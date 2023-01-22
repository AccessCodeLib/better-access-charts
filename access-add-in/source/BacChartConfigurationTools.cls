VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BacChartConfigurationTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/BacChartConfigurationTools.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const DefaultTableName As String = "USys_Properties"
Private Const DefaultPropNameDataField As String = "PropName"
Private Const DefaultJsonDataField As String = "JSON"
Private Const DefaultUseAddInDatabase As Boolean = False

Private m_TableName As String
Private m_PropNameDataField As String
Private m_JsonDataField As String
Private m_UseAddInDatabase As Boolean

Private Sub Class_Initialize()
    InitTable DefaultTableName, DefaultPropNameDataField, DefaultJsonDataField, DefaultUseAddInDatabase
End Sub

Public Sub InitTable( _
            Optional ByVal TableName As String = DefaultTableName, _
            Optional ByVal PropNameDataField As String = DefaultPropNameDataField, _
            Optional ByVal JsonDataField As String = DefaultJsonDataField, _
            Optional ByVal UseAddInDatabase As Boolean = False)
            
    m_TableName = TableName
    m_PropNameDataField = PropNameDataField
    m_JsonDataField = JsonDataField
    m_UseAddInDatabase = UseAddInDatabase
    
End Sub

Public Sub DeleteFormProperty(ByVal FormName As String, ByVal PropertyName As String)
    
    Dim ao As AccessObject
    Dim Prop As AccessObjectProperty
    
    Set ao = CurrentProject.AllForms(FormName)
    
    For Each Prop In ao.Properties
        If Prop.Name = PropertyName Then
            ao.Properties.Remove Prop.Name
            Exit For
        End If
    Next
    
End Sub

Public Function GetFormProperty(ByVal FormName As String, ByVal PropertyName As String) As BAC_Properties

    Dim ao As AccessObject
    Dim Prop As AccessObjectProperty
    Dim PropJson As String
    Dim ChartProperties As BAC_Properties
    
    Set ao = CurrentProject.AllForms(FormName)
    
    For Each Prop In ao.Properties
        If Prop.Name = PropertyName Then
            PropJson = Trim(Nz(Prop.Value, vbNullString))
            Exit For
        End If
    Next
    
    If Len(PropJson) = 0 Then
        Set GetFormProperty = Nothing
        Exit Function
    End If
    
    With New BAC_JsonConverter
        Set ChartProperties = .Json2Properties(PropJson)
    End With
    
    Set GetFormProperty = ChartProperties

End Function

Public Sub SaveToFormProperty(ByVal FormName As String, ByVal PropertyName As String, ByVal ChartProperites As BAC_Properties)

    'AccessObjectProperties
    Dim ao As AccessObject
    Dim Prop As AccessObjectProperty
    Dim PropExists As Boolean
    Dim PropJson As String
    
    With New BAC_JsonConverter
        PropJson = .Properties2Json(ChartProperites)
    End With
    
    Set ao = CurrentProject.AllForms(FormName)
    
    For Each Prop In ao.Properties
        If Prop.Name = PropertyName Then
            PropExists = True
            Exit For
        End If
    Next
    
    If PropExists Then
        Prop.Value = PropJson
    Else
        ao.Properties.Add PropertyName, PropJson
    End If

End Sub

Public Function GetFromTable(ByVal PropName As String) As BAC_Properties

    Dim JsonString As String
    Dim Properties As BAC_Properties
    
    With InstanceDb.OpenRecordset("select " & m_JsonDataField & " from " & m_TableName & " where " & m_PropNameDataField & " = '" & PropName & "'", dbOpenSnapshot)
        If Not .EOF Then
            JsonString = Nz(.Fields(0).Value, vbNullString)
        End If
        .Close
    End With
    
    If Len(JsonString) = 0 Then
        Set Properties = Nothing
    Else
        With New BAC_JsonConverter
            Set Properties = .Json2Properties(JsonString)
        End With
    End If
    
    Set GetFromTable = Properties

End Function

Public Sub SaveToTable(ByVal PropName As String, ByVal Properties As BAC_Properties)

    Dim JsonString As String
    
    With New BAC_JsonConverter
        JsonString = .Properties2Json(Properties)
    End With
    
    With InstanceDb.OpenRecordset("select " & m_PropNameDataField & ", " & m_JsonDataField & " from " & m_TableName & " where " & m_PropNameDataField & " = '" & PropName & "'", dbOpenDynaset)
        If .EOF Then
            .AddNew
            .Fields("PropName").Value = PropName
        Else
            .Edit
        End If
        .Fields("JSON").Value = JsonString
        .Update
        .Close
    End With
    
End Sub

Private Property Get InstanceDb() As DAO.Database

    If m_UseAddInDatabase Then
        Set InstanceDb = CodeDb
    Else
        Set InstanceDb = CurrentDb
    End If

End Property
