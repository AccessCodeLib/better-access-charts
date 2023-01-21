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

Public Sub DeleteFormProperty(ByVal FormName As String, ByVal ChartName As String)
    
    Dim ao As AccessObject
    Dim Prop As AccessObjectProperty
    
    Set ao = CurrentProject.AllForms(FormName)
    
    For Each Prop In ao.Properties
        If Prop.Name = ChartName Then
            ao.Properties.Remove Prop.Name
            Exit For
        End If
    Next
    
End Sub

Public Function GetFormProperty(ByVal FormName As String, ByVal ChartName As String) As BAC_Properties

    Dim ao As AccessObject
    Dim Prop As AccessObjectProperty
    Dim PropJson As String
    Dim ChartProperties As BAC_Properties
    
    Set ao = CurrentProject.AllForms(FormName)
    
    For Each Prop In ao.Properties
        If Prop.Name = ChartName Then
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

Public Sub SaveFormProperty(ByVal FormName As String, ByVal ChartName As String, ByVal ChartProperites As BAC_Properties)

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
        If Prop.Name = ChartName Then
            PropExists = True
            Exit For
        End If
    Next
    
    If PropExists Then
        ao.Properties(ChartName).Value = PropJson
    Else
        ao.Properties.Add ChartName, PropJson
    End If

End Sub
