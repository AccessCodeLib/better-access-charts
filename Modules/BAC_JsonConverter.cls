VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_JsonConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const SubPropIndentLen As Long = 2

Public Function Json2Properties(ByVal JsonString As String) As BAC_Properties

    JsonString = OptimizeJsonString(JsonString)
    
    If Left(JsonString, 1) <> "{" Then
        Err.Raise vbObjectError, "JsonConverter.Parse", "Missing JSON structure"
    End If
    
    If Right(JsonString, 1) <> "}" Then
        Err.Raise vbObjectError, "JsonConverter.Parse", "Missing JSON structure"
    End If
    
    Set Json2Properties = ParseProperties(JsonString)
    
End Function

Public Function Properties2Json(ByVal Properties As BAC_Properties) As String
    Properties2Json = ConvertProperties(Properties)
End Function

'#############################################
' Parse JSON

Private Function OptimizeJsonString(ByVal JsonString As String) As String

    JsonString = Replace(JsonString, vbNewLine, vbNullString)
    JsonString = Replace(JsonString, vbTab, vbNullString)
    JsonString = Replace(JsonString, vbCr, vbNullString)
    JsonString = Replace(JsonString, vbLf, vbNullString)
    
    JsonString = Trim(JsonString)
    
    OptimizeJsonString = JsonString

End Function

Private Function ParseProperties(ByVal JsonString As String) As BAC_Properties

    Dim Props As BAC_Properties
    Set Props = New BAC_Properties
    
    Dim PropStringArray() As String
    Dim i As Long
    
    JsonString = Trim(JsonString)
    If Left(JsonString, 1) = "{" And Right(JsonString, 1) = "}" Then ' if
        JsonString = Mid(JsonString, 2, Len(JsonString) - 2)
    Else
        Err.Raise vbObjectError, "JsonConverter.ParseProperties", "Missing JSON structure"
    End If
    
    
    PropStringArray = Split(SetSplitMarker(JsonString), ",|")
    
    For i = LBound(PropStringArray) To UBound(PropStringArray)
        Props.Append ConvertString2Property(PropStringArray(i))
    Next
    
    Set ParseProperties = Props

End Function

Private Function ConvertString2Property(ByVal JsonString As String) As BAC_Property
'   "Name" : "abc" => Value
'   "Name" : 123   => Value
'   "Name" : { ... => Properties

    Dim Prop As BAC_Property
    
    Dim PropName As String
    Dim PropValueString As Variant
    
    Dim CutPos As Long
    
    Set Prop = New BAC_Property
    
    CutPos = InStr(1, JsonString, ":")
    
    PropName = Trim(Left(JsonString, CutPos - 1))
    If Left(PropName, 1) = """" And Right(PropName, 1) = """" Then
        PropName = Mid(PropName, 2, Len(PropName) - 2)
    End If
    Prop.Name = PropName
    
    PropValueString = Trim(Mid(JsonString, CutPos + 1))
    If Left(PropValueString, 1) = "{" Then ' new properties
        Set Prop.Properties = ParseProperties(PropValueString)
    ElseIf PropValueString = "null" Then
        Prop.Value = Null
    ElseIf PropValueString = "true" Then
        Prop.Value = True
    ElseIf PropValueString = "false" Then
        Prop.Value = False
    ElseIf Left(PropValueString, 1) = """" Then 'string
        Prop.Value = Mid(PropValueString, 2, Len(PropValueString) - 2)
    Else ' numeric (other data types not used in BAC)
        Prop.Value = Val(PropValueString)
    End If
    
    Set ConvertString2Property = Prop
    
End Function

Private Function SetSplitMarker(ByVal JsonString As String) As String

    Dim Pos As Long
    Dim LeftPart As String
    
    JsonString = Replace(JsonString, ",", ", ") ' place to insert marker

    Pos = InStr(1, JsonString, ",")
    
    Do While Pos > 0
        If Not PosIsInSubObject(Left(JsonString, Pos - 1)) Then
            Mid(JsonString, Pos + 1, 1) = "|"
        End If
        Pos = InStr(Pos + 1, JsonString, ",")
        Debug.Print JsonString, Pos
    Loop
    
    SetSplitMarker = Replace(JsonString, ", ", ",")

End Function

Private Function PosIsInSubObject(ByVal StringToCheck As String) As Boolean

    Dim CntOpeningBrackets As Long
    Dim CntClosingBrackets As Long
    
'Debug.Print StringToCheck,
    
    CntOpeningBrackets = Len(StringToCheck) - Len(Replace(StringToCheck, "{", vbNullString))
    CntClosingBrackets = Len(StringToCheck) - Len(Replace(StringToCheck, "}", vbNullString))
    
    If CntOpeningBrackets > CntClosingBrackets Then
        PosIsInSubObject = True
    Else
        PosIsInSubObject = False
    End If
    
'Debug.Print PosIsInSubObject

End Function


'##################
' Build JSON
'
Private Function ConvertProperties(ByVal Properties As BAC_Properties, Optional ByVal PropLevel As Long = 0) As String

    Dim Prop As Property
    Dim i As Long
    Dim PropCnt As Long
    
    Dim sb As BAC_StringBuilder
    Set sb = New BAC_StringBuilder
    
    PropCnt = Properties.Count
    
    sb.AddLine "{"
    
    For i = 1 To PropCnt
        sb.Add ConvertProperty(Properties.Item(i), PropLevel + 1)
        If i < PropCnt Then
            sb.Add "," & vbNewLine
        End If
    Next
    
    'sb.AddLine vbNullString
    sb.Add vbNewLine & String(PropLevel * SubPropIndentLen, " ") & "}"
    
    ConvertProperties = sb.ToString

End Function

Private Function ConvertProperty(ByVal Property As BAC_Property, Optional ByVal PropLevel As Long = 0) As String

    With New BAC_StringBuilder
        .Add String(PropLevel * SubPropIndentLen, " ") & """" & Property.Name & """ : "
        
        If Property.Properties.Count > 0 Then
            .Add ConvertProperties(Property.Properties, PropLevel)
        Else
            .Add ConvertValue(Property.Value)
        End If
        
        ConvertProperty = .ToString
        
    End With
    
End Function

Private Function ConvertValue(ByVal Value As Variant) As String

    Dim ValueString As String

    Select Case VarType(Value)
        Case VbVarType.vbBoolean
            If Value = True Then
                ValueString = "true"
            Else
                ValueString = "false"
            End If
        Case VbVarType.vbString
            ValueString = """" & Value & """"
        Case VbVarType.vbNull
            ValueString = "null"
        Case Else
            ValueString = Trim(str(Value))
    End Select
    
    ConvertValue = ValueString

End Function
