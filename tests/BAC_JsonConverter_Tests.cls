VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_JsonConverter_Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/tests/BAC_JsonConverter_Tests.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
'
'@AccUnit
' Test environment: https://github.com/AccessCodeLib/AccUnit/tree/main/access-add-in
' start tests with: TestSuite.Add(new BAC_JsonConverter_Tests).Run
'
Option Compare Database
Option Explicit

Const SimpleJsonTestString As String = _
"{" & vbNewLine & _
"  ""NumProp"" : 123," & vbNewLine & _
"  ""BoolProp"" : true," & vbNewLine & _
"  ""StrProp"" : ""abc""" & vbNewLine & _
"}"

Const JsonStringWithSubObject As String = _
"{" & vbNewLine & _
"  ""NumProp1"" : 123," & vbNewLine & _
"  ""StrProp1"" : ""abc""," & vbNewLine & _
"  ""Properties1"" : {" & vbNewLine & _
"    ""NumProp"" : 456," & vbNewLine & _
"    ""StrProp"" : ""xyz""" & vbNewLine & _
"  }" & vbNewLine & _
"}"


'###############
' TESTS
'

Public Sub ParseSimpleJsonString()

    Dim P As BAC_Properties
    
    With New BAC_JsonConverter
        Set P = .Json2Properties(SimpleJsonTestString)
    End With
    
    Assert.That P.Item("NumProp").Value, Iz.EqualTo(123)
    Assert.That P.Item("BoolProp").Value, Iz.EqualTo(True)
    Assert.That P.Item("StrProp").Value, Iz.EqualTo("abc")
    
End Sub

Public Sub ParseJsonStringWithSubProperties()

    Dim P As BAC_Properties
    
    With New BAC_JsonConverter
        Set P = .Json2Properties(JsonStringWithSubObject)
    End With
    
    Assert.That P.Item("NumProp1").Value, Iz.EqualTo(123)
    Assert.That P.Item("StrProp1").Value, Iz.EqualTo("abc")
    Assert.That P.Item("Properties1").Properties.Item("NumProp").Value, Iz.EqualTo(456)
    Assert.That P.Item("Properties1").Properties.Item("StrProp").Value, Iz.EqualTo("xyz")
    
End Sub

Public Sub TestCreateSimpleJsonString()

    Dim P As BAC_Properties
    Dim JsonString As String
    
    With New BAC_JsonConverter
        Set P = .Json2Properties(SimpleJsonTestString)
        JsonString = .Properties2Json(P)
    End With
    
    Debug.Print "----"
    Debug.Print SimpleJsonTestString
    Debug.Print "----"
    Debug.Print JsonString
    Debug.Print "----"
    
    Assert.That JsonString, Iz.EqualTo(SimpleJsonTestString)

End Sub

Public Sub TestCreateJsonStringWithSubObject()

    Dim P As BAC_Properties
    Dim JsonString As String
    
    With New BAC_JsonConverter
        Set P = .Json2Properties(JsonStringWithSubObject)
        JsonString = .Properties2Json(P)
    End With
    
    Assert.That JsonString, Iz.EqualTo(JsonStringWithSubObject)

End Sub
