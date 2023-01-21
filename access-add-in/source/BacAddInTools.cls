VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BacAddInTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/BacAddInTools.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public Function NewCodeModuleSupport() As CodeModuleSupport
    
    Set NewCodeModuleSupport = New CodeModuleSupport
    
End Function

Public Function NewConfigurationForm(ByRef ReturnProperties As BAC_Properties, Optional ByVal ChartProperties As BAC_Properties = Nothing) As Form

    Dim ConfigFrm As Form
    Set ConfigFrm = New Form_frm_Configuration
    
    If ReturnProperties Is Nothing Then
        Set ReturnProperties = New BAC_Properties
        ReturnProperties.NewProperty "Commit"
        ReturnProperties.NewProperty "ChartProperties"
        ReturnProperties.NewProperty "ScriptProperties"
    End If
    
    ConfigFrm.Init ReturnProperties

    If Not (ChartProperties Is Nothing) Then
        ConfigFrm.LoadConfig ChartProperties
    End If
    
    ConfigFrm.Visible = True
    
    Set NewConfigurationForm = ConfigFrm
        
End Function

Public Property Get ChartConfiguration() As BacChartConfigurationTools
    Set ChartConfiguration = New BacChartConfigurationTools
End Property
