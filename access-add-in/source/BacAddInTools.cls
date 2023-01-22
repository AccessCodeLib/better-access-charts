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

Private m_ScriptPropertiesLoaded As Boolean

Public Property Get BAC() As BAC__Factory

    If Not m_ScriptPropertiesLoaded Then
        DefineScripts DefaultScriptProperties
        m_ScriptPropertiesLoaded = True
    End If
    
    Set BAC = BetterAccessCharts
    
End Property

Private Sub DefineScripts(ByVal ScriptProperties As BAC_Properties)

    ' ## Scripts
    'Define chart.js
    With ScriptProperties.Item("ChartJS").Properties
        BetterAccessCharts.Scripts.ChartJS.Source = .Item("Source").Value
        BetterAccessCharts.Scripts.ChartJS.CDN = .Item("CDN").Value
        BetterAccessCharts.Scripts.ChartJS.PathFilename = .Item("PathFilename").Value
    End With
    
    'Define color scheme plug-in
    With ScriptProperties.Item("Plugin_ColorSchemes").Properties
        BetterAccessCharts.Scripts.Plugin_ColorSchemes.Source = .Item("Source").Value
        BetterAccessCharts.Scripts.Plugin_ColorSchemes.CDN = .Item("CDN").Value
        BetterAccessCharts.Scripts.Plugin_ColorSchemes.PathFilename = .Item("PathFilename").Value
    End With
    
    'Define data label plug in
    With ScriptProperties.Item("Plugin_DataLabels").Properties
        BetterAccessCharts.Scripts.Plugin_DataLabels.Source = .Item("Source").Value
        BetterAccessCharts.Scripts.Plugin_DataLabels.CDN = .Item("CDN").Value
        BetterAccessCharts.Scripts.Plugin_DataLabels.PathFilename = .Item("PathFilename").Value
    End With
    
End Sub

Private Property Get DefaultScriptProperties() As BAC_Properties
    
    Const ScriptPropertiesName As String = "DefaultScriptProperties"

    Dim ScriptProperties As BAC_Properties
    
    Set ScriptProperties = Me.ChartConfiguration.GetFromTable(ScriptPropertiesName)
    
    If ScriptProperties Is Nothing Then
        With New Form_frm_Configuration
            Set ScriptProperties = .ScriptProperties
        End With
    End If
    
    Set DefaultScriptProperties = ScriptProperties
    
End Property

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
