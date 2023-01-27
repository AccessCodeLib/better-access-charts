VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModuleSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/CodeModuleSupport.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const ErrNumber91_ObjectVariableOrWithBlockVariableNotSet As Long = 91

Private Const BAC_EarlyBindingConstName As String = "BAC_EarlyBinding"
Private Const ConditionalCompilationArgumentsOptionName As String = "Conditional Compilation Arguments"
Private Const BetterAccessChartsReferenceName As String = "BetterAccessChartsLib"

Private Const FormCodeMarker_Start As String = "'### BetterAccessCharts support - START"
Private Const FormCodeMarker_End As String = "'### BetterAccessCharts support - END"

#Const EARLYBINDING = 0
#Const DEBUGMODE = 0

#If EARLYBINDING = 1 Then
#Else

'   Public Enum vbext_ComponentType
'      vbext_ct_StdModule = 1
'      vbext_ct_ClassModule = 2
'      vbext_ct_MSForm = 3
'      vbext_ct_ActiveXDesigner = 11
'      vbext_ct_Document = 100
'   End Enum
   
#End If

#If EARLYBINDING = 1 Then
Private Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Private Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING = 1 Then
   Dim Proj As VBProject
   Dim Temp As VBProject
#Else
   Dim Proj As Object
   Dim Temp As Object
#End If
   Dim strCurrentDbName As String

   Set Proj = VBE.ActiveVBProject
   'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
   strCurrentDbName = UncPath(CurrentDb.Name)
   If Proj.FileName <> strCurrentDbName Then
      Set Proj = Nothing
      For Each Temp In VBE.VBProjects
         If Temp.FileName = strCurrentDbName Then
            Set Proj = Temp
            Exit For
         End If
      Next
   End If
   
   Set CurrentVbProject = Proj

End Property

Public Property Get EARLYBINDING() As Boolean

   Dim ConditionalCompilationArgumentsString As String
   Dim ConditionalCompilationArguments() As String
   Dim TestString As String
   Dim i As Long
   
   ConditionalCompilationArgumentsString = Trim(Nz(Application.GetOption(ConditionalCompilationArgumentsOptionName), vbNullString))
   
   If Len(ConditionalCompilationArgumentsString) = 0 Then
      EARLYBINDING = False
      Exit Property
   End If
   
   ConditionalCompilationArguments = GetConditionalCompilationArgumentsArray(ConditionalCompilationArgumentsString)
   For i = LBound(ConditionalCompilationArguments) To UBound(ConditionalCompilationArguments)
      TestString = Replace(ConditionalCompilationArguments(i), " ", vbNullString)
      If TestString Like BAC_EarlyBindingConstName & "=*" Then
         EARLYBINDING = (TestString <> BAC_EarlyBindingConstName & "=0")
         Exit For
      End If
   Next
   
End Property

Public Property Let EARLYBINDING(ByVal UseEarlyBinding As Boolean)

   Dim ConditionalCompilationArgumentsString As String
   Dim ConditionalCompilationArguments() As String
   Dim TestString As String
   Dim ArgExists As Boolean
   Dim i As Long
   
   ConditionalCompilationArgumentsString = Trim(Nz(Application.GetOption("Conditional Compilation Arguments"), vbNullString))
   
   If Len(ConditionalCompilationArgumentsString) = 0 Then
      Application.SetOption ConditionalCompilationArgumentsOptionName, BAC_EarlyBindingConstName & " = " & Abs(UseEarlyBinding)
      Exit Property
   End If
   
   ConditionalCompilationArguments = GetConditionalCompilationArgumentsArray(ConditionalCompilationArgumentsString)
   For i = LBound(ConditionalCompilationArguments) To UBound(ConditionalCompilationArguments)
      If Replace(ConditionalCompilationArguments(i), " ", vbNullString) Like BAC_EarlyBindingConstName & "=*" Then
         ConditionalCompilationArguments(i) = BAC_EarlyBindingConstName & " = " & Abs(UseEarlyBinding)
         ArgExists = True
         Exit For
      End If
   Next
   
   If ArgExists Then
      ConditionalCompilationArgumentsString = Join(ConditionalCompilationArguments, ":")
   Else
      ConditionalCompilationArgumentsString = ConditionalCompilationArgumentsString & " : " & BAC_EarlyBindingConstName & " = " & Abs(UseEarlyBinding)
   End If

   Application.SetOption ConditionalCompilationArgumentsOptionName, ConditionalCompilationArgumentsString
   
   If UseEarlyBinding Then
      AddReference
   Else
      RemoveReference
   End If

End Property

Private Function GetConditionalCompilationArgumentsArray(ByVal FullString As String) As String()
   GetConditionalCompilationArgumentsArray = Split(FullString, ":")
End Function

Private Sub AddReference()
   If Not ReferenceExits() Then
    Application.References.AddFromFile CodeProject.FullName
   End If
End Sub

Private Function GetAddInLocation() As String
   Dim strLocation As String
   strLocation = CodeProject.Path
   GetAddInLocation = strLocation
End Function

Private Function ReferenceExits() As Boolean
   Dim ref As Reference
   Dim BacRef As Reference
   For Each ref In Application.References
      If ref.Name = BetterAccessChartsReferenceName Then
         ReferenceExits = True
         Exit Function
      End If
   Next
   ReferenceExits = False
End Function

Private Sub RemoveReference()
   Dim ref As Reference
   Dim BacRef As Reference
   For Each ref In Application.References
      If ref.Name = BetterAccessChartsReferenceName Then
         Set BacRef = ref
      End If
   Next
   If Not (BacRef Is Nothing) Then
      Application.References.Remove BacRef
   End If
End Sub


Public Sub CreateModulesInApp()
   CreateModuleInApp BAC_LoaderModuleName, GetBetterAccessChartsLoaderModuleString
   CreateModuleInApp BAC_FactoryModuleName, GetBetterAccessChartsFactoryModuleString
End Sub

Public Sub CreateModuleInApp(ByVal ModuleName As String, ByVal CodeString As String)

On Error GoTo Fehler

#If EARLYBINDING = 1 Then
    Dim cm As CodeModule
    Dim vbc As VBComponent
    Dim vbp As VBProject
#Else
    Dim cm As Object
    Dim vbc As Object
    Dim vbp As Object
#End If
    
    If CodeModuleExists(ModuleName) Then
        Set vbc = CurrentVbProject.VBComponents(ModuleName)
    Else 'Modul anlegen und benennen
        Set vbc = CurrentVbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
        vbc.Name = ModuleName
    End If

   'Code einfügen
    Set cm = vbc.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString CodeString
    
    'Leerzeilen am Ende säubern
    Do While Len(Trim$(cm.Lines(cm.CountOfLines, 1))) = 0 And cm.CountOfLines > 1
        cm.DeleteLines cm.CountOfLines
    Loop
    
    Call SaveModule
    
Ende_Cleanup:
    On Error Resume Next
    Set cm = Nothing
    Set vbc = Nothing
    Set vbp = Nothing
    Exit Sub

Fehler:
    If Err.Number = ErrNumber91_ObjectVariableOrWithBlockVariableNotSet Then
        RaiseError ErrNumber91_ObjectVariableOrWithBlockVariableNotSet, "TransferCodeModule.CreateModuleInApp", Err.Source, "Keine Datenbank geöffnet!"
    Else
        RaiseError Err.Number, "TransferCodeModule.CreateModuleInApp", Err.Source, Err.Description
    End If
    Resume Ende_Cleanup
    
End Sub

Private Sub SaveModule()
    DoCmd.Save acModule, BAC_FactoryModuleName
    DoCmd.Close acModule, BAC_FactoryModuleName
End Sub

Public Sub DeleteModuleFromApp()

On Error GoTo Fehler

#If EARLYBINDING = 1 Then
    Dim vbc As VBComponent
#Else
    Dim vbc As Object
#End If
    
    Set vbc = CurrentVbProject.VBComponents(BAC_FactoryModuleName)
    CurrentVbProject.VBComponents.Remove vbc
    
Ende_Cleanup:
    On Error Resume Next
    Set vbc = Nothing
    Exit Sub

Fehler:
    If Err.Number = ErrNumber91_ObjectVariableOrWithBlockVariableNotSet Then
        RaiseError ErrNumber91_ObjectVariableOrWithBlockVariableNotSet, "TransferCodeModule.DeleteModuleFromApp", Err.Source, "Keine Datenbank geöffnet!"
    Else
        RaiseError Err.Number, "TransferCodeModule.DeleteModuleFromApp", Err.Source, Err.Description
    End If
    Resume Ende_Cleanup

End Sub

Public Function CodeModuleExists( _
                ByVal CodeModulName As String) As Boolean

#If EARLYBINDING = 1 Then
   Dim VbcCol As VBIDE.VBComponents
   Dim vbc As VBIDE.VBComponent
#Else
   Dim VbcCol As Object
   Dim vbc As Object
#End If
   
   Set VbcCol = CurrentVbProject.VBComponents
   For Each vbc In VbcCol
      If vbc.Name = CodeModulName Then
         CodeModuleExists = True
         Exit For
      End If
   Next
   
End Function

Private Sub RaiseError( _
                    ByVal ErrNumber As Long, _
                    ByVal ErrSource As String, _
                    ByVal ErrSubSource As String, _
                    ByVal ErrDescription As String)

    If ErrSubSource <> "sqlDebugPrint" Then
        ErrSource = ErrSource & "->" & ErrSubSource
    End If
    Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub
                    
Private Function GetBetterAccessChartsFactoryModuleString() As String

    Dim str As String
    
    str = "'---------------------------------------------------------------------------------------" & vbNewLine & _
"'<codelib>" & vbNewLine & _
"'  <file>%AppFolder%/source/BetterAccessChartsFactory.bas</file>" & vbNewLine & _
"'</codelib>" & vbNewLine & _
"'---------------------------------------------------------------------------------------" & vbNewLine & _
"Option Compare Database" & vbNewLine & _
"Option Explicit" & vbNewLine & _
"" & vbNewLine & _
"Private m_BAC As Object" & vbNewLine & _
"Private m_BACx As Object" & vbNewLine & _
"" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"#Else" & vbNewLine
         
   str = str & vbNewLine & _
"Public Enum chChartType" & vbNewLine & _
"    Line = 1" & vbNewLine & _
"    Bar = 2" & vbNewLine & _
"    HorizontalBar = 3" & vbNewLine & _
"    Radar = 4" & vbNewLine & _
"    Doughnut = 5" & vbNewLine & _
"    Pie = 6" & vbNewLine & _
"    PolarArea = 7" & vbNewLine & _
"    Bubble = 8" & vbNewLine & _
"    Scatter = 9" & vbNewLine & _
"End Enum" & vbNewLine & _
"" & vbNewLine & _
"Public Enum chDataSourceType" & vbNewLine & _
"    dstDemo" & vbNewLine & _
"    dstTableName" & vbNewLine & _
"    dstQueryName" & vbNewLine & _
"    dstSQLStament" & vbNewLine & _
"    dstRecordset" & vbNewLine & _
"    dstData" & vbNewLine & _
"    dstEmpty" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chPosition" & vbNewLine & _
"    posTop = 1" & vbNewLine & _
"    posLeft = 2" & vbNewLine & _
"    posBottom = 3" & vbNewLine & _
"    posRight = 4" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chScriptSource" & vbNewLine & _
"    CDN = 1" & vbNewLine & _
"    LocalFile = 2" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chAlign" & vbNewLine & _
"    alStart = 1" & vbNewLine & _
"    alCenter = 2" & vbNewLine & _
"    alEnd = 3" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chDataLabelAnchor" & vbNewLine & _
"    anStart = 1" & vbNewLine & _
"    anCenter = 2" & vbNewLine & _
"    anEnd = 3" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chDisplayIn" & vbNewLine & _
"    chWebBrowserControl = 1" & vbNewLine & _
"    chWebBrowserActiveX = 2" & vbNewLine & _
"    chImageControl = 3" & vbNewLine & _
"    chSystemBrowser = 4" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chEasing" & vbNewLine & _
"    linear = 0" & vbNewLine & _
"    easeInQuad = 1" & vbNewLine & _
"    easeOutQuad = 2" & vbNewLine & _
"    easeInOutQuad = 3" & vbNewLine & _
"    easeInCubic = 4" & vbNewLine & _
"    easeOutCubic = 5" & vbNewLine & _
"    easeInOutCubic = 6" & vbNewLine & _
"    easeInQuart = 7" & vbNewLine & _
"    easeOutQuart = 8" & vbNewLine & _
"    easeInOutQuart = 9" & vbNewLine & _
"    easeInQuint = 10" & vbNewLine & _
"    easeOutQuint = 11" & vbNewLine & _
"    easeInOutQuint = 12" & vbNewLine & _
"    easeInSine = 13" & vbNewLine & _
"    easeOutSine = 14" & vbNewLine & _
"    easeInOutSine = 15" & vbNewLine & _
"    easeInExpo = 16" & vbNewLine

   str = str & _
"    easeOutExpo = 17" & vbNewLine & _
"    easeInOutExpo = 18" & vbNewLine & _
"    easeInCirc = 19" & vbNewLine & _
"    easeOutCirc = 20" & vbNewLine & _
"    easeInOutCirc = 21" & vbNewLine & _
"    easeInElastic = 22" & vbNewLine & _
"    easeOutElastic = 23" & vbNewLine & _
"    easeInOutElastic = 24" & vbNewLine & _
"    easeInBack = 25" & vbNewLine & _
"    easeOutBack = 26" & vbNewLine & _
"    easeInOutBack = 27" & vbNewLine & _
"    easeInBounce = 28" & vbNewLine & _
"    easeOutBounce = 29" & vbNewLine & _
"    easeInOutBounce = 30" & vbNewLine & _
"End Enum" & vbNewLine
   
   str = str & vbNewLine & _
"#End If"

   str = str & vbNewLine & _
"' Factory" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"Public Function BAC() As " & BetterAccessChartsReferenceName & ".BAC__Factory" & vbNewLine & _
"#Else" & vbNewLine & _
"Public Function BAC() As Object" & vbNewLine & _
"#End If" & vbNewLine & _
"    If m_Bac Is Nothing Then" & vbNewLine & _
"        Set m_BAC = BACx.BAC" & vbNewLine & _
"    End If" & vbNewLine & _
"    Set BAC = m_Bac" & vbNewLine & _
"End Function" & vbNewLine
       
       
   str = str & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"Public Function BACx() As " & BetterAccessChartsReferenceName & ".BacAddInTools" & vbNewLine & _
"#Else" & vbNewLine & _
"Public Function BACx() As Object" & vbNewLine & _
"#End If" & vbNewLine & _
"    If m_BACx Is Nothing Then" & vbNewLine & _
"        Set m_BACx = BetterAccessChartsLoader.GetBetterAccessChartsAddInTools" & vbNewLine & _
"    End If" & vbNewLine & _
"    Set BACx = m_BACx" & vbNewLine & _
"End Function" & vbNewLine

    GetBetterAccessChartsFactoryModuleString = str

End Function

Private Function GetBetterAccessChartsLoaderModuleString() As String

   Dim str As String
   
    str = "'---------------------------------------------------------------------------------------" & vbNewLine & _
"'<codelib>" & vbNewLine & _
"'  <file>%AppFolder%/source/BetterAccessChartsLoader.bas</file>" & vbNewLine & _
"'</codelib>" & vbNewLine & _
"'---------------------------------------------------------------------------------------" & vbNewLine & _
"Option Compare Database" & vbNewLine & _
"Option Explicit" & vbNewLine & _
"" & vbNewLine & _
"' Set BAC_EarlyBinding in project property for compiler arguments" & vbNewLine & _
"' BAC_EarlyBinding = 1 ... CheckBetterAccessChartsReference add reference to accda (BAC created by accda-Reference)" & vbNewLine & _
"' BAC_EarlyBinding = 0 ... CheckBetterAccessChartsReference removes reference (BAC created by add-in)" & vbNewLine & _
"" & vbNewLine & _
"Private Const BetterAccessChartsFileName As String = ""BetterAccessCharts""" & vbNewLine & _
"Private Const BetterAccessChartsReferenceName As String = ""BetterAccessChartsLib""" & vbNewLine & _
"Private Const BetterAccessChartsAddInToolsName As String = ""BetterAccessChartsAddInTools""" & vbNewLine

   str = str & vbNewLine & _
"Public Sub CheckBetterAccessChartsReference()" & vbNewLine & _
"   CheckReference" & vbNewLine & _
"End Sub" & vbNewLine

   str = str & vbNewLine & _
"Public Function GetBetterAccessChartsAddInTools() As Object" & vbNewLine & _
"   CheckReference" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"    Set GetBetterAccessChartsAddInTools = " & BetterAccessChartsReferenceName & ".BetterAccessChartsAddInTools" & vbNewLine & _
"#Else" & vbNewLine & _
"    Set GetBetterAccessChartsAddInTools = Application.Run(GetAddInLocation & BetterAccessChartsFileName & ""."" & BetterAccessChartsAddInToolsName)" & vbNewLine & _
"#End If" & vbNewLine & _
"End Function" & vbNewLine

   str = str & vbNewLine & _
"Private Function GetAddInLocation() As String" & vbNewLine & _
"   Dim strLocation As String" & vbNewLine & _
"   strLocation = GetAppDataLocation & ""\Microsoft\AddIns\""" & vbNewLine & _
"   If Len(VBA.Dir(strLocation & BetterAccessChartsFileName & "".accda"")) = 0 Then" & vbNewLine & _
"      strLocation = CodeProject.Path & ""\""" & vbNewLine & _
"      If Len(VBA.Dir(strLocation & BetterAccessChartsFileName & "".accda"")) = 0 Then" & vbNewLine & _
"         Err.Raise vbObjectError, ""BetterAccessChartsLoader.GetAddInLocation"", ""Add-In file is missing""" & vbNewLine & _
"      End If" & vbNewLine & _
"   End If" & vbNewLine & _
"   GetAddInLocation = strLocation" & vbNewLine & _
"End Function" & vbNewLine

   str = str & vbNewLine & _
"Private Function GetAppDataLocation()" & vbNewLine & _
"   With CreateObject(""WScript.Shell"")" & vbNewLine & _
"      GetAppDataLocation = .ExpandEnvironmentStrings(""%APPDATA%"") & """"" & vbNewLine & _
"   End With" & vbNewLine & _
"End Function" & vbNewLine

   str = str & vbNewLine & _
"Private Sub CheckReference()" & vbNewLine & _
"   Static m_ReferenceChecked As Boolean" & vbNewLine & _
"   Static m_UseEarlyBindingState As Boolean" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"   If m_UseEarlyBindingState = False Then" & vbNewLine & _
"      m_ReferenceChecked = False" & vbNewLine & _
"      m_UseEarlyBindingState = True" & vbNewLine & _
"   End If" & vbNewLine & _
"#Else" & vbNewLine & _
"   If m_UseEarlyBindingState = True Then" & vbNewLine & _
"      m_ReferenceChecked = False" & vbNewLine & _
"      m_UseEarlyBindingState = False" & vbNewLine & _
"   End If" & vbNewLine & _
"#End If" & vbNewLine & _
"   If m_ReferenceChecked Then" & vbNewLine & _
"      Exit Sub" & vbNewLine & _
"   End If" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"    AddReference" & vbNewLine & _
"#Else" & vbNewLine & _
"    RemoveReference" & vbNewLine & _
"#End If" & vbNewLine & _
"End Sub" & vbNewLine

   str = str & vbNewLine & _
"Private Sub AddReference()" & vbNewLine & _
"   If Not ReferenceExits() Then" & vbNewLine & _
"    Application.References.AddFromFile GetAddInLocation & BetterAccessChartsFileName & "".accda""" & vbNewLine & _
"   End If" & vbNewLine & _
"End Sub" & vbNewLine

   str = str & vbNewLine & _
"Private Function ReferenceExits() As Boolean" & vbNewLine & _
"   Dim ref As Reference" & vbNewLine & _
"   Dim BacRef As Reference" & vbNewLine & _
"   For Each ref In Application.References" & vbNewLine & _
"      If ref.Name = BetterAccessChartsReferenceName Then" & vbNewLine & _
"         ReferenceExits = True" & vbNewLine & _
"         Exit Function" & vbNewLine & _
"      End If" & vbNewLine & _
"   Next" & vbNewLine & _
"   ReferenceExits = False" & vbNewLine & _
"End Function" & vbNewLine

   str = str & vbNewLine & _
"Private Sub RemoveReference()" & vbNewLine & _
"   Dim ref As Reference" & vbNewLine & _
"   Dim BacRef As Reference" & vbNewLine & _
"   For Each ref In Application.References" & vbNewLine & _
"      If ref.Name = BetterAccessChartsReferenceName Then" & vbNewLine & _
"         Set BacRef = ref" & vbNewLine & _
"      End If" & vbNewLine & _
"   Next" & vbNewLine & _
"   If Not (BacRef Is Nothing) Then" & vbNewLine & _
"      Application.References.Remove BacRef" & vbNewLine & _
"   End If" & vbNewLine & _
"End Sub" & vbNewLine

   GetBetterAccessChartsLoaderModuleString = str

End Function


Public Sub InsertFormCode( _
                ByVal ChartFormName As String, _
                ByVal ChartControlName As String _
            )

    
    Dim cmw As CodeModuleWriter
    Set cmw = New CodeModuleWriter
    
    DoCmd.OpenForm ChartFormName, acDesign
    With Forms(ChartFormName)
      If .HasModule = False Then
         .HasModule = True
      End If
    End With
    
    cmw.Load "Form_" & ChartFormName
    
    CheckWebCtlControlSource Forms(ChartFormName).Controls(ChartControlName)
    
    RemoveBACFormCode cmw
    
    cmw.AddCodeToEventProcedure "Load", "Form", "    Me.TimerInterval = 2 ' load chart after form is completed"
    cmw.AddCodeToEventProcedure "UnLoad", "Form", "    Set m_Configuration = Nothing"
    cmw.AddCodeToEventProcedure "Timer", "Form", "    InitChart"
    cmw.AddCodeToEventProcedure "Timer", "Form", "    Me.TimerInterval = 0"
    
    cmw.AddCode GetBetterAccessChartFormCode(ChartControlName)

End Sub

Private Sub CheckWebCtlControlSource(ByVal WebCtl As Control)
   With WebCtl
      If .ControlSource <> "=""about:blank""" Then
         .ControlSource = "=""about:blank"""
      End If
   End With
End Sub

Private Sub RemoveBACFormCode(ByVal cmw As CodeModuleWriter)
    
#If EARLYBINDING = 1 Then
    Dim cm As CodeModule
#Else
    Dim cm As Object
#End If

    Dim StartLine As Long
    Dim EndLine As Long
    
    Dim StartMarkerLine As Long
    Dim EndMarkerLine As Long
    
    Set cm = cmw.CodeModule
    
    cm.Find FormCodeMarker_Start, StartMarkerLine, 0, 0, 0, True
    cm.Find FormCodeMarker_End, 0, 0, EndMarkerLine, 0, True
    
    If StartMarkerLine > 0 Then
        cm.DeleteLines StartMarkerLine, EndMarkerLine - StartMarkerLine + 1
    End If
    
End Sub

Private Function GetBetterAccessChartFormCode(ByVal ChartControlName As String) As String

    Dim str As String
   
    str = FormCodeMarker_Start & vbNewLine
    
    str = str & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"    Private m_Chart As BAC_Chart" & vbNewLine & _
"#Else" & vbNewLine & _
"    Private m_Chart As Object" & vbNewLine & _
"#End If" & vbNewLine
    str = str & vbNewLine & _
"Private WithEvents m_Configuration As Form" & vbNewLine & _
"Private m_ConfigFormReturnProperties As Object" & vbNewLine
    
    str = str & vbNewLine & _
"Private Sub InitChart()" & vbNewLine & _
"" & vbNewLine & _
"    Dim ChartProperties As Object" & vbNewLine & _
"" & vbNewLine & _
"    Set ChartProperties = BACx.ChartConfiguration.GetFormProperty(Me.Name, Me." & ChartControlName & ".Name & ""_BACproperties"")" & vbNewLine & _
"    If Not (ChartProperties Is Nothing) Then" & vbNewLine & _
"        CreateChart ChartProperties" & vbNewLine & _
"    End If" & vbNewLine & _
"" & vbNewLine & _
"End Sub" & vbNewLine & _
"" & vbNewLine & _
"Private Sub CreateChart(ByVal ChartProperties As Object)" & vbNewLine & _
"" & vbNewLine & _
"    Set m_Chart = BAC.Chart(Me." & ChartControlName & ")" & vbNewLine & _
"" & vbNewLine & _
"    m_Chart.Properties.SetProperties ChartProperties" & vbNewLine & _
"    m_Chart.DisplayIn = chDisplayIn.chWebBrowserControl" & vbNewLine & _
"    m_Chart.CreateChart" & vbNewLine & _
"" & vbNewLine & _
"End Sub" & vbNewLine

    str = str & vbNewLine & _
"Private Property Get ConfigurationForm() As Form" & vbNewLine & _
"" & vbNewLine & _
"    Dim ChartProperties As Object" & vbNewLine & _
"" & vbNewLine & _
"    If m_Configuration Is Nothing Then" & vbNewLine & _
"        If Not (m_Chart Is Nothing) Then" & vbNewLine & _
"            Set ChartProperties = m_Chart.Properties" & vbNewLine & _
"        Else" & vbNewLine & _
"            Set ChartProperties = BACx.ChartConfiguration.GetFormProperty(Me.Name, Me." & ChartControlName & ".Name & ""_BACproperties"")" & vbNewLine & _
"        End If" & vbNewLine & _
"        Set m_Configuration = BACx.NewConfigurationForm(m_ConfigFormReturnProperties, Me.Name, Me." & ChartControlName & ".Name, ChartProperties)" & vbNewLine & _
"        Me.SetFocus" & vbNewLine & _
"    End If" & vbNewLine & _
"" & vbNewLine & _
"    Set ConfigurationForm = m_Configuration" & vbNewLine & _
"" & vbNewLine & _
"End Property" & vbNewLine

    str = str & vbNewLine & _
"Private Sub " & ChartControlName & "_DblClick(Cancel As Integer)" & vbNewLine & _
"    If BACx.ChartShortcutMenu(Me.ctlWebbrowser) = 1 Then" & vbNewLine & _
"        ConfigurationForm.SetFocus" & vbNewLine & _
"    End If" & vbNewLine & _
"End Sub" & vbNewLine

    str = str & vbNewLine & _
"Private Sub m_Configuration_Unload(Cancel As Integer)" & vbNewLine & _
"    Set m_Configuration = Nothing" & vbNewLine & _
"    If Not (m_ConfigFormReturnProperties Is Nothing) Then" & vbNewLine & _
"        With m_ConfigFormReturnProperties" & vbNewLine & _
"            If .Item(""Commit"").Value = True Then" & vbNewLine & _
"                CreateChart .Item(""ChartProperties"").Properties" & vbNewLine & _
"            End If" & vbNewLine & _
"        End With" & vbNewLine & _
"        Set m_ConfigFormReturnProperties = Nothing" & vbNewLine & _
"    End If" & vbNewLine & _
"End Sub" & vbNewLine

    str = str & vbNewLine & FormCodeMarker_End & vbNewLine

    GetBetterAccessChartFormCode = str

End Function
