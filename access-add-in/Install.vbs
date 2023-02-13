'Author: Josef Poetzl

const AddInName = "BetterAccessCharts"
const AddInFileName = "BetterAccessCharts.accda"
const MsgBoxTitle = "Install Better-Access-Charts"

MsgBox "Before updating the add-in file, the add-in must not be loaded!" & chr(13) & _
       "For security, close all Access instances.", , MsgBoxTitle & ": Hint"

Select Case MsgBox("Do you want to use the add-in as accde?" + chr(13) & _
                   "(Add-in is copied compiled to the add-in directory.)", 3, MsgBoxTitle)
   case 6 ' vbYes
      CreateMde GetSourceFileFullName, GetDestFileFullName
	  MsgBox "Add-in has been composed and changed to '" + GetAddInLocation + "' Stored", , MsgBoxTitle
   case 7 ' vbNo
      FileCopy GetSourceFileFullName, GetDestFileFullName
	  MsgBox "Add-in was added to '" + GetAddInLocation + "' Stored", , MsgBoxTitle
   case else
      
End Select


'##################################################
' Auxiliary functions:

Function GetSourceFileFullName()
   GetSourceFileFullName = GetScriptLocation & AddInFileName 
End Function

Function GetDestFileFullName()
   GetDestFileFullName = GetAddInLocation & AddInFileName 
End Function

Function GetScriptLocation()

   With WScript
      GetScriptLocation = Replace(.ScriptFullName & ":", .ScriptName & ":", "") 
   End With

End Function

Function GetAddInLocation()

   GetAddInLocation = GetAppDataLocation & "Microsoft\AddIns\"

End Function

Function GetAppDataLocation()

   Set wsShell = CreateObject("WScript.Shell")
   GetAppDataLocation = wsShell.ExpandEnvironmentStrings("%APPDATA%") & "\"

End Function

Function FileCopy(SourceFilePath, DestFilePath)

   set fso = CreateObject("Scripting.FileSystemObject") 
   fso.CopyFile SourceFilePath, DestFilePath

End Function

Function CreateMde(SourceFilePath, DestFilePath)

   Set AccessApp = CreateObject("Access.Application")
   AccessApp.SysCmd 603, (SourceFilePath), (DestFilePath)

End Function