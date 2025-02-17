VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiWindowTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module: WinApiWindowTools
'---------------------------------------------------------------------------------------
'/**
' <summary>
' WinAPI-Funktionen: Window Handle & Co.
' </summary>
' <remarks>
' </remarks>
' \ingroup WinAPI
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiWindowTools.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Die Prozeduren (GetMDI, GetHeaderSection, GetDetailSection, GetFooterSection und GetControl
' stammen aus dem AEK10-Vortrag von J�rg Ostendorp
'
'----------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare PtrSafe Function ClientToScreen Lib "user32.dll" ( _
         ByVal Hwnd As LongPtr, _
         ByRef lpPoint As POINTAPI _
      ) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
         ByVal HWnd1 As LongPtr, _
         ByVal HWnd2 As LongPtr, _
         ByVal Lpsz1 As String, _
         ByVal Lpsz2 As String _
      ) As LongPtr

'---------------------------------------------------------------------------------------
' Function: GetMDI
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle des MDI-Client-Fensters.
' </summary>
' <returns>Handle (Long)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetMDI() As LongPtr
   Dim h As LongPtr
   h = Application.hWndAccessApp
   'Erstes (und einziges) "MDIClient"-Kindfenster des Applikationsfensters suchen
   GetMDI = FindWindowEx(h, 0&, "MDIClient", vbNullString)
End Function

'---------------------------------------------------------------------------------------
' Function: GetHeaderSection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle f�r den Kopfbereich eines Formulares
' </summary>
' <param name="fHwnd">Handle des Formulars (Form.Hwnd)</param>
' <returns>Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetHeaderSection(ByVal Hwnd As LongPtr) As LongPtr
   Dim h As LongPtr
   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln
   h = FindWindowEx(Hwnd, 0&, "OformSub", vbNullString)
   GetHeaderSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetDetailSection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle f�r den Detailbereich eines Formulares
' </summary>
' <param name="fHwnd">Handle des Formulars (Form.Hwnd)</param>
' <returns>Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetDetailSection(ByVal Hwnd As LongPtr) As LongPtr
   Dim h As LongPtr
   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln, beginnend
   'nach dem Kopfbereich
   h = GetHeaderSection(Hwnd)
   h = FindWindowEx(Hwnd, h, "OformSub", vbNullString)
   GetDetailSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetFooterSection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle f�r den Fu�bereich eines Formulares
' </summary>
' <param name="fHwnd">Handle des Formulars (Form.Hwnd)</param>
' <returns>Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFooterSection(ByVal Hwnd As Long) As LongPtr
   Dim h As LongPtr
   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln, beginnend
   'nach dem Detailbereich
   h = GetDetailSection(Hwnd)
   h = FindWindowEx(Hwnd, h, "OformSub", vbNullString)
   GetFooterSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetControl
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle eines beliebigen Controls
' </summary>
' <param name="frm">Formular-Referenz</param>
' <param name="sHwnd">Handle des Bereichs, auf dem sich das Control befindet (Header, Detail, Footer)</param>
' <param name="ClassName">Name der Fensterklasse des Controls</param>
' <param name="ControlName">Name des Controls</param>
' <returns>Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetControl(ByRef FrmRef As Access.Form, ByVal Hwnd As LongPtr, _
                           ByVal ClassName As String, ByVal ControlName As String) As LongPtr

   'Ermittelt den Handle eines beliebigen Controls

   'Parameter:
   ' frm - Formular
   ' Handle des Bereichs, auf dem sich das Control befindet (Header, Detail, Footer)
   ' ControName - Name der Fensterklasse des Controls
   ' ControlName - Name des Controls


   'Exitieren mehrere Controls der gleichen Klasse auf einem Formular, z.B. TabControls, besteht das Problem, da�
   'deren Reihenfolge nicht definiert ist (anders also als bei den Sektionsfenstern)
   'In diesem Fall kann man alle Kindfenster dieser Klasse in einer Schleife durchlaufen
   'und z.B. pr�fen, ob die Position des Fensters des zur�ckgegebenen Handles
   'mit der des Access-Steuerelementes �bereinstimmt.
   'Nachfolgend wird hierf�r die undokumentierte Funktion accHittest verwendet.
   'Dieser werden als Parameter die Screenkoordinaten der linken oberen Ecke eines
   'Steuerelementes �bergeben. Befindet sich dort ein Objekt, erh�lt man dieses als R�ckgabewert.
   'Ist der Name des Objektes identisch mit dem �bergebenen Steuerelementnamen, so
   'hat man das Handle ermittelt:

On Error Resume Next

   Dim h As LongPtr
   Dim obj As Object
   Dim pt As POINTAPI

   h = 0

   Do
      'Erstes (h=0)/n�chstes (h<>0) Control auf dem Sektionsfenster ermitteln
      h = FindWindowEx(Hwnd, h, ClassName, vbNullString)

      'Bildschirmkoordinaten dieses Controls ermitteln
      'daf�r die Punktkoordinaten aus dem letzten Durchlauf zur�cksetzen, sonst wird addiert!
      pt.X = 0
      pt.Y = 0
      ClientToScreen h, pt

      'Objekt bei den Koordinaten ermitteln
      Set obj = FrmRef.accHitTest(pt.X, pt.Y)

      'Wenn Objektname = Tabname Ausstieg aus der Schleife
      If obj.Name = ControlName Then
         Exit Do
      End If
   Loop While h <> 0

   'Handle zur�ckgeben
   GetControl = h

End Function
