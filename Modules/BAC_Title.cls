VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2022 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.30.02  published: 30.12.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_Font As BAC_Font

'### Members
Private m_Properties As BAC_Properties

' ### Properties
Public Property Get Properties() As BAC_Properties
   Set Properties = m_Properties
End Property

Public Property Get Font() As BAC_Font
    Set Font = m_Font
End Property

Public Property Let Show(ByVal This_Show As Boolean)
    m_Properties.Item("Show").Value = This_Show
End Property
Public Property Get Show() As Boolean
    Show = Nz(m_Properties.Item("Show").Value, False)
End Property

Public Property Let Position(ByVal This_Position As chPosition)
    m_Properties.Item("Position").Value = This_Position
End Property
Public Property Get Position() As chPosition
    Position = Nz(m_Properties.Item("Position").Value, 0)
End Property

Public Property Let Text(ByVal This_Text As String)
    m_Properties.Item("Text").Value = This_Text
End Property
Public Property Get Text() As String
    Text = Nz(m_Properties.Item("Text").Value, vbNullString)
End Property


'### Instancing

Private Sub Class_Initialize()

    Set m_Font = New BAC_Font
    InitProperties
    
End Sub

Private Sub InitProperties()

   Set m_Properties = New BAC_Properties
   Set m_Properties.NewProperty("Font").Properties = m_Font.Properties
   m_Properties.NewProperty "Show"
   m_Properties.NewProperty "Position"
   m_Properties.NewProperty "Text"
   
End Sub


' ### Generator Function

Friend Function getTitle() As String

    With New BAC_StringBuilder
        If Me.Show = False Then
            .AddLine "            title: {"
            .AddLine "                display: false,"
            .AddLine "            },"
        Else
            .AddLine "            title: {"
            .AddLine "                display: true,"
            .AddLine "                position: " & BAC.Helper.getPositionString(Me.Position) & ","
            .AddLine "                text: '" & Me.Text & "',"
            If Me.Font.Family <> vbNullString Then
                .AddLine "                    fontFamily: '" & Me.Font.Family & "',"
            End If
            If Me.Font.Size > 0 Then
                .AddLine "                    fontSize: " & Me.Font.Size & ","
            End If
            If Me.Font.Color <> vbNullString Then
                .AddLine "                    fontColor: '" & Me.Font.Color & "',"
            End If
            If Me.Font.IsBold = True And Me.Font.IsItalic = True Then
                .AddLine "                    fontStyle: 'bold italic'"
            ElseIf Me.Font.IsBold = True Then
                .AddLine "                    fontStyle: 'bold'"
            ElseIf Me.Font.IsItalic = True Then
                .AddLine "                    fontStyle: 'italic'"
            Else
                'Do nothing
            End If
            .CutTrailingComma
            .AddLine "            },"
        End If
        
        getTitle = .ToString
    End With
    
End Function
