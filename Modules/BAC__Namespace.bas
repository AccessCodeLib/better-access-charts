Attribute VB_Name = "BAC__Namespace"
'###########################################################################################
'# Copyright (c) 2020 - 2023 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 3.04.16  published: 27.03.2023                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_BetterAccessCharts As BAC__Factory


'### Properties

Public Property Get BetterAccessCharts() As BAC__Factory
  If m_BetterAccessCharts Is Nothing Then Set m_BetterAccessCharts = New BAC__Factory
  Set BetterAccessCharts = m_BetterAccessCharts
End Property

Public Property Get BAC() As BAC__Factory
  Set BAC = BetterAccessCharts
End Property


