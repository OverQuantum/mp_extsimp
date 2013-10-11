VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' mp_extsimp
' Generalization of complex junctions and two ways roads
' from OpenStreetMap data
'
' Copyright © 2012-2013 OverQuantum
'
' Please refer to mp_extsimp.bas for details
'

' This form is only for visualization of optimization progress

Option Explicit

Private Sub Form_Load()
    Form1.Visible = True 'display form
    DoEvents
    Call OptimizeRouting(Command)
    End
End Sub
