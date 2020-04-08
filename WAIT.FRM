VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analizando proyecto ..."
   ClientHeight    =   495
   ClientLeft      =   2865
   ClientTop       =   4125
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Me.Refresh
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmWait = Nothing
End Sub


