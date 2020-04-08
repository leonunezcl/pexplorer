VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones varias"
   ClientHeight    =   1785
   ClientLeft      =   2160
   ClientTop       =   1980
   ClientWidth     =   6060
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "A&plicar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   7
      Top             =   135
      Width           =   1395
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   6
      Top             =   585
      Width           =   1395
   End
   Begin VB.Frame fraMisc 
      Caption         =   "Misceláneas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   390
      TabIndex        =   1
      Top             =   30
      Width           =   4065
      Begin VB.CheckBox chkOpc 
         Caption         =   "&Recargar proyecto al reanalizar"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1185
         Width           =   2595
      End
      Begin VB.CheckBox chkOpc 
         Caption         =   "&Colorizar código"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Width           =   1500
      End
      Begin VB.CheckBox chkOpc 
         Caption         =   "&Ejecutar sonidos"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   465
         Width           =   1500
      End
      Begin VB.CheckBox chkOpc 
         Caption         =   "C&olorizar código analizado"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   705
         Width           =   2385
      End
      Begin VB.CheckBox chkOpc 
         Caption         =   "&Analizar proyecto automáticamente"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   945
         Width           =   2910
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   1725
      Left            =   0
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private cc As New GCommonDialog

Private Sub cmd_Click(Index As Integer)

    Dim k As Integer
    
    If Index = 0 Then
        'grabar opciones de archivo
        For k = 1 To UBound(Ana_Archivo)
            Call GrabaIni(C_INI, "ana_opciones", CStr(k), chkOpc(k - 1).Value)
            Ana_Opciones(k).Value = chkOpc(k - 1).Value
        Next k
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    
    Call Hourglass(hwnd, True)
    Call CenterWindow(hwnd)
    
    For k = 1 To UBound(Ana_Opciones)
        chkOpc(k - 1).Value = Ana_Opciones(k).Value
    Next k
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Call Hourglass(hwnd, False)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    On Local Error Resume Next
    
    Set cc = Nothing
    Set frmOpciones = Nothing
    
    Err = 0
    
End Sub


