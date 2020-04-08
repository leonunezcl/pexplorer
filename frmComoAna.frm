VERSION 5.00
Begin VB.Form frmComoAna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione forma de analizar"
   ClientHeight    =   3480
   ClientLeft      =   2460
   ClientTop       =   2250
   ClientWidth     =   4455
   Icon            =   "frmComoAna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   60
      TabIndex        =   7
      Top             =   1410
      Width           =   2970
      Begin VB.Label lblhelp 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   2790
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
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
      Index           =   1
      Left            =   3165
      TabIndex        =   6
      Top             =   585
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
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
      Left            =   3165
      TabIndex        =   5
      Top             =   150
      Width           =   1200
   End
   Begin VB.Frame fra 
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   2985
      Begin VB.OptionButton opt 
         Caption         =   "&Personalizada"
         Height          =   210
         Index           =   3
         Left            =   135
         TabIndex        =   4
         ToolTipText     =   "Analizar segun opciones definidas"
         Top             =   1035
         Width           =   1650
      End
      Begin VB.OptionButton opt 
         Caption         =   "Mí&nima"
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "Analizar solo lo no usado"
         Top             =   795
         Width           =   1530
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Media"
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Analizar los elementos no usados y medir complejidad"
         Top             =   540
         Width           =   1470
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Full"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   1
         ToolTipText     =   "Analizar todo el software"
         Top             =   285
         Value           =   -1  'True
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmComoAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If opt(0).Value = True Then
            glbComoAnalizar = FULL
        ElseIf opt(1).Value = True Then
            glbComoAnalizar = MEDIA
        ElseIf opt(2).Value = True Then
            glbComoAnalizar = MINIMA
        Else
            glbComoAnalizar = PERSONALIZADA
            MsgBox "Opción aun no soportada.", vbCritical
            Exit Sub
        End If
    Else
        glbComoAnalizar = CANCELADO
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    CenterWindow hwnd
    
    opt_Click 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmComoAna = Nothing
End Sub


Private Sub opt_Click(Index As Integer)
    
    Dim Msg As String
    
    If Index = 0 Then
        Msg = "Se analizaran todos los problemas que pueda presentar el proyecto a nivel de optimización, funcionalidad y complejidad de este."
    ElseIf Index = 1 Then
        Msg = "Solo se analizaran todos los elementos que componen el proyecto y que esten declarados pero no esten siendo usados y la complejidad de los procedimientos."
    ElseIf Index = 2 Then   'Minima
        Msg = "Solo se analizaran todos los elementos que componen el proyecto y que esten declarados pero no esten siendo usados."
    Else
        Msg = "Se analizaran todas los opciones definidas solo por el usuario."
    End If
    
    lblhelp.Caption = Msg
    
End Sub


