VERSION 5.00
Begin VB.Form frmAnaFiltro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Filtro"
   ClientHeight    =   1440
   ClientLeft      =   2205
   ClientTop       =   1755
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnaFiltro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione opciones de filtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   375
      TabIndex        =   4
      Top             =   -15
      Width           =   2895
      Begin VB.OptionButton opt 
         Caption         =   "&Todo"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   8
         Top             =   1035
         Width           =   1335
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Optimización"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   7
         Top             =   780
         Width           =   1335
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Funcionalidad"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   6
         Top             =   525
         Width           =   1395
      End
      Begin VB.OptionButton opt 
         Caption         =   "&Estilo"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   255
         Width           =   720
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   1380
      Left            =   0
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   0
      Width           =   360
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
      Left            =   3330
      TabIndex        =   1
      Top             =   120
      Width           =   1200
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
      Left            =   3330
      TabIndex        =   0
      Top             =   555
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   405
      TabIndex        =   2
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmAnaFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
'filtra el listview de analisis
Private Sub FiltrarAnalisis()
                
    glbFiltroAnalisis = True
    
    Call Hourglass(hwnd, True)
    
    Main.lvwInfoAna.ListItems.Clear

Volver:
    If opt(3).Value Then
        glbFiltroVariables = True
        glbFiltroConstantes = True
        glbFiltroApis = True
        Call CargaProblemasAplicacion(0)
    Else
        If opt(0).Value Then CargaProblemasAplicacion (2)
        If opt(1).Value Then CargaProblemasAplicacion (3)
        If opt(2).Value Then CargaProblemasAplicacion (1)
    End If
    
    Call Hourglass(hwnd, False)
    
    glbFiltroAnalisis = False
    
End Sub
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call FiltrarAnalisis
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    CenterWindow hwnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff("Filtro", picDraw)
    
    picDraw.Refresh
    
    opt(3).Value = True
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmAnaFiltro = Nothing
    
End Sub



