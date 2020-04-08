VERSION 5.00
Begin VB.Form frmResumen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen"
   ClientHeight    =   4095
   ClientLeft      =   2745
   ClientTop       =   1845
   ClientWidth     =   4110
   Icon            =   "Resumen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   1440
      TabIndex        =   17
      Top             =   3600
      Width           =   1245
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   330
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen de Proyecto"
      Height          =   3465
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4035
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3000
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2670
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2340
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1980
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1650
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1290
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total de Líneas de Código"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Arreglos"
         Height          =   195
         Index           =   7
         Left            =   210
         TabIndex        =   9
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Funciones"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   7
         Top             =   2340
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Subs"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   6
         Top             =   2010
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Apis"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   5
         Top             =   2670
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Tipos"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   4
         Top             =   1650
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Enumeraciones"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   990
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Constantes"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Variables"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   360
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    txt(0).Text = Proyecto.TotalVariables
    txt(1).Text = Proyecto.TotalConstantes
    txt(2).Text = Proyecto.TotalEnumeraciones
    txt(3).Text = Proyecto.TotalArray
    txt(4).Text = Proyecto.TotalTipos
    txt(5).Text = Proyecto.TotalSubs
    txt(6).Text = Proyecto.TotalFunciones
    txt(7).Text = Proyecto.TotalApi
    txt(8).Text = Proyecto.TotalLineasDeCodigo
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmResumen = Nothing
    
End Sub


