VERSION 5.00
Begin VB.Form frmResumenDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Detallado de Proyecto"
   ClientHeight    =   6045
   ClientLeft      =   3405
   ClientTop       =   2025
   ClientWidth     =   4080
   Icon            =   "ResumenDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Resumen de Proyecto"
      Height          =   5355
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   4035
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   4950
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   4620
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   4290
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3960
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3630
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   3300
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2970
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   300
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1290
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1650
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1980
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2310
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Funciones Públicas"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   24
         Top             =   4620
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Subs Públicas"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Tipos Públicos"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   22
         Top             =   3300
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Arreglos Públicos"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   21
         Top             =   2610
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Enumeraciones Públicas"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   20
         Top             =   1980
         Width           =   2370
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Constantes Publicas"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   1290
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Variables Públicas"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Variables Privadas"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Constantes Privadas"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   990
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Enumeraciones Privadas"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1650
         Width           =   2385
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Tipos Privados"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   2970
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Apis"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   4920
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Subs Privadas"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   3630
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Funciones Privadas"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   4290
         Width           =   2025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Arreglos Privados"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2310
         Width           =   1860
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   1380
      TabIndex        =   0
      Top             =   5550
      Width           =   1245
   End
End
Attribute VB_Name = "frmResumenDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    txt(0).Text = Proyecto.TotalVariablesPrivadas
    txt(1).Text = Proyecto.TotalVariablesPublicas
    
    txt(2).Text = Proyecto.TotalConstantesPrivadas
    txt(3).Text = Proyecto.TotalConstantesPublicas
    
    txt(4).Text = Proyecto.TotalEnumeracionesPrivadas
    txt(5).Text = Proyecto.TotalEnumeracionesPublicas
    
    txt(6).Text = Proyecto.TotalArrayPrivadas
    txt(7).Text = Proyecto.TotalArrayPublicas
    
    txt(8).Text = Proyecto.TotalTiposPrivadas
    txt(9).Text = Proyecto.TotalTiposPublicas
    
    txt(10).Text = Proyecto.TotalSubsPrivadas
    txt(11).Text = Proyecto.TotalSubsPublicas
    
    txt(12).Text = Proyecto.TotalFuncionesPrivadas
    txt(13).Text = Proyecto.TotalFuncionesPublicas
    
    txt(14).Text = Proyecto.TotalApi
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmResumenDetalle = Nothing
    
End Sub


