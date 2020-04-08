VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmEstadisticas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadisticas"
   ClientHeight    =   5325
   ClientLeft      =   2070
   ClientTop       =   2985
   ClientWidth     =   7200
   Icon            =   "Estadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      Caption         =   "Estadisticas del Proyecto"
      Height          =   5250
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   7125
      Begin VB.Frame fraTab 
         BorderStyle     =   0  'None
         Height          =   4185
         Left            =   165
         TabIndex        =   4
         Top             =   900
         Width           =   6750
         Begin VB.Frame fra 
            Caption         =   "Archivos de Menor Tamaño en KBytes"
            Height          =   1965
            Index           =   2
            Left            =   0
            TabIndex        =   16
            Top             =   2190
            Width           =   6750
            Begin VB.TextBox txtP 
               Height          =   285
               Index           =   4
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   1560
               Width           =   4470
            End
            Begin VB.TextBox txtP 
               Height          =   285
               Index           =   3
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   1245
               Width           =   4470
            End
            Begin VB.TextBox txtP 
               Height          =   285
               Index           =   2
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   915
               Width           =   4470
            End
            Begin VB.TextBox txtP 
               Height          =   285
               Index           =   1
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   585
               Width           =   4470
            End
            Begin VB.TextBox txtP 
               Height          =   285
               Index           =   0
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   255
               Width           =   4470
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Página de Propiedades"
               Height          =   195
               Left            =   135
               TabIndex        =   26
               Top             =   1590
               Width           =   1650
            End
            Begin VB.Label Label8 
               Caption         =   "Control de Usuario"
               Height          =   225
               Left            =   135
               TabIndex        =   25
               Top             =   1260
               Width           =   1365
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Módulo de Clase"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   945
               Width           =   1185
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Módulo BAS"
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   615
               Width           =   885
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Formulario"
               Height          =   195
               Left            =   135
               TabIndex        =   22
               Top             =   285
               Width           =   720
            End
         End
         Begin VB.Frame fra 
            Caption         =   "Archivos de Mayor Tamaño en KBytes"
            Height          =   1965
            Index           =   1
            Left            =   0
            TabIndex        =   5
            Top             =   165
            Width           =   6750
            Begin VB.TextBox txtG 
               Height          =   285
               Index           =   0
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   255
               Width           =   4470
            End
            Begin VB.TextBox txtG 
               Height          =   285
               Index           =   1
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   585
               Width           =   4470
            End
            Begin VB.TextBox txtG 
               Height          =   285
               Index           =   2
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   915
               Width           =   4470
            End
            Begin VB.TextBox txtG 
               Height          =   285
               Index           =   3
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   1245
               Width           =   4470
            End
            Begin VB.TextBox txtG 
               Height          =   285
               Index           =   4
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   1560
               Width           =   4470
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Formulario"
               Height          =   195
               Left            =   135
               TabIndex        =   15
               Top             =   285
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Módulo BAS"
               Height          =   195
               Left            =   120
               TabIndex        =   14
               Top             =   615
               Width           =   885
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Módulo de Clase"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   945
               Width           =   1185
            End
            Begin VB.Label Label5 
               Caption         =   "Control de Usuario"
               Height          =   225
               Left            =   135
               TabIndex        =   12
               Top             =   1260
               Width           =   1365
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Página de Propiedades"
               Height          =   195
               Left            =   135
               TabIndex        =   11
               Top             =   1590
               Width           =   1650
            End
         End
      End
      Begin ComctlLib.TabStrip tabEst 
         Height          =   4590
         Left            =   90
         TabIndex        =   3
         Top             =   570
         Width           =   6930
         _ExtentX        =   12224
         _ExtentY        =   8096
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Archivos del Proyecto"
               Key             =   ""
               Object.Tag             =   ""
               Object.ToolTipText     =   "Estadisticas de sobre los tamaños de archivos"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblProyecto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "NOMBRE DEL PROYECTO"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   750
         TabIndex        =   2
         Top             =   300
         Width           =   6240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre "
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   285
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    CenterWindow hwnd
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmEstadisticas = Nothing
    
End Sub


