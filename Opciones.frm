VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración"
   ClientHeight    =   5340
   ClientLeft      =   3090
   ClientTop       =   2370
   ClientWidth     =   7035
   Icon            =   "Opciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      Caption         =   "Visor de Analisis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   6870
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4560
         Left            =   2400
         TabIndex        =   21
         Top             =   450
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   8043
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"Opciones.frx":030A
      End
      Begin VB.ListBox lstCodigo 
         Height          =   2010
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   2175
      End
      Begin VB.PictureBox pic 
         Height          =   540
         Left            =   220
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   1
         Top             =   2520
         Width           =   1980
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   17
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   16
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   480
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   15
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   720
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   14
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   960
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   13
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C000C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   1200
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   12
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   1440
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   11
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   1680
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   10
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   0
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   9
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   240
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   480
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   720
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   960
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   1200
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   4
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   1440
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   3
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   1680
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   2
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Texto"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ejemplo"
         Height          =   195
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    CenterWindow hWnd
    
    lstCodigo.AddItem "Código Normal"
    lstCodigo.AddItem "Llamada a nivel de Módulo"
    lstCodigo.AddItem ""
    lstCodigo.AddItem ""
    
    fra(1).Left = fra(0).Left
    fra(1).Top = fra(0).Top
    fra(2).Left = fra(0).Left
    fra(2).Top = fra(0).Top
    
End Sub

Private Sub tabConf_Click()
        
End Sub


