VERSION 5.00
Begin VB.Form frmCorregir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corregir aplicación"
   ClientHeight    =   6255
   ClientLeft      =   1590
   ClientTop       =   1755
   ClientWidth     =   5940
   Icon            =   "frmCorregir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFiles 
      Appearance      =   0  'Flat
      Height          =   4305
      Left            =   390
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   1920
      Width           =   4155
   End
   Begin VB.CheckBox chkSel 
      Caption         =   "&Todos"
      Height          =   195
      Left            =   3660
      TabIndex        =   10
      Top             =   1695
      Value           =   1  'Checked
      Width           =   900
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Detener"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   4650
      TabIndex        =   9
      Top             =   660
      Width           =   1215
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
      Height          =   360
      Index           =   1
      Left            =   4650
      TabIndex        =   8
      Top             =   1095
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Corregir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4650
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   390
      TabIndex        =   3
      Top             =   525
      Width           =   4155
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "- Arrays , Constantes, Variables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   495
         Width           =   2685
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "- Funciones y Procedimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   5
         Top             =   735
         Width           =   2505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "La correción del proyecto comenta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   3090
      End
   End
   Begin VB.TextBox txtArchivo 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   225
      Width           =   4140
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   6225
      Left            =   0
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar archivos a corregir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   12
      Top             =   1695
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Proyecto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   1
      Top             =   30
      Width           =   765
   End
End
Attribute VB_Name = "frmCorregir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Private Sub CargarArchivos()

    Dim k As Integer
    
    For k = 1 To UBound(Proyecto.aArchivos)
        lstFiles.AddItem Proyecto.aArchivos(k).Descripcion
        lstFiles.Selected(lstFiles.NewIndex) = True
    Next k
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
    
    ElseIf Index = 1 Then
    
    Else
    
    End If
    
End Sub

Private Sub Form_Load()

    CenterWindow hwnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Call CargarArchivos
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Set frmCorregir = Nothing
    
End Sub


Private Sub Label5_Click()

End Sub


