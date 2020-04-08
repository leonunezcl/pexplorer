VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar Versión"
   ClientHeight    =   3480
   ClientLeft      =   1665
   ClientTop       =   1695
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
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
      Height          =   420
      Index           =   1
      Left            =   3180
      TabIndex        =   4
      Top             =   2955
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   1275
      TabIndex        =   3
      Top             =   2955
      Width           =   1215
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3465
      Left            =   0
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   2
      Top             =   0
      Width           =   360
   End
   Begin VB.Frame fra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2880
      Left            =   405
      TabIndex        =   0
      Top             =   -30
      Width           =   4935
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         Left            =   75
         TabIndex        =   1
         Top             =   195
         Width           =   4785
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        pShell C_WEB_PAGE_PE, hWnd
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    CenterWindow hWnd
    
    Dim Msg As String
    
    Msg = "Esta versión de " & App.Title & " tiene más de 3 meses." & vbNewLine & vbNewLine
    Msg = Msg & "Una nueva versión de este software puede estar disponible en la página web." & vbNewLine & vbNewLine
    Msg = Msg & "Desea comprobar una versión más reciente ?" & vbNewLine
    
    lblMsg.Caption = Msg
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff("Actualizar versión", picDraw)
    
    picDraw.Refresh
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmUpdate = Nothing
    
End Sub


