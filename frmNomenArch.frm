VERSION 5.00
Begin VB.Form frmNomenArch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nomenclatura de Archivos"
   ClientHeight    =   3405
   ClientLeft      =   1905
   ClientTop       =   4695
   ClientWidth     =   4845
   Icon            =   "frmNomenArch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3525
      TabIndex        =   4
      Top             =   105
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Seleccione archivo a nomenclar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   405
      TabIndex        =   1
      Top             =   0
      Width           =   3045
      Begin VB.TextBox txtArchivo 
         Height          =   300
         Left            =   105
         TabIndex        =   3
         Top             =   2940
         Width           =   2850
      End
      Begin VB.ListBox lisArchivo 
         Height          =   2595
         Left            =   90
         TabIndex        =   2
         Top             =   255
         Width           =   2865
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   3345
      Left            =   0
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmNomenArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

'carga los archivos nomenclados
Private Sub CargaArchivosNomenclados()

    Dim k As Integer
    
    For k = 1 To UBound(glbAnaArchivos)
        lisArchivo.AddItem glbAnaArchivos(k).Clase
    Next k
    
    lisArchivo.ListIndex = 0
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Unload Me
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
    
    With mGradient
        .Angle = 90
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Call CargaArchivosNomenclados
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmNomenArch = Nothing
    
End Sub


Private Sub lisArchivo_Click()

    txtArchivo.Text = glbAnaArchivos(lisArchivo.ListIndex + 1).Nomenclatura
    
End Sub


Private Sub txtArchivo_Change()

    glbAnaArchivos(lisArchivo.ListIndex + 1).Nomenclatura = txtArchivo.Text
    
End Sub


