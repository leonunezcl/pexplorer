VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sugerencia del día"
   ClientHeight    =   3330
   ClientLeft      =   3405
   ClientTop       =   2880
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   3330
      Left            =   0
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   6
      Top             =   0
      Width           =   360
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Mostrar sugerencias al iniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   420
      TabIndex        =   3
      Top             =   2970
      Width           =   2970
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Siguiente sugerencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4215
      TabIndex        =   2
      Top             =   525
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   405
      ScaleHeight     =   2895
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   15
      Width           =   3735
      Begin VB.Image imgTip 
         Height          =   480
         Left            =   45
         Picture         =   "frmTip.frx":030A
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sabía que..."
         Height          =   255
         Left            =   630
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1965
         Left            =   90
         TabIndex        =   4
         Top             =   630
         Width           =   3510
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
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
      Left            =   4215
      TabIndex        =   0
      Top             =   45
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

' La base de datos en memoria de sugerencias.
Dim Tips As New Collection

' Nombre del archivo de sugerencias
Const TIP_FILE = "TIPOFDAY.TXT"

' Índice en la colección de la sugerencia actualmente mostrada.
Dim CurrentTip As Long

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Private Sub DoNextTip()

    ' Seleccionar una sugerencia aleatoriamente.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' O recorrer secuencialmente las sugerencias

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Mostrar.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Leer cada sugerencia desde archivo.
    Dim InFile As Integer   ' Descriptor para archivo.
    
    ' Obtener el siguiente descriptor de archivo libre.
    InFile = FreeFile
    
    ' Asegurarse de que se especifica un archivo.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Asegurarse de que el archivo existe antes de intentar abrirlo.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Leer la colección desde un archivo de texto.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Mostrar una sugerencia aleatoriamente.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' guardar si este formulario debe mostrarse o no al iniciar
    SaveSetting App.ExeName, "Opciones", "Mostrar sugerencias al iniciar", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
            
    CenterWindow hwnd
    
    ' Establecer la casilla de verificación, que obligará a que el valor se vuelva a escribir en el Registro
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Semilla aleatoria
    Randomize
    
    ' Leer el archivo de sugerencias y mostrar una sugerencia aleatoriamente.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "de que no se ha encontrado el archivo " & TIP_FILE & vbCrLf & vbCrLf & _
           "Cree un archivo de texto llamado " & TIP_FILE & " con el Bloc de notas, con una sugerencia por línea. " & _
           "A continuación, colóquelo en el mismo directorio que la aplicación."
    End If

    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
End Sub

