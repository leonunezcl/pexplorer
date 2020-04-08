VERSION 5.00
Begin VB.Form frmInfoAna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información de Análisis"
   ClientHeight    =   5700
   ClientLeft      =   3555
   ClientTop       =   1740
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoAna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   5625
      Left            =   15
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   10
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
      Height          =   435
      Left            =   6315
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Información"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      Left            =   435
      TabIndex        =   0
      Top             =   30
      Width           =   5790
      Begin VB.TextBox txtComentario 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1695
         Width           =   5550
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
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
         Left            =   105
         TabIndex        =   7
         Top             =   1455
         Width           =   990
      End
      Begin VB.Label lblTipo 
         Caption         =   "TIPO"
         Height          =   270
         Left            =   960
         TabIndex        =   6
         Top             =   1200
         Width           =   4710
      End
      Begin VB.Label lblUbicacion 
         Caption         =   "UBICACION"
         Height          =   270
         Left            =   975
         TabIndex        =   5
         Top             =   900
         Width           =   4710
      End
      Begin VB.Label lblProblema 
         Caption         =   "PROBLEMA"
         Height          =   615
         Left            =   975
         TabIndex        =   4
         Top             =   300
         Width           =   4710
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Index           =   2
         Left            =   105
         TabIndex        =   3
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación"
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
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   885
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Problema"
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
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmInfoAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Itmx As ListItem
Private mGradient As New clsGradient
'muestra la info
Private Sub MuestraInformacion()

    Dim nro As Integer
    Dim k As Integer
    Dim sNro As String
        
    Set Itmx = Main.lvwInfoAna.SelectedItem
    
    lblProblema.Caption = Itmx.SubItems(1)
    lblUbicacion.Caption = Itmx.SubItems(2)
    lblTipo.Caption = Itmx.SubItems(3)
    
    For k = 1 To UBound(Arr_Analisis)
        If Arr_Analisis(k).Problema = Itmx.SubItems(1) Then
            If Arr_Analisis(k).Ubicacion = Itmx.SubItems(2) Then
                If Arr_Analisis(k).Tipo = Itmx.SubItems(3) Then
                    If Arr_Analisis(k).Help <> -1 Then
                        txtComentario.text = LoadResString(Arr_Analisis(k).Help)
                        Exit For
                    End If
                End If
            End If
        End If
    Next k
    
End Sub

Private Sub cmd_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
    
    Call MuestraInformacion
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Itmx = Nothing
    Set mGradient = Nothing
    Set frmInfoAna = Nothing
    
End Sub


