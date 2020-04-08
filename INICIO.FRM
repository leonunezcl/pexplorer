VERSION 5.00
Begin VB.Form frmInicio 
   BorderStyle     =   0  'None
   ClientHeight    =   1155
   ClientLeft      =   2400
   ClientTop       =   2235
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMainSkin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   0
      Picture         =   "Inicio.frx":0000
      ScaleHeight     =   1350
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    CenterWindow hWnd
    
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE)
     
    Dim WindowRegion As Long
    
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
    
End Sub


