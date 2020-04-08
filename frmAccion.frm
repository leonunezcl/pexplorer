VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Respaldando ..."
   ClientHeight    =   1695
   ClientLeft      =   2670
   ClientTop       =   4965
   ClientWidth     =   5265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5265
   Begin VB.Frame fra 
      Caption         =   "Respaldando archivos. Espere un momento ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5175
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   345
         Left            =   90
         TabIndex        =   2
         Top             =   1125
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   135
         Picture         =   "frmAccion.frx":030A
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "lblArchivo"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   885
         Width           =   4950
      End
   End
End
Attribute VB_Name = "frmAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public total As Integer

Private Sub Form_Activate()
    Me.Refresh
End Sub

Private Sub Form_Load()

    Dim ret As Long
    Dim X As Integer
    Dim Y As Integer
    Dim wRect As RECT
    
    Call CenterWindow(hWnd)
    'Call CargarPantalla(Me)
    
    pgb.Min = 1
    pgb.Max = total
        
    ret = GetWindowRect(hWnd, wRect)
    
    X = (GetSystemMetrics(SM_CXSCREEN) - (wRect.Right - wRect.Left)) / 2
    Y = (GetSystemMetrics(SM_CYSCREEN) - (wRect.Bottom - wRect.Top + GetSystemMetrics(SM_CYCAPTION))) / 2
    
    ret = SetWindowPos(hWnd, HWND_TOPMOST, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim ret As Long
    
    ret = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmAccion = Nothing
    
End Sub


