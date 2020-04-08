VERSION 5.00
Begin VB.Form frmFind 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar"
   ClientHeight    =   1185
   ClientLeft      =   2805
   ClientTop       =   4005
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   79
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   1140
      Left            =   0
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   6
      Top             =   0
      Width           =   360
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1155
      TabIndex        =   1
      Top             =   150
      Width           =   3405
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   4665
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "B&uscar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   4665
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "S&olo palabras completas"
      Height          =   255
      Left            =   435
      TabIndex        =   3
      Top             =   870
      Width           =   2805
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "&Según Modelo"
      Height          =   255
      Left            =   435
      TabIndex        =   2
      Top             =   570
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Buscar:"
      Height          =   195
      Left            =   435
      TabIndex        =   0
      Top             =   180
      Width           =   660
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

'busca texto seleccionado
Sub FindText()

    On Local Error GoTo SalirFindText
    
    Dim lWhere, lPos As Long
    Dim sTmp As String
    Dim texto As String
    Dim iComp As Integer
    
    If gbMatchCase = 0 Then iComp = 1 Else iComp = 0
        
    If gbLastPos = 0 Then
        lPos = 1
    Else
        lPos = gbLastPos
    End If
    
    texto = UCase$(Main.txtRutina.text)
    
    Do While lPos < Len(texto)
        
        sTmp = Mid$(texto, lPos, Len(texto))
        
        lWhere = InStr(sTmp, UCase$(gsFindText))
        lPos = lPos + lWhere
        
        If lWhere Then   ' If found,
            Main.txtRutina.SetFocus
            Main.txtRutina.SelStart = lPos - 2   ' set selection start and
            Main.txtRutina.SelLength = Len(gsFindText)   ' set selection length.   Else
            
            gbLastPos = lPos
            Exit Do
        Else
            gbLastPos = 0
            Exit Do 'we are ready
        End If
    Loop
    
    Exit Sub
    
SalirFindText:
    gbLastPos = 0
    Err = 0
    
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    gbMatchCase = chkMatchCase.Value
    gbWholeWord = chkWholeWord.Value
    gsFindText = txtFind.text
    Call FindText
End Sub

Private Sub Form_Load()
    
    CenterWindow hwnd
    
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
    
    chkMatchCase.Value = gbMatchCase
    chkWholeWord.Value = gbWholeWord
    txtFind.text = gsFindText
    txtFind.SelLength = Len(gsFindText)
            
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmFind = Nothing
    
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = (txtFind.text <> "")
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
    End If
    
End Sub

