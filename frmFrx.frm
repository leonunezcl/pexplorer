VERSION 5.00
Begin VB.Form frmFrx 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visualizador de Recursos binarios"
   ClientHeight    =   4770
   ClientLeft      =   1830
   ClientTop       =   2025
   ClientWidth     =   8400
   ClipControls    =   0   'False
   Icon            =   "frmFrx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4545
      TabIndex        =   5
      Top             =   4125
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   4680
      ScaleHeight     =   3315
      ScaleWidth      =   3570
      TabIndex        =   3
      Top             =   360
      Width           =   3630
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Guardar gráfico seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2055
      TabIndex        =   2
      Top             =   4125
      Width           =   2010
   End
   Begin VB.ListBox List1 
      Height          =   3360
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   4275
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   3780
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "File type                Hdr Size     File offset   Image size"
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   150
      Width           =   4275
   End
End
Attribute VB_Name = "frmFrx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' Demonstrates how to extract any and all graphics in VB FRX, CTX,
' DSX, DOX, and PGX binary files, display them, and save them to file.

' Key info: the original graphics files themselves are stored in their
' entirety in VB binary files. It's just a matter of finding where they
' start and end, loading them in a byte array, determining what
' graphic format they are, and they're yours. That's it...

Private Const LB_SETTABSTOPS = &H192
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

Private m_cff As New cFrxFile
Private m_sCaption As String
'
Public sPath As String
Private Sub CargaRecursos()

    Dim cfi As cFrxItem
    
    MousePointer = vbHourglass
    ' the file is read on this assignment, and may take some time if invalid.
    m_cff.Path = sPath
    MousePointer = vbNormal
    
    If m_cff.Count Then
      Caption = m_sCaption & " - " & sPath
      List1.Clear
      Label2 = "File size: " & m_cff.FileSize
      Set Picture1 = Nothing
      cmdSave.Enabled = False
      DoEvents
      
      For Each cfi In m_cff
        List1.AddItem cfi.FileTypeName & vbTab & cfi.HeaderSize & _
                              vbTab & cfi.FileOffset & vbTab & cfi.ImageSize
      Next
    Else
      MsgBox "No se encontraron archivos graficos en " & sPath, vbCritical
      cmdSave.Enabled = False
    End If
    
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Dim adwTabs(3) As Long
  
  Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
  m_sCaption = Caption
  Label1 = "File type                Hdr Size     File offset   Image size"
  Label2 = ""
  Picture1.TabStop = 0
  cmdSave.Enabled = False
  
  adwTabs(0) = 0
  adwTabs(1) = -85
  adwTabs(2) = -125
  adwTabs(3) = -165
  Call SendMessage(List1.hwnd, LB_SETTABSTOPS, 4, adwTabs(0))
    Call CargaRecursos
End Sub

Private Sub Form_Resize()
  On Error GoTo Out
  
  'Picture1.Width = (ScaleWidth - Picture1.Left) - (cmdOpen.Height * 0.5)
  'Picture1.Height = (ScaleHeight - Picture1.Top) - (cmdOpen.Height * 0.5)

Out:
End Sub

Private Sub List1_Click()
  
  With m_cff(List1.ListIndex + 1)
    If .PictureType Then
      Set Picture1 = .Picture
    Else
      Set Picture1 = Nothing
      Picture1.Cls
      ' StrConv will err if passed an uninitialized byte array
      ' and overflows with strings > 32K chars
      If .ImageSize And (.ImageSize < 2 ^ 15) Then
        Picture1.Print StrConv(.Bits, vbUnicode)
      Else
        Picture1.Print "Unable to display data"
      End If
    End If
  End With
  
  cmdSave.Enabled = True

End Sub

Private Sub cmdSave_Click()
  Dim sFilter As String
  Dim sPath As String
  Dim ff As Integer
  Dim ab() As Byte
  
  If (List1.ListIndex = -1) Then
    cmdSave.Enabled = False
    Exit Sub
  End If
  
  With m_cff(List1.ListIndex + 1)
    sFilter = .FileTypeName & " (*." & .FileExtension & ")" & vbNullChar & "*." & .FileExtension & vbNullChar & _
                 "All Picture Files" & vbNullChar & "*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur" & vbNullChar & _
                 "All Files (*.*)" & vbNullChar & "*.*"
  End With
  
  If GetSaveFilePath(hwnd, sFilter, 0, m_cff(List1.ListIndex + 1).FileExtension, "", "", "", sPath) Then
    On Error Resume Next
    Kill sPath
    On Error GoTo 0
    ff = FreeFile
    Open sPath For Binary As ff
    ' have to copy the array or the entire SafeArray struct gets
    ' written... (the Bits prop is a Variant)
    ab() = m_cff(List1.ListIndex + 1).Bits
    Put #ff, , ab()
    Close ff
  End If
  
End Sub
