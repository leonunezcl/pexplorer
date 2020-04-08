VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   2655
   ClientTop       =   1815
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6000
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   4650
      TabIndex        =   3
      Top             =   120
      Width           =   1245
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   3420
      Width           =   4515
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2460
      TabIndex        =   1
      Top             =   90
      Width           =   2085
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   2385
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim archivo As String
    Dim linea As String
    Dim Procedimiento As String
    Dim Funcion As String
    
    archivo = Dir1.Path & "\" & File1.List(File1.ListIndex)
    
    Combo1.Clear
    
    Open archivo For Input As #1
        Do While Not EOF(1)
            Line Input #1, linea
            linea = Trim$(linea)
            If InStr(linea, "Private Sub ") Then
                Procedimiento = linea 'Mid$(linea, InStr(linea, "Private Sub "))
                Combo1.AddItem Procedimiento
            ElseIf InStr(linea, "Public Sub ") Then
                Procedimiento = linea 'Mid$(linea, InStr(linea, "Public Sub "))
                Combo1.AddItem Procedimiento
            ElseIf InStr(linea, "Sub ") Then
                If Left$(linea, 3) = "Sub" Then
                    Procedimiento = linea ' Mid$(linea, InStr(linea, "Sub "))
                    Combo1.AddItem Procedimiento
                End If
            ElseIf InStr(linea, "Private Function ") Then
                Funcion = linea 'Mid$(linea, InStr(linea, "Private Sub "))
                Combo1.AddItem Funcion
            ElseIf InStr(linea, "Public Function ") Then
                Funcion = linea 'Mid$(linea, InStr(linea, "Private Sub "))
                Combo1.AddItem Funcion
            ElseIf InStr(linea, "Function ") Then
                If Left$(linea, 8) = "Function" Then
                    Funcion = linea 'Mid$(linea, InStr(linea, "Private Sub "))
                    Combo1.AddItem Funcion
                End If
            End If
        Loop
    Close #1
    
    ComboAutoSize Me, Combo1
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub


