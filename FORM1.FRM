VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   2745
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7200
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_FILENAME_LEN = 260
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Sub Form_Load()
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
   Dim i As Integer, s2 As String
   Const sFile = "C:\Windows\leame.htm"
   
   'Check if the file exists
   If Dir(sFile) = "" Or sFile = "" Then
        MsgBox "File not found!", vbCritical
        Exit Sub
   End If
   'Create a buffer
   s2 = String(MAX_FILENAME_LEN, 32)
   'Retrieve the name and handle of the executable, associated with this file
   i = FindExecutable(sFile, vbNullString, s2)
   If i > 32 Then
      MsgBox Left$(s2, InStr(s2, Chr$(0)) - 1)
   Else
      MsgBox "No association found !"
   End If
End Sub

