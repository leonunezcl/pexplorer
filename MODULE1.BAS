Attribute VB_Name = "Module1"
Option Explicit

Public glbrutina As String

Public Function ClearEnterInString(ByVal sText As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    For k = 1 To Len(sText)
        If Chr$(Asc(Mid$(sText, k, 1))) <> Chr$(13) And Chr$(Asc(Mid$(sText, k, 1))) <> Chr$(10) Then
            ret = ret & Mid$(sText, k, 1)
        Else
            ret = ret & " "
        End If
    Next k
    
    ClearEnterInString = ret
    
End Function

