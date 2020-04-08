Attribute VB_Name = "MSubclass"
Option Explicit

' declares:
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const GWL_WNDPROC = (-4)

' SubTimer is independent of VBCore, so it hard codes error handling

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private m_iCurrentMessage As Long
Private m_iProcOld As Long

Public Property Get CurrentMessage() As Long
   CurrentMessage = m_iCurrentMessage
End Property

Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
        Case eeCantSubclass
            sText = "Can't subclass window"
        Case eeAlreadyAttached
            sText = "Message already handled by another class"
        Case eeInvalidWindow
            sText = "Invalid window"
        Case eeNoExternalWindow
            sText = "Can't modify external window"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
End Sub

Sub AttachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    Dim procOld As Long, f As Long, c As Long
    Dim iC As Long, bFail As Boolean
    
    ' Validate window
    If IsWindow(hWnd) = False Then ErrRaise eeInvalidWindow
    If IsWindowLocal(hWnd) = False Then ErrRaise eeNoExternalWindow

    ' Get the message count
    c = GetProp(hWnd, "C" & hWnd)
    If c = 0 Then
        ' Subclass window by installing window procecure
        procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
        If procOld = 0 Then ErrRaise eeCantSubclass
        ' Associate old procedure with handle
        f = SetProp(hWnd, hWnd, procOld)
        'Debug.Assert f <> 0
        ' Count this message
        c = 1
        f = SetProp(hWnd, "C" & hWnd, c)
    Else
        ' Count this message
        c = c + 1
        f = SetProp(hWnd, "C" & hWnd, c)
    End If
    'Debug.Assert f <> 0
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
    c = GetProp(hWnd, hWnd & "#" & iMsg & "C")
    If (c > 0) Then
        For iC = 1 To c
            If (GetProp(hWnd, hWnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
                ErrRaise eeAlreadyAttached
                bFail = True
                Exit For
            End If
        Next iC
    End If
                
    If Not (bFail) Then
        c = c + 1
        ' Increase count for hWnd/Msg:
        f = SetProp(hWnd, hWnd & "#" & iMsg & "C", c)
        'Debug.Assert f <> 0
        
        ' Associate object with message at the count:
        f = SetProp(hWnd, hWnd & "#" & iMsg & "#" & c, ObjPtr(iwp))
        'Debug.Assert f <> 0
    End If
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    Dim procOld As Long, f As Long, c As Long
    Dim iC As Long, iP As Long, lPtr As Long
    
    ' Get the message count
    c = GetProp(hWnd, "C" & hWnd)
    If c = 1 Then
        ' This is the last message, so unsubclass
        procOld = GetProp(hWnd, hWnd)
        'Debug.Assert procOld <> 0
        ' Unsubclass by reassigning old window procedure
        Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
        ' Remove unneeded handle (oldProc)
        RemoveProp hWnd, hWnd
        ' Remove unneeded count
        RemoveProp hWnd, "C" & hWnd
    Else
        ' Uncount this message
        c = GetProp(hWnd, "C" & hWnd)
        c = c - 1
        f = SetProp(hWnd, "C" & hWnd, c)
    End If
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
    
    ' How many instances attached to this hwnd/msg?
    c = GetProp(hWnd, hWnd & "#" & iMsg & "C")
    If (c > 0) Then
        ' Find this iwp object amongst the items:
        For iC = 1 To c
            If (GetProp(hWnd, hWnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
                iP = iC
                Exit For
            End If
        Next iC
    
        If (iP <> 0) Then
             ' Remove this item:
             For iC = iP + 1 To c
                lPtr = GetProp(hWnd, hWnd & "#" & iMsg & "#" & iC)
                SetProp hWnd, hWnd & "#" & iMsg & "#" & (iC - 1), lPtr
             Next iC
        End If
        ' Decrement the count
        RemoveProp hWnd, hWnd & "#" & iMsg & "#" & c
        c = c - 1
        SetProp hWnd, hWnd & "#" & iMsg & "C", c
    
    End If
End Sub

Private Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long) _
                            As Long
    Dim procOld As Long, pSubclass As Long, f As Long
    Dim iwp As ISubclass, iwpT As ISubclass
    Dim iPC As Long, iP As Long, bNoProcess As Long
    Dim bCalled As Boolean
    
    ' Get the old procedure from the window
    procOld = GetProp(hWnd, hWnd)
    'Debug.Assert procOld <> 0
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
    
    ' Get the number of instances for this msg/hwnd:
    bCalled = False
    iPC = GetProp(hWnd, hWnd & "#" & iMsg & "C")
    If (iPC > 0) Then
        ' For each instance attached to this msg/hwnd, call the subclass:
        For iP = 1 To iPC
            bNoProcess = False
            ' Get the object pointer from the message
            pSubclass = GetProp(hWnd, hWnd & "#" & iMsg & "#" & iP)
            If pSubclass = 0 Then
                ' This message not handled, so pass on to old procedure
                WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                            wParam, ByVal lParam)
                bNoProcess = True
            End If
            
            If Not (bNoProcess) Then
                ' Turn the pointer into an illegal, uncounted interface
                CopyMemory iwpT, pSubclass, 4
                ' Do NOT hit the End button here! You will crash!
                ' Assign to legal reference
                Set iwp = iwpT
                ' Still do NOT hit the End button here! You will still crash!
                ' Destroy the illegal reference
                CopyMemory iwpT, 0&, 4
                ' OK, hit the End button if you must--you'll probably still crash,
                ' but it will be because of the subclass, not the uncounted reference
                
                ' Store the current message, so the client can check it:
                m_iCurrentMessage = iMsg
                m_iProcOld = procOld
                
                ' Use the interface to call back to the class
                With iwp
                    ' Preprocess (only check this the first time around):
                    If (iP = 1) Then
                        If .MsgResponse = emrPreprocess Then
                           If Not (bCalled) Then
                              WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                                        wParam, ByVal lParam)
                              bCalled = True
                           End If
                        End If
                    End If
                    ' Consume (this message is always passed to all control
                    ' instances regardless of whether any single one of them
                    ' requests to consume it):
                    WindowProc = .WindowProc(hWnd, iMsg, wParam, ByVal lParam)
                    ' PostProcess (only check this the last time around):
                    If (iP = iPC) Then
                        If .MsgResponse = emrPostProcess Then
                           If Not (bCalled) Then
                              WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                                        wParam, ByVal lParam)
                              bCalled = True
                           End If
                        End If
                    End If
                End With
            End If
        Next iP
    Else
        ' This message not handled, so pass on to old procedure
        WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                    wParam, ByVal lParam)
    End If
End Function
Public Function CallOldWindowProc( _
      ByVal hWnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
   CallOldWindowProc = CallWindowProc(m_iProcOld, hWnd, iMsg, wParam, lParam)

End Function

' Cheat! Cut and paste from MWinTool rather than reusing
' file because reusing file would cause many unneeded dependencies
Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function
'



