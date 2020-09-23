Attribute VB_Name = "mSubClass"
Option Explicit

'This is only a very simple possibility to subclass as form.

Public Declare Function CallWindowProcA Lib "user32" ( _
    ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Public oldProc As Long

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Static VPressed As Boolean
    WndProc = 0
    If uMsg = WM_HOTKEY Then
        If Form1.vChk1(6).Checked = True Then
            If hotkeyPrompt = True Then
                Dim tempRep As String
                tempRep = InputBox("Enter Password to Continue", "Enter Password")
                If tempRep <> protectPass Then
                    MsgBox "Invalid password", vbCritical, "Invalid Password"
                Else
                    Form1.WindowState = 0
                    Form1.Show
                    Form1.vChk1_ButtonClicked 0, True
                End If
            End If
        End If
    Else
        WndProc = CallWindowProcA(oldProc, hwnd, uMsg, wParam, lParam)
    End If
End Function

