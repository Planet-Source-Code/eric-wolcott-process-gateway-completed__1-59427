Attribute VB_Name = "SubClass"
Option Explicit

Public colTrackMouse As New Collection
Public Declare Function CallWindowProcA Lib "user32" ( _
    ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Function procTrackMouse(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Dim tmItem As New CTrackMouse
    Set tmItem = colTrackMouse.Item("TM" & hwnd)
    If Not (tmItem Is Nothing) Then procTrackMouse = tmItem.MessageReceived(wMsg, wParam, lParam)
End Function

