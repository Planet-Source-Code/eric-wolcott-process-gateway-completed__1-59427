Attribute VB_Name = "mHotKey"
Option Explicit

Private Declare Function RegisterHotKey Lib "user32" _
    (ByVal hwnd As Long, ByVal id As Long, _
    ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" _
    (ByVal hwnd As Long, ByVal id As Long) As Long

Public Const MOD_ALT = &H1
Public Const mod_control = &H2
Public Const mod_shift = &H4
Private m_hkCount As Long

Function HotKeyActivate(ByVal hwnd As Long, _
    Modifier As Integer, Optional KeyCode As Integer) As Long
    
    m_hkCount = m_hkCount + 1
    
    HotKeyActivate = RegisterHotKey(hwnd, m_hkCount, mod_control + mod_shift, KeyCode)
    keyOn = 1
End Function

Function HotKeyDeactivate(ByVal hwnd As Long)
    Dim i As Integer
    For i = 1 To m_hkCount
        UnregisterHotKey hwnd, i
    Next i
    m_hkCount = 0
End Function
