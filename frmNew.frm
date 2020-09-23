VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New Access Attempt!"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form2"
   ScaleHeight     =   4695
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ProcessGateway.UserControl3 UserControl31 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   8916
   End
   Begin ProcessGateway.UserControl2 UserControl21 
      Height          =   4695
      Left            =   495
      TabIndex        =   0
      Top             =   0
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   8281
      Begin VB.Timer tmr 
         Interval        =   1000
         Left            =   4440
         Top             =   1560
      End
      Begin ProcessGateway.UserControl7 uc7Allow 
         Height          =   360
         Left            =   240
         TabIndex        =   5
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         Hold_Caption    =   "Allow"
      End
      Begin ProcessGateway.UserControl7 uc7Block 
         Height          =   360
         Left            =   3720
         TabIndex        =   4
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         Hold_Caption    =   "Block"
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   4800
         TabIndex        =   22
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Process Will be blocked in:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblListWin 
         BackStyle       =   0  'Transparent
         Caption         =   "List Windows"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3720
         MousePointer    =   13  'Arrow and Hourglass
         TabIndex        =   20
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblLastAction 
         BackStyle       =   0  'Transparent
         Caption         =   "[Allowed/Blocked]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2520
         TabIndex        =   19
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label lblAttempts 
         BackStyle       =   0  'Transparent
         Caption         =   "[Attempts]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblFilename 
         BackStyle       =   0  'Transparent
         Caption         =   "[File Name]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2520
         TabIndex        =   17
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lblChild 
         BackStyle       =   0  'Transparent
         Caption         =   "[Children]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblParent 
         BackStyle       =   0  'Transparent
         Caption         =   "[Parent PID]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Action Taken:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Access Attempts:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Child Windows:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent PID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblPID 
         BackStyle       =   0  'Transparent
         Caption         =   "[PID]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblProcName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[Process Name]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   4560
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ALERT!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblMain 
         BackStyle       =   0  'Transparent
         Caption         =   "New Access Attempt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Access Attempts:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tPID As Long
Dim tPName As String
Dim tempBypass As Boolean

Public Sub popForm(Index As Integer)
    If Form1.vChk3(3).Checked = True Then
        lblCount.Caption = "N/A"
    End If
    tPID = procinfo(Index).th32ProcessID
    tPName = procinfo(Index).procName
    lblProcName.Caption = procinfo(Index).procName
    lblpID.Caption = procinfo(Index).th32ProcessID
    lblParent.Caption = procinfo(Index).th32ParentProcessID
    lblChild.Caption = procinfo(Index).childWnd
    lblFilename.Caption = procinfo(Index).szexeFile
    Call checkforJailed(procinfo(Index).procName)
End Sub

Private Sub Form_Load()
    uc7Allow.SubClassMe
    uc7Block.SubClassMe
    unloadOK = False
'-------------------------------------------
    If Form1.vChk3(1).Checked = True Then
        tempAccPass = True
        Call uc7Allow_Clicked
    ElseIf Form1.vChk3(2).Checked = True Then
        tempBypass = True
        Call uc7Block_Clicked
    End If
'-------------------------------------------
    Call popForm(frmIndex)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If unloadOK = False Then
        Cancel = 1
        Exit Sub
    End If
    uc7Block.UnSubClassMe
    uc7Allow.UnSubClassMe
End Sub

Public Sub checkforJailed(fName As String)
Dim found As Integer
    found = 0
    For i = 0 To UBound(jailInfo)
        If UCase(fName) = UCase(jailInfo(i).exeName) Then
           found = 1
           lblAttempts.Caption = jailInfo(i).attempts
           lblLastAction.Caption = jailInfo(i).prevAction
           Exit For
        End If
    Next i
    If found = 0 Then
        lblAttempts.Caption = 0
        lblLastAction.Caption = "N/A"
    End If
End Sub

Private Sub lblListWin_Click()
    winID = Int(val(lblpID.Caption))
    winName = lblProcName.Caption
    ReDim wndNames(99) As String
    EnumWindows AddressOf EnumWindowsProc2, ByVal 0&
    Load frmWindows
    frmWindows.Show
End Sub

Private Sub tmr_Timer()
    If Form1.vChk3(3).Checked = True Then
        tmr.Enabled = False
        Exit Sub
    End If
    lblCount.Caption = lblCount.Caption - 1
    If lblCount.Caption = "0" Then
        tempBypass = True
        Call uc7Block_Clicked
    End If
End Sub

Private Sub uc7Allow_Clicked()
Dim found As Integer
    If tempAccPass = False Then
        If protectAccess = True Then
            Dim tempPass As String
            tempPass = InputBox("Password set for Access Control. Enter password to continue", "Enter Password")
            If tempPass <> protectPass Then
                MsgBox "Invalid password", vbCritical, "Invalid Password"
                Exit Sub
            End If
        End If
    End If
    unloadOK = True
    found = 0
    ResumeThreads tPID
    If taskmgrFrozen = True Then
        taskmgrFrozen = False
        'Form1.ProcessHack.Timer3_Timer
    End If
    For i = 0 To UBound(jailInfo)
        If UCase(jailInfo(i).exeName) = UCase(tPName) Then
            jailInfo(i).attempts = jailInfo(i).attempts + 1
            jailInfo(i).prevAction = "Allowed"
            jailInfo(i).lastTime = time()
            jailInfo(i).onNow = True
            jailInfo(i).jailPID = tPID
            ReDim Preserve jailInfo(i).attemptTimes(jailInfo(i).attempts)
            jailInfo(i).attemptTimes(jailInfo(i).attempts) = time()
            found = 1
            Call addLog(jailInfo(i))
            Exit For
        End If
    Next i
    
    If found = 0 Then
        If jailInfo(UBound(jailInfo)).exeName = "" Then
            jailInfo(UBound(jailInfo)).exeName = tPName
            jailInfo(UBound(jailInfo)).attempts = 1
            jailInfo(UBound(jailInfo)).prevAction = "Allowed"
            jailInfo(UBound(jailInfo)).firstTime = time()
            jailInfo(UBound(jailInfo)).lastTime = time()
            jailInfo(UBound(jailInfo)).dateOf = Date
            jailInfo(UBound(jailInfo)).onNow = True
            jailInfo(UBound(jailInfo)).jailPID = tPID
            ReDim Preserve jailInfo(UBound(jailInfo)).attemptTimes(jailInfo(UBound(jailInfo)).attempts)
            jailInfo(UBound(jailInfo)).attemptTimes(jailInfo(UBound(jailInfo)).attempts) = time()
            Call addLog(jailInfo(UBound(jailInfo)))
        Else
            ReDim Preserve jailInfo(UBound(jailInfo) + 1)
            jailInfo(UBound(jailInfo)).exeName = tPName
            jailInfo(UBound(jailInfo)).attempts = 1
            jailInfo(UBound(jailInfo)).prevAction = "Allowed"
            jailInfo(UBound(jailInfo)).firstTime = time()
            jailInfo(UBound(jailInfo)).lastTime = time()
            jailInfo(UBound(jailInfo)).dateOf = Date
            jailInfo(UBound(jailInfo)).onNow = True
            jailInfo(UBound(jailInfo)).jailPID = tPID
            ReDim Preserve jailInfo(UBound(jailInfo)).attemptTimes(jailInfo(UBound(jailInfo)).attempts)
            jailInfo(UBound(jailInfo)).attemptTimes(jailInfo(UBound(jailInfo)).attempts) = time()
            Call addLog(jailInfo(UBound(jailInfo)))
        End If
    End If
    
    If tempAccPass = True Then
        tempAccPass = False
    End If
    
    Unload Me
End Sub

Private Sub uc7Block_Clicked()
Dim found As Integer
    If tempBypass = False Then
        If protectAccess = True Then
            Dim tempPass As String
            tempPass = InputBox("Password set for Access Control. Enter password to continue", "Enter Password")
            If tempPass <> protectPass Then
                MsgBox "Invalid password", vbCritical, "Invalid Password"
                Exit Sub
            End If
        End If
    End If
    unloadOK = True
    found = 0
    For i = 0 To UBound(jailInfo)
        If UCase(jailInfo(i).exeName) = UCase(tPName) Then
            jailInfo(i).attempts = jailInfo(i).attempts + 1
            jailInfo(i).prevAction = "Blocked"
            jailInfo(i).lastTime = time()
            jailInfo(i).onNow = True
            jailInfo(i).jailPID = tPID
            ReDim Preserve jailInfo(i).attemptTimes(jailInfo(i).attempts)
            jailInfo(i).attemptTimes(jailInfo(i).attempts) = time()
            found = 1
            Call addLog(jailInfo(i))
            Exit For
        End If
    Next i
    
    If found = 0 Then
        If jailInfo(UBound(jailInfo)).exeName = "" Then
            jailInfo(UBound(jailInfo)).exeName = tPName
            jailInfo(UBound(jailInfo)).attempts = 1
            jailInfo(UBound(jailInfo)).prevAction = "Blocked"
            jailInfo(UBound(jailInfo)).firstTime = time()
            jailInfo(UBound(jailInfo)).lastTime = time()
            jailInfo(UBound(jailInfo)).dateOf = Date
            jailInfo(UBound(jailInfo)).onNow = True
            jailInfo(UBound(jailInfo)).jailPID = tPID
            ReDim Preserve jailInfo(UBound(jailInfo)).attemptTimes(jailInfo(UBound(jailInfo)).attempts)
            jailInfo(UBound(jailInfo)).attemptTimes(jailInfo(UBound(jailInfo)).attempts) = time()
            Call addLog(jailInfo(UBound(jailInfo)))
        Else
            ReDim Preserve jailInfo(UBound(jailInfo) + 1)
            jailInfo(UBound(jailInfo)).exeName = tPName
            jailInfo(UBound(jailInfo)).attempts = 1
            jailInfo(UBound(jailInfo)).prevAction = "Blocked"
            jailInfo(UBound(jailInfo)).firstTime = time()
            jailInfo(UBound(jailInfo)).lastTime = time()
            jailInfo(UBound(jailInfo)).dateOf = Date
            jailInfo(UBound(jailInfo)).onNow = True
            jailInfo(UBound(jailInfo)).jailPID = tPID
            ReDim Preserve jailInfo(UBound(jailInfo)).attemptTimes(jailInfo(UBound(jailInfo)).attempts)
            jailInfo(UBound(jailInfo)).attemptTimes(jailInfo(UBound(jailInfo)).attempts) = time()
            Call addLog(jailInfo(UBound(jailInfo)))
        End If
    End If
    Call refreshJail
    tempBypass = False
    Unload Me
End Sub
