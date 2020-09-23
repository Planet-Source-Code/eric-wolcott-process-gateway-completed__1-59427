VERSION 5.00
Begin VB.Form frmProcessHack 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Process Hacker"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form2"
   ScaleHeight     =   5115
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin ProcessGateway.UserControl3 UserControl31 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   8705
   End
   Begin ProcessGateway.UserControl2 UserControl21 
      Height          =   5055
      Left            =   1575
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   8916
      Begin ProcessGateway.UserControl7 uc7Hack 
         Height          =   360
         Left            =   5880
         TabIndex        =   6
         Top             =   4320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         Hold_Caption    =   "Enable Process Hacker"
         Hold_Enabled    =   0   'False
      End
      Begin VB.ListBox lstHackProcTo 
         Height          =   3570
         Left            =   4920
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.ListBox lstProcto 
         Height          =   3570
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rename Hacked Process To:"
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
         Left            =   4440
         TabIndex        =   4
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label lblMain 
         BackStyle       =   0  'Transparent
         Caption         =   "Rename This Process To:"
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
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmProcessHack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appExe As String
Dim host As String
Dim hostIdent As String

Private Sub Form_Load()
    appExe = App.exeName & ".exe"
    uc7Hack.SubClassMe
    For i = 0 To UBound(procinfo)
        If procinfo(i).procName = "" Then
            Exit For
        End If
        If UCase(procinfo(i).procName) <> UCase(App.exeName & ".exe") Then
            If Len(appExe) = Len(procinfo(i).procName) Then
                lstProcto.AddItem procinfo(i).procName
            End If
            lstHackProcTo.AddItem procinfo(i).procName
        End If
    Next i
    If lstProcto.ListCount = 0 Then
        MsgBox ".Exe Name does not match the length of any processes. Shorten or increase the exe name " & appExe, vbCritical, "Error"
        Me.Hide
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    uc7Hack.UnSubClassMe
End Sub

Private Sub lstHackProcTo_Click()
    uc7Hack.Enabled = True
End Sub

Private Sub uc7Hack_Clicked()
    App.Title = ""
    
    host = lstProcto.List(lstProcto.ListIndex)
    hostIdent = lstHackProcTo.List(lstHackProcTo.ListIndex)
    
    If Len(host) > Len(hostIdent) Then
        hostIdent = hostIdent & Space(Len(host) - Len(hostIdent))
    End If
    
    
    Form1.ProcessHack.i_HostProcess = host
    Form1.ProcessHack.i_NewHostIdentity = hostIdent
    Form1.ProcessHack.i_NewHostProcess = appExe
    Form1.ProcessHack.Enable
    Form1.lblHackStat.Caption = "[ENABLED]"
    Form1.lblHackStat.ForeColor = &H8000&
    Form1.lblAppName.Caption = appExe
    Form1.lblHostProc.Caption = host
    Form1.lblIdent.Caption = hostIdent
    Me.Hide
End Sub
