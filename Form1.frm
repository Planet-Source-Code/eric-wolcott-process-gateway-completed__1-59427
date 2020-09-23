VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00D6D1D0&
   Caption         =   "Process Gateway"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ProcessGateway.ProcessHack ProcessHack 
      Height          =   135
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
   End
   Begin ProcessGateway.UserControl4 UserControl42 
      Height          =   1065
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1879
   End
   Begin ProcessGateway.UserControl1 UserControl12 
      Height          =   375
      Left            =   15
      TabIndex        =   5
      Top             =   2745
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   661
      strCaption      =   "Options"
   End
   Begin ProcessGateway.UserControl4 UserControl41 
      Height          =   1500
      Left            =   45
      TabIndex        =   4
      Top             =   975
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   2646
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   9300
      TabIndex        =   3
      Top             =   5550
      Width           =   9300
   End
   Begin ProcessGateway.UserControl3 UserControl31 
      Height          =   2565
      Left            =   0
      TabIndex        =   2
      Top             =   2985
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4524
   End
   Begin ProcessGateway.UserControl2 UserControl21 
      Height          =   5430
      Left            =   2490
      TabIndex        =   1
      Top             =   555
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   9578
      Begin VB.Frame frmMain2 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   4
         Left            =   2160
         TabIndex        =   69
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Frame frmMain2 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   3
         Left            =   1920
         TabIndex        =   68
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         Begin ProcessGateway.CheckBox vChk3 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   104
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "Automatically Allow"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk3 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   105
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "Automatically Block"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk3 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   106
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   450
            Caption         =   "Disable Auto Block Timer (NOT RECCOMENDED)"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk3 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   107
            Top             =   2280
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            Caption         =   "Log New Access Attempts Only"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   108
            Top             =   1920
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            Caption         =   "Disable Logging"
            Bold            =   0   'False
         End
      End
      Begin VB.Frame frmMain2 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   2
         Left            =   1680
         TabIndex        =   67
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         Begin ProcessGateway.CheckBox vChk1 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   97
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            Caption         =   "Do not show Process Gateway in TaskBar"
            Bold            =   0   'False
         End
         Begin VB.Timer tmrFindRun 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   4320
            Top             =   1560
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Advanced Process Hacker"
            Height          =   1695
            Left            =   120
            TabIndex        =   74
            Top             =   2160
            Width           =   6015
            Begin ProcessGateway.CheckBox vChk1 
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   103
               Top             =   1320
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   450
               Caption         =   "Disguise Process Name in Ctrl+Alt+Del"
               Bold            =   0   'False
            End
            Begin VB.Label lblIdent 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4440
               MousePointer    =   10  'Up Arrow
               TabIndex        =   82
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblHostProc 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4680
               MousePointer    =   10  'Up Arrow
               TabIndex        =   81
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblAppName 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1560
               MousePointer    =   10  'Up Arrow
               TabIndex        =   80
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label lblHackStat 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "[DISABLED]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   720
               MousePointer    =   10  'Up Arrow
               TabIndex        =   79
               Top             =   360
               Width           =   1815
            End
            Begin VB.Line Line8 
               X1              =   0
               X2              =   6000
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Label Label23 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Application Name:"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               MousePointer    =   10  'Up Arrow
               TabIndex        =   78
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label22 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Status:"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               MousePointer    =   10  'Up Arrow
               TabIndex        =   77
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label21 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Host Renamed To:"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3000
               MousePointer    =   10  'Up Arrow
               TabIndex        =   76
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label20 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Current Host Process:"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3000
               MousePointer    =   10  'Up Arrow
               TabIndex        =   75
               Top             =   360
               Width           =   1575
            End
         End
         Begin ProcessGateway.CheckBox vChk1 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   98
            Top             =   600
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   450
            Caption         =   "Do not show Process Gateway in System Tray"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk1 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   99
            Top             =   960
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   450
            Caption         =   "Hide Process From Ctrl+Alt+Del (Windows 9x Only)"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk1 
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   100
            Top             =   1320
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            Caption         =   "Hide Process Gateway When minimized"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk1 
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   101
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "Disable RUN Menu"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk1 
            Height          =   255
            Index           =   5
            Left            =   4320
            TabIndex        =   102
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "Disable Ctrl+Alt+Del"
            Bold            =   0   'False
         End
      End
      Begin VB.Frame frmMain2 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   1
         Left            =   1440
         TabIndex        =   66
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            Caption         =   "Run on Startup"
            Bold            =   0   'False
         End
         Begin ProcessGateway.UserControl7 uc7Password 
            Height          =   360
            Left            =   4080
            TabIndex        =   72
            Top             =   3360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            Hold_Caption    =   "Apply"
            Hold_Enabled    =   0   'False
         End
         Begin VB.TextBox txtPass 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3360
            PasswordChar    =   "*"
            TabIndex        =   70
            Top             =   2880
            Width           =   2175
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   89
            Top             =   960
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            Caption         =   "Enable HotKey Password"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   90
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            Caption         =   "Password Protect Options"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            Caption         =   "Retain Options Upon Exit"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   92
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   450
            Caption         =   "Password Protect Access Attempt "
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   93
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            Caption         =   "Password Protect Logs"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   7
            Left            =   2760
            TabIndex        =   94
            Top             =   1320
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            Caption         =   "Password Protect Process Information"
            Bold            =   0   'False
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   6
            Left            =   2760
            TabIndex        =   95
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            Caption         =   "Password Protect All"
         End
         Begin ProcessGateway.CheckBox vChk 
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   96
            Top             =   2400
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            Caption         =   "Change Password"
         End
         Begin VB.Line Line6 
            X1              =   2400
            X2              =   2400
            Y1              =   120
            Y2              =   3960
         End
         Begin VB.Line Line5 
            X1              =   2400
            X2              =   6240
            Y1              =   2160
            Y2              =   2160
         End
      End
      Begin VB.Frame frmMain 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   5
         Left            =   1200
         TabIndex        =   65
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         Begin ProcessGateway.UserControl7 uc7DeleteLog 
            Height          =   360
            Left            =   4200
            TabIndex        =   86
            Top             =   3000
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            Hold_Caption    =   "Delete Selected Entry"
         End
         Begin ProcessGateway.UserControl7 uc7SaveLog 
            Height          =   360
            Left            =   120
            TabIndex        =   84
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            Hold_Caption    =   "Save Log"
         End
         Begin MSComctlLib.ListView lstVwLog 
            Height          =   2295
            Left            =   120
            TabIndex        =   83
            Top             =   600
            Width           =   6005
            _ExtentX        =   10583
            _ExtentY        =   4048
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ColHdrIcons     =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   2640
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":2A8B2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":2D064
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin ProcessGateway.UserControl7 uc7LoadLog 
            Height          =   360
            Left            =   120
            TabIndex        =   85
            Top             =   3480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   635
            Hold_Caption    =   "Load Log"
         End
         Begin ProcessGateway.UserControl7 uc7MoreInfo 
            Height          =   360
            Left            =   4560
            TabIndex        =   87
            Top             =   3480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            Hold_Caption    =   "Get More Info"
         End
         Begin VB.Label lblLogStatus 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "[Logging Enabled]"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   405
            Left            =   1560
            TabIndex        =   71
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.Frame frmMain 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         Begin ProcessGateway.UserControl7 uc7MoreInfoRun 
            Height          =   360
            Left            =   2640
            TabIndex        =   42
            Top             =   3420
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            Hold_Caption    =   "More Information"
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Process Refresh"
            Height          =   2655
            Left            =   4440
            TabIndex        =   10
            Top             =   270
            Width           =   1695
            Begin VB.TextBox txtProcRef 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   600
               TabIndex        =   12
               Text            =   "0"
               Top             =   360
               Width           =   495
            End
            Begin ProcessGateway.UserControl7 uc7Interval 
               Height          =   360
               Left            =   240
               TabIndex        =   11
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   635
               Hold_Caption    =   "Set Interval"
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "Sets the rate in seconds at which processes with be refreshed. Leave it blank to stop. 0 for Default."
               Height          =   1215
               Left            =   120
               TabIndex        =   13
               Top             =   1200
               Width           =   1455
            End
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Height          =   2955
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   4200
         End
         Begin ProcessGateway.UserControl7 uc7Suspend 
            Height          =   360
            Left            =   4440
            TabIndex        =   14
            Top             =   3420
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            Hold_Caption    =   "Suspend Thread"
         End
         Begin ProcessGateway.UserControl7 uc7Resume 
            Height          =   360
            Left            =   4440
            TabIndex        =   15
            Top             =   2970
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            Hold_Caption    =   "Resume Thread"
         End
      End
      Begin VB.Frame frmMain 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   4
         Left            =   720
         TabIndex        =   33
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         Begin ProcessGateway.UserControl7 uc7ProcWin 
            Height          =   360
            Left            =   3960
            TabIndex        =   50
            Top             =   3000
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            Hold_Caption    =   "List Accociated Windows"
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4800
            TabIndex        =   64
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblThreads 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4800
            TabIndex        =   63
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Process Priority:"
            Height          =   255
            Left            =   3360
            TabIndex        =   62
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Threads:"
            Height          =   255
            Left            =   3480
            TabIndex        =   61
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblChd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   60
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Child Windows:"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label lblFirstAccess 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   58
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "First Access Date:"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblAccessAttempts 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   46
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lbllastAccess 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   45
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblParentID 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   44
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblProcIDs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   43
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblLastAct 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   41
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Last Access Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Access Attempts:"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Previous Action Taken:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Parent PID:"
            Height          =   255
            Left            =   480
            TabIndex        =   37
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Process ID:"
            Height          =   255
            Left            =   480
            TabIndex        =   36
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblProcID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   35
            Top             =   240
            Width           =   2055
         End
         Begin VB.Line Line4 
            X1              =   240
            X2              =   5760
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lblProcName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame frmMain 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   3
         Left            =   480
         TabIndex        =   25
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         Begin ProcessGateway.UserControl7 uc7MoreInfoJail 
            Height          =   360
            Left            =   3840
            TabIndex        =   32
            Top             =   2760
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            Hold_Caption    =   "More Information"
         End
         Begin ProcessGateway.UserControl7 uc7Unblock 
            Height          =   360
            Left            =   3840
            TabIndex        =   31
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            Hold_Caption    =   "Unblock Process"
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Process Information"
            Height          =   1815
            Left            =   3720
            TabIndex        =   27
            Top             =   270
            Width           =   2415
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "Last Access:"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label lbllTime 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   1320
               TabIndex        =   55
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblPrevAttempts 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   1320
               TabIndex        =   54
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label lblTime 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   1320
               TabIndex        =   53
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lbljailPID 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   1320
               TabIndex        =   52
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "Prev. Attempts:"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "First Access:"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "Process ID:"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.ListBox List2 
            Height          =   3180
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame frmMain 
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   4320
         Width           =   6255
         Begin ProcessGateway.UserControl7 uc7Monitor 
            Height          =   360
            Left            =   3840
            TabIndex        =   47
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            Hold_Caption    =   "Monitor Processes"
         End
         Begin ProcessGateway.UserControl7 uc7ProcJailed 
            Height          =   360
            Left            =   3840
            TabIndex        =   18
            Top             =   2040
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            Hold_Caption    =   "View Processes"
         End
         Begin ProcessGateway.UserControl7 uc7ProcRun 
            Height          =   360
            Left            =   3840
            TabIndex        =   17
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            Hold_Caption    =   "View Processes"
         End
         Begin VB.Label lblRefresh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Refresh Processes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1680
            MousePointer    =   10  'Up Arrow
            TabIndex        =   51
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblMonitor 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   5160
            TabIndex        =   49
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Status:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblProcJailed 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Jailed Processes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   23
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   5400
            Y1              =   1880
            Y2              =   1880
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Jailed Processes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label lblProcRun 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   975
         End
         Begin VB.Label label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Running Processes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   20
            Top             =   840
            Width           =   2055
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   5400
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Running Processes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Timer tmrProcRef 
         Enabled         =   0   'False
         Left            =   6360
         Top             =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   375
         X2              =   5925
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label lblMain 
         BackStyle       =   0  'Transparent
         Caption         =   "Process Home"
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
         Height          =   645
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   6150
      End
   End
   Begin ProcessGateway.UserControl1 UserControl11 
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   600
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   661
      strCaption      =   "Main Menu"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Load tray icon
    Call loadIcon
    'Initalize Log Listview
    Call rdyLog
    'Load options from previous run if needed
    Call loadSett
    'Set inital values for certin variables
    noList = False
    tmrON = False
    keyOn = 0
    monitorOn = False
    firstRun = True
    refProc = True
    unloadOK = False
    logOn = True
    logNew = False
    protectOpt = False
    protectAccess = False
    showGo = False
    hotkeyPrompt = False
    taskmgrFrozen = False
    tempAccPass = False
    protectPass = ""
    prevIndex = 1
    prevCapt = "Process Home"
    Me.BackColor = 14078416
    'Add side buttons
    UserControl11.SubClassMe
    UserControl41.AddButton "Process Home"
    UserControl41.AddButton "Process Control Center"
    UserControl41.AddButton "Process Info Center"
    UserControl41.AddButton "View Process Jail"
    UserControl41.AddButton "Process Logs"
    UserControl12.SubClassMe
    UserControl42.AddButton "General Options"
    UserControl42.AddButton "Stealth Options"
    UserControl42.AddButton "Access/Log Options"
    'UserControl42.AddButton "Logging Options"
    UserControl12.Top = UserControl11.Top + UserControl11.Height + 5
    'Subclass command buttons
    uc7Resume.SubClassMe
    uc7Suspend.SubClassMe
    uc7Interval.SubClassMe
    uc7ProcRun.SubClassMe
    uc7LoadLog.SubClassMe
    uc7SaveLog.SubClassMe
    uc7ProcJailed.SubClassMe
    uc7MoreInfo.SubClassMe
    uc7Unblock.SubClassMe
    uc7DeleteLog.SubClassMe
    uc7MoreInfoJail.SubClassMe
    uc7MoreInfoRun.SubClassMe
    uc7Monitor.SubClassMe
    uc7ProcWin.SubClassMe
    uc7Password.SubClassMe
    'Arranges frames on form and bring up main frame.
    Call arrangeFrm
    'Dim arrarys
    ReDim procinfo(150) As PROCESSENTRY32
    ReDim jailInfo(1) As jailedProc
    'Take snapshot of inital processes.
    Call enumProc
    firstRun = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim rep As String
    If List2.ListCount > 0 Then
        rep = MsgBox("You have " & List2.ListCount & " Jailed processes. Do you want to resume these processes?", vbYesNoCancel, "Exit")
        If rep = vbYes Then
          For i = 0 To UBound(jailInfo)
            For z = 0 To List2.ListCount
                If UCase(List2.List(z)) = UCase(jailInfo(i).exeName) Then
                    ResumeThreads jailInfo(i).jailPID
                End If
            Next z
          Next i
        ElseIf rep = vbCancel Then
            Cancel = 1
            Exit Sub
        End If
    End If
    bCancel = True
    UserControl11.UnSubClassMe
    UserControl41.UnSubClass
    UserControl12.UnSubClassMe
    UserControl42.UnSubClass
    uc7Resume.UnSubClassMe
    uc7Suspend.UnSubClassMe
    uc7Interval.UnSubClassMe
    uc7ProcRun.UnSubClassMe
    uc7ProcJailed.UnSubClassMe
    uc7MoreInfo.UnSubClassMe
    uc7Unblock.UnSubClassMe
    uc7MoreInfoJail.UnSubClassMe
    uc7MoreInfoRun.UnSubClassMe
    uc7Monitor.UnSubClassMe
    uc7LoadLog.UnSubClassMe
    uc7DeleteLog.UnSubClassMe
    uc7ProcWin.UnSubClassMe
    uc7SaveLog.UnSubClassMe
    uc7Password.UnSubClassMe
    DoEvents
    If vChk(8).Checked = True Then
        Call setSett
    End If
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = Form1.hwnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
    End
End Sub

Private Sub Form_Resize()
    If vChk1(6).Checked = True Then
        If Form1.WindowState = 1 Then
            Me.Hide
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HotKeyDeactivate Me.hwnd
    SetWindowLongA Me.hwnd, GWL_WNDPROC, oldProc
End Sub

Private Sub lblRefresh_Click()
    ReDim procinfo(150) As PROCESSENTRY32
    Call enumProc
End Sub

Private Sub List2_Click()
    For i = 1 To UBound(jailInfo)
        If UCase(List2.List(List2.ListIndex)) = UCase(jailInfo(i).exeName) Then
            lbljailPID.Caption = jailInfo(i).jailPID
            lblTime.Caption = jailInfo(i).firstTime
            lbllTime.Caption = jailInfo(i).lastTime
            lblPrevAttempts.Caption = jailInfo(i).attempts
        End If
    Next i
End Sub

Private Sub tmrFindRun_Timer()
    Const WM_CLOSE = &H10
    Dim winHwnd As Long
    Dim RetVal As Long
    Dim cName As String
    cName = Space(255)
    winHwnd = FindWindow(vbNullString, "Run")
    If winHwnd <> 0 Then
        GetClassName winHwnd, cName, 255
        If InStr(1, cName, "#32770") Then
            PostMessage winHwnd, WM_CLOSE, 0&, 0&
        End If
    End If
    DoEvents
    
End Sub

Private Sub tmrProcRef_Timer()
ReDim procinfo(150) As PROCESSENTRY32
    If refProc = True Then
        List1.Clear
        Call enumProc
        refProc = False
    Else
        Call enumProc
    End If
End Sub

Private Sub uc7DeleteLog_Clicked()
Dim isChecked As Boolean
    isChecked = False
    For i = 1 To lstVwLog.ListItems.Count
        If lstVwLog.ListItems(i).Selected = True Then
            isChecked = True
            Exit For
        End If
    Next i
    If isChecked = False Then
        MsgBox "No item selected", vbCritical, "Error"
        Exit Sub
    End If
    lstVwLog.SetFocus
    If lstVwLog.SelectedItem.Index = 0 Then
        MsgBox "No item selected", vbCritical, "Error"
    Else
       lstVwLog.ListItems.Remove (lstVwLog.SelectedItem.Index)
    End If
End Sub

Private Sub uc7Interval_Clicked()
    If txtProcRef = "" Then
        tmrON = True
        Call setMonitor
    ElseIf Trim(txtProcRef) = 0 Then
        tmrProcRef.Interval = 250
    Else
        tmrProcRef.Interval = txtProcRef.Text * 1000
    End If
End Sub

Private Sub uc7LoadLog_Clicked()
Dim pgName As String
Dim pgPID As String
Dim pgTime As String
Dim pgAttempts As String
Dim pgAction As String
    lstVwLog.ListItems.Clear
    cd.Filter = "Process Gateway Logs (*.pgl)|*.pgl"
    cd.ShowOpen
    If cd.FileName = "" Then Exit Sub
    Open cd.FileName For Input As #1
    Do Until EOF(1)
        Input #1, pgName, pgPID, pgAction, pgTime, pgAttempts
        Set lstItem = Form1.lstVwLog.ListItems.Add(, , pgName)
        lstItem.SubItems(1) = pgPID
        lstItem.SubItems(2) = pgAction
        lstItem.SubItems(3) = pgTime
        lstItem.SubItems(4) = pgAttempts
        If pgAction = "Blocked" Then
            lstItem.SmallIcon = 1
        Else
            lstItem.SmallIcon = 2
        End If
    Loop
    Close #1
End Sub

Private Sub uc7Monitor_Clicked()
    Call setMonitor
End Sub

Private Sub uc7MoreInfo_Clicked()
Dim isChecked As Boolean
    isChecked = False
    For i = 1 To lstVwLog.ListItems.Count
        If lstVwLog.ListItems(i).Selected = True Then
            isChecked = True
            Exit For
        End If
    Next i
    If isChecked = False Then
        MsgBox "No item selected", vbCritical, "Error"
        Exit Sub
    End If
    Call applyInfo(Trim(lstVwLog.SelectedItem.Text))
    Call UserControl41_ButtonClick(3)
    UserControl41.forceClick (3)
    UserControl41.Width = 2460
    UserControl41.ShowMe
End Sub

Private Sub uc7MoreInfoJail_Clicked()
    If List2.List(List2.ListIndex) = "" Then
        MsgBox "No Process Selected", vbCritical, "Error"
        Exit Sub
    End If
    UserControl41.ShowMe
    Call applyInfo(List2.List(List2.ListIndex))
    Call UserControl41_ButtonClick(3)
    UserControl41.forceClick (3)
    UserControl41.Width = 2460
    UserControl41.ShowMe
End Sub

Private Sub uc7MoreInfoRun_Clicked()
    If List1.List(List1.ListIndex) = "" Then
        MsgBox "No Process Selected", vbCritical, "Error"
        Exit Sub
    End If
    UserControl41.ShowMe
    Call applyInfo(List1.List(List1.ListIndex))
    Call UserControl41_ButtonClick(3)
    UserControl41.forceClick (3)
    UserControl41.Width = 2460
    UserControl41.ShowMe
End Sub

Private Sub uc7Password_Clicked()
    If protectPass <> "" Then
        Dim tempPass As String
        tempPass = InputBox("Enter Previous Password", "Change Password")
        If tempPass <> protectPass Then
            MsgBox "Invalid password", vbCritical, "Invalid Password"
            Exit Sub
        End If
    End If
    If txtPass.Text = "" Then
        MsgBox "You must enter a valid password", vbCritical, "Password Change"
    Else
        protectPass = txtPass.Text
        txtPass.Text = ""
        vChk(2).Checked = False
        MsgBox "Password Changed", vbOKOnly, "Password Change"
    End If
End Sub

Private Sub uc7ProcJailed_Clicked()
    UserControl41.ShowMe
    Call UserControl41_ButtonClick(4)
    UserControl41.forceClick (4)
    UserControl41.Width = 2460
    UserControl41.ShowMe
End Sub

Private Sub uc7ProcRun_Clicked()
    UserControl41.ShowMe
    Call UserControl41_ButtonClick(2)
    UserControl41.forceClick (2)
    UserControl41.Width = 2460
    UserControl41.ShowMe
End Sub

Private Sub uc7ProcWin_Clicked()
    If lblProcName.Caption = "" Then
        MsgBox "No Process Selected", vbCritical, "Error"
        Exit Sub
    End If
    winID = Int(val(lblProcID.Caption))
    winName = lblProcName.Caption
    ReDim wndNames(99) As String
    EnumWindows AddressOf EnumWindowsProc2, ByVal 0&
    Load frmWindows
    frmWindows.Show
End Sub

Private Sub uc7Resume_Clicked()
Dim rep As String
    If List1.List(List1.ListIndex) = "" Then
        MsgBox "No Process Selected", vbCritical, "Error"
        Exit Sub
    End If
    For i = 0 To List2.ListCount
        If UCase(List1.List(List1.ListIndex)) = UCase(List2.List(i)) Then
            rep = MsgBox("This will remove " & List1.List(List1.ListIndex) & " from jail, Are you sure?", vbYesNo, "Unjail Processes")
            If rep = vbYes Then
                List2.RemoveItem i
            Else
                Exit Sub
            End If
        End If
    Next i
    If List1.ListIndex >= 0 Then
        ResumeThreads (procinfo(List1.ListIndex).th32ProcessID)
    End If
End Sub

Private Sub uc7SaveLog_Clicked()
Dim pgName As String
Dim pgPID As String
Dim pgTime As String
Dim pgAttempts As String
Dim pgAction As String
    cd.Filter = "Process Gateway Logs (*.pgl)|*.pgl"
    cd.ShowSave
    If cd.FileName = "" Then Exit Sub
    Open cd.FileName For Append As #1
    Close #1
    Open cd.FileName For Output As #1
    For i = 1 To lstVwLog.ListItems.Count
        pgName = lstVwLog.ListItems(i).Text
        pgPID = lstVwLog.ListItems(i).ListSubItems(1).Text
        pgAction = lstVwLog.ListItems(i).ListSubItems(2).Text
        pgTime = lstVwLog.ListItems(i).ListSubItems(3).Text
        pgAttempts = lstVwLog.ListItems(i).ListSubItems(4).Text
        Write #1, pgName, pgPID, pgAction, pgTime, pgAttempts
    Next i
    Close #1
End Sub

Private Sub uc7Suspend_Clicked()
    If List1.List(List1.ListIndex) = "" Then
        MsgBox "No Process Selected", vbCritical, "Error"
        Exit Sub
    End If
    If List1.ListIndex >= 0 Then
        SuspendThreads (procinfo(List1.ListIndex).th32ProcessID)
        glbPID = procinfo(List1.ListIndex).th32ProcessID
        EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    End If
End Sub

Private Sub uc7Unblock_Clicked()
    If List2.List(List2.ListIndex) = "" Then
        MsgBox "No Process Selected", vbCritical, "Error"
        Exit Sub
    End If
    For i = 0 To UBound(jailInfo)
        If UCase(List2.List(List2.ListIndex)) = UCase(jailInfo(i).exeName) Then
            ResumeThreads jailInfo(i).jailPID
        End If
    Next i
    List2.RemoveItem List2.ListIndex
End Sub

Private Sub UserControl11_Clicked(State As Integer)
    UserControl12.Reset
    UserControl42.Reset
    UserControl42.Visible = False
    Select Case State
        Case 0
            UserControl41.Visible = False
            UserControl12.Top = UserControl11.Top + UserControl11.Height + 5
        Case 1
            UserControl41.Left = UserControl11.Left
            UserControl41.Top = UserControl11.Top + UserControl11.Height
            UserControl41.Width = UserControl11.Width
            UserControl41.ShowMe
            UserControl41.Visible = True
            UserControl12.Top = UserControl41.Top + UserControl41.Height + 5
    End Select
End Sub

Private Sub UserControl12_Clicked(State As Integer)
    UserControl11.Reset
    UserControl41.Reset
    UserControl41.Visible = False
    UserControl12.Top = UserControl11.Top + UserControl11.Height + 5
    Select Case State
        Case 0
            UserControl42.Visible = False
        Case 1
            UserControl42.Left = UserControl12.Left
            UserControl42.Top = UserControl12.Top + UserControl12.Height
            UserControl42.Width = UserControl12.Width
            UserControl42.ShowMe
            UserControl42.Visible = True
    End Select
End Sub

Private Sub UserControl41_ButtonClick(Index As Integer)
    Select Case Index
        Case 1
            lblMain.Caption = "Process Home"
            Call showFrm(2)
            If showGo = True Then
                prevIndex = 1
                prevCapt = "Process Home"
            End If
            showGo = False
        Case 2
            'ReDim procinfo(150) As PROCESSENTRY32
            lblMain.Caption = "Process Control Center"
            Call showFrm(1)
            If showGo = True Then
                prevIndex = 2
                prevCapt = "Process Control Center"
            End If
            showGo = False
            'List1.Clear
            tempClear = 1
            'Call enumProc
        Case 3
            lblMain.Caption = "Process Information Center"
            Call showFrm(4)
            If showGo = True Then
                prevIndex = 3
                prevCapt = "Process Information Center"
            End If
            showGo = False
            'ReDim procinfo(150) As PROCESSENTRY32
            'List1.Clear
            'noList = True
            'Call enumProc
            'noList = False
            'EnumWindows AddressOf EnumWindowsProc, ByVal 0&
            'For i = 0 To arrLen - 1
            '    List1.AddItem procinfo(i).childWnd & "=" & procinfo(i).th32ProcessID
            'Next i
        Case 4
            lblMain.Caption = "Jailed Processes"
            Call showFrm(3)
            If showGo = True Then
                prevIndex = 4
                prevCapt = "Jailed Processes"
            End If
            showGo = False
        Case 5
            lblMain.Caption = "Activity Log"
            Call showFrm(5)
            If showGo = True Then
                prevIndex = 5
                prevCapt = "Activity Log"
            End If
            showGo = False
            lstVwLog.SetFocus
    End Select
End Sub

Private Sub UserControl42_ButtonClick(Index As Integer)
If protectOpt = True Then
Dim tempPass As String
    tempPass = InputBox("Password For options set. Enter password to continue", "Protected")
    If tempPass <> protectPass Then
        MsgBox "Invalid password", vbCritical, "Invalid Password"
        Exit Sub
    End If
End If
    Select Case Index
        Case 1
            lblMain.Caption = "General Options"
            Call showFrm2(1)
            prevCapt = "General Options"
        Case 2
            lblMain.Caption = "Stealth Options"
            Call showFrm2(2)
            prevCapt = "Stealth Options"
        Case 3
            lblMain.Caption = "Access Control Options"
            Call showFrm2(3)
            prevCapt = "Access Control Options"
        Case 4
            lblMain.Caption = "Logging Options"
            Call showFrm2(4)
            prevCapt = "Logging Options"
    End Select
End Sub

Private Sub arrangeFrm()
    For i = 1 To frmMain.UBound
        frmMain(i).Top = 720
        frmMain(i).Left = 360
    Next i
    For i = 1 To frmMain2.UBound
        frmMain2(i).Top = 720
        frmMain2(i).Left = 360
    Next i
End Sub

Private Sub showFrm(Index As Integer)
    If protectLogs = True Then
        If Index = 5 Then
            Dim tempPass2 As String
            tempPass2 = InputBox("Password set for Logs. Enter Password to continue", "Enter Password")
            If tempPass2 <> protectPass Then
                MsgBox "Invalid password", vbCritical, "Invalid password"
                UserControl41.forceClick (prevIndex)
                lblMain.Caption = prevCapt
                Exit Sub
            End If
        End If
    End If
    If protectInfo = True Then
        If Index = 4 Then
            Dim tempPass As String
            tempPass = InputBox("Password set for Process Information. Enter Password to continue", "Enter Password")
            If tempPass <> protectPass Then
                MsgBox "Invalid password", vbCritical, "Invalid password"
                UserControl41.forceClick (prevIndex)
                lblMain.Caption = prevCapt
                Exit Sub
            End If
        End If
    End If
    For i = 1 To frmMain.UBound
        If Index = i Then
            frmMain(i).Visible = True
        Else
            frmMain(i).Visible = False
        End If
    Next i
    showGo = True
    Call showFrm2(99)
End Sub

Private Sub showFrm2(Index As Integer)
    For i = 1 To frmMain2.UBound
        If Index = i Then
            frmMain2(i).Visible = True
        Else
            frmMain2(i).Visible = False
        End If
    Next i
End Sub

Private Sub applyInfo(prNameof As String)
Dim iDex As Integer
Dim iDex2 As Integer
Dim foundIN As Boolean
    foundIN = False
    For i = 0 To UBound(procinfo)
        If prNameof = procinfo(i).procName Then
            iDex = i
            foundIN = True
            Exit For
        End If
    Next i
    For i = 0 To UBound(jailInfo)
        If prNameof = jailInfo(i).exeName Then
            iDex2 = i
            Exit For
        End If
    Next i
    If foundIN = True Then
        lblProcName.Caption = procinfo(iDex).procName
        lblProcID.Caption = procinfo(iDex).th32ProcessID
        lblProcIDs.Caption = procinfo(iDex).th32ProcessID
        lblParentID.Caption = procinfo(iDex).th32ParentProcessID
        lblThreads.Caption = procinfo(iDex).cntThreads
        lblChd.Caption = procinfo(iDex).childWnd
    '-----------Process Priority----------------
        Select Case getPriority(procinfo(iDex).th32ProcessID)
            Case 32
                lblPriority.Caption = "Normal"
            Case 64
                lblPriority.Caption = "Idle"
            Case 128
                lblPriority.Caption = "High"
            Case 256
                lblPriority.Caption = "RealTime"
        End Select
    '-------------------------------------------
    Else
        lblProcName.Caption = jailInfo(iDex2).exeName
        lblProcID.Caption = "N/A"
        lblProcIDs.Caption = "N/A"
        lblParentID.Caption = "N/A"
        lblThreads.Caption = "N/A"
        lblChd.Caption = "N/A"
    End If
    If iDex2 <> 0 Then
        lblAccessAttempts.Caption = jailInfo(iDex2).attempts
        lblLastAct.Caption = jailInfo(iDex2).prevAction
        If jailInfo(i).prevAction = "Blocked" Then
            lblLastAct.ForeColor = &HC0&
        Else
            lblLastAct.ForeColor = &H8000&
        End If
        lblFirstAccess.Caption = jailInfo(iDex2).firstTime
        lbllastAccess.Caption = jailInfo(iDex2).lastTime
    Else
        lblAccessAttempts.Caption = 0
        lblLastAct.Caption = "N/A"
        lblLastAct.ForeColor = &H80000012
        lblFirstAccess.Caption = "N/A"
        lbllastAccess.Caption = "N/A"
    End If
End Sub

Private Sub setMonitor()
    If tmrON = True Then
        tmrProcRef.Enabled = False
        lblMonitor.Caption = "OFF"
        lblMonitor.ForeColor = &HC0&
        uc7Monitor.Caption = "Monitor Processes"
        tmrON = False
    Else
        monitorOn = True
        tmrProcRef.Enabled = True
        Call uc7Interval_Clicked
        lblMonitor.Caption = "ON"
        lblMonitor.ForeColor = &H8000&
        uc7Monitor.Caption = "Stop Monitoring"
        tmrON = True
    End If
End Sub

Public Function getPriority(pid As Long)
Dim hwnd2 As Long
    hwnd2 = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    pri = GetPriorityClass(hwnd2)
    CloseHandle hwnd2
    getPriority = pri
End Function

Public Sub setPass()
retry:
        protectPass = InputBox("Choose password:", "Enter New Password")
        If protectPass = "" Then
            MsgBox "Invalid password", vbCritical, "Error"
            GoTo retry
        End If
End Sub

Private Sub loadIcon()
    traymsg = "Process Gateway"
    TrayI.cbSize = Len(TrayI)
    'Set the window's handle (this will be used to hook the specified window)
    TrayI.hwnd = Form1.hwnd
    'Application-defined identifier of the taskbar icon
    TrayI.uId = 1&
    'Set the flags
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'Set the callback message
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    'Set the picture (must be an icon!)
    TrayI.hIcon = Form1.Icon
    'Set the tooltiptext
    TrayI.szTip = traymsg & Chr$(0)
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI
End Sub

Private Sub removeIcon()
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = Form1.hwnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
End Sub

Private Sub rdyLog()
    Set colHead = lstVwLog.ColumnHeaders.Add(lstVwLog.ColumnHeaders.Count + 1, "Process Name", "Process Name", TextWidth("Process Name") * 1.5)
    Set colHead = lstVwLog.ColumnHeaders.Add(lstVwLog.ColumnHeaders.Count + 1, "Process ID", "Process ID", TextWidth("Process ID") * 1.3)
    Set colHead = lstVwLog.ColumnHeaders.Add(lstVwLog.ColumnHeaders.Count + 1, "Action Taken", "Action Taken", TextWidth("Action Taken") * 1.5)
    Set colHead = lstVwLog.ColumnHeaders.Add(lstVwLog.ColumnHeaders.Count + 1, "Time", "Time", TextWidth("Time") * 3.1)
    Set colHead = lstVwLog.ColumnHeaders.Add(lstVwLog.ColumnHeaders.Count + 1, "Attempts", "Attempts", TextWidth("Attempts") * 1.5)
End Sub

Private Sub vChk_ButtonClicked(Index As Integer, Checked As Boolean)
    Select Case Index
    
        Case 0
            Dim val As String
            If vChk(0).Checked = True Then
                val = App.Path & "\" & App.exeName & ".exe"
                RegOpenKeyEx HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 0, KEY_ALL_ACCESS, result
                RegSetValueEx result, "Process Gateway", 0, REG_SZ, ByVal val, Len(val)
                RegCloseKey result
            Else
                RegOpenKeyEx HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 0, KEY_ALL_ACCESS, result
                RegDeleteValue result, "Process Gateway"
            End If
        Case 1
            If vChk(1).Checked = True Then
                If protectPass = "" Then
                    Call setPass
                End If
                protectOpt = True
                MsgBox "Password set for options", vbOKOnly, "Options Protected"
                UserControl42.Reset
                UserControl41.forceClick (1)
                UserControl41.Width = 2460
                Call showFrm(2)
            Else
                protectOpt = False
            End If
        Case 2
            If vChk(2).Checked = True Then
                txtPass.Enabled = True
                uc7Password.Enabled = True
            Else
                txtPass.Enabled = False
                uc7Password.Enabled = False
            End If
        Case 3
            If vChk(3).Checked = True Then
                If protectPass = "" Then
                    Call setPass
                End If
                protectLogs = True
                MsgBox "Password set for Logs", vbOKOnly, "Logs Protected"
            Else
                protectLogs = False
            End If
        Case 4
            If vChk(4).Checked = True Then
                logOn = False
                lblLogStatus.Caption = "[Logging Disabled]"
                lblLogStatus.ForeColor = &HC0&
            Else
                logOn = True
                lblLogStatus.Caption = "[Logging Enabled]"
                lblLogStatus.ForeColor = &H8000&
            End If
        Case 5
            If vChk(5).Checked = True Then
                If protectPass = "" Then
                    Call setPass
                End If
                protectAccess = True
                MsgBox "Password set for Access Control", vbOKOnly, "Access Control Protected"
            Else
                protectAccess = False
            End If
        Case 6
            If vChk(6).Checked = True Then
                vChk(7).Checked = True
                vChk(3).Checked = True
                vChk(5).Checked = True
                vChk(1).Checked = True
                vChk(7).Enabled = False
                vChk(3).Enabled = False
                vChk(5).Enabled = False
                vChk(1).Enabled = False
            Else
                vChk(7).Enabled = True
                vChk(3).Enabled = True
                vChk(5).Enabled = True
                vChk(1).Enabled = True
                vChk(7).Checked = False
                vChk(3).Checked = False
                vChk(5).Checked = False
                vChk(1).Checked = False
            End If
        Case 7
            If vChk(7).Checked = True Then
                If protectPass = "" Then
                    Call setPass
                End If
                protectInfo = True
                MsgBox "Password set for Process Information", vbOKOnly, "Process Information Protected"
            Else
                protectInfo = False
            End If
        Case 8
            If vChk(8).Checked = True Then
                Call setSett
            Else
                SaveSetting "PGateway", "Main", "Load", "No"
            End If
        Case 9
            If vChk(9).Checked = True Then
                If protectPass = "" Then
                    Call setPass
                End If
                hotkeyPrompt = True
            Else
                hotkeyPrompt = False
            End If
    End Select
End Sub

Public Sub vChk1_ButtonClicked(Index As Integer, Checked As Boolean)
    Select Case Index
    
        Case 0
            Dim style As Long
            Hide
            style = GetWindowLong(hwnd, GWL_EXSTYLE)
            If vChk1(0).Checked = True Then
                If style And WS_EX_APPWINDOW Then
                    style = style - WS_EX_APPWINDOW
                End If
            Else
                style = style Or WS_EX_APPWINDOW
            End If
            SetWindowLong hwnd, GWL_EXSTYLE, style
            App.TaskVisible = vChk1(0).Checked
            Show
        Case 1
            If vChk1(1).Checked = True Then
                removeIcon
            Else
                loadIcon
            End If
        Case 2
            If vChk1(2).Checked = True Then
                Load frmProcessHack
                frmProcessHack.Show
            Else
                frmProcessHack.lstHackProcTo.Selected(frmProcessHack.lstHackProcTo.ListIndex) = False
                frmProcessHack.lstProcto.Selected(frmProcessHack.lstProcto.ListIndex) = False
                frmProcessHack.uc7Hack.Enabled = False
                lblHostProc.Caption = ""
                lblIdent.Caption = ""
                lblAppName.Caption = ""
                lblHackStat.Caption = "[DISABLED]"
                lblHackStat.ForeColor = &H80&
                ProcessHack.Disable
            End If
        Case 3
            If vChk1(3).Checked = True Then
               App.TaskVisible = False
            Else
               App.TaskVisible = True
            End If
        Case 4
            If vChk1(4).Checked = True Then
                tmrFindRun.Enabled = True
            Else
                tmrFindRun.Enabled = False
            End If
        Case 5
            If vChk1(5).Checked = True Then
                Open "C:\Windows\System32\taskmgr.exe" For Binary As #1
            Else
                Close #1
            End If
        Case 6
            If vChk1(6).Checked = True Then
                If keyOn = 0 Then
                    Dim rep As String
                    rep = MsgBox("You must enable a HotKey for ProcessGateway. Do you wish to do this now?", vbYesNo, "Enable Hotkey")
                    If rep = vbYes Then
                        oldProc = SetWindowLongA(Me.hwnd, GWL_WNDPROC, AddressOf WndProc)
                        HotKeyActivate Me.hwnd, 1, Asc("P")
                    Else
                        vChk1(0).Checked = False
                    End If
                End If
                If keyOn = 1 Then
                    MsgBox "Hot key enabled. Process gateway will hide itself when minimized, to restore it, press CTRL+SHIFT+P", vbOKOnly, "Hotkey Enabled"
                End If
            Else
                MsgBox "Hot key disabled"
            End If
    End Select
End Sub

Private Sub vChk3_ButtonClicked(Index As Integer, Checked As Boolean)
    Select Case Index
        Case 0
            If vChk3(0).Checked = True Then
                logNew = True
            Else
                logNew = False
            End If
        Case 1
            If vChk3(1).Checked = True Then
                If vChk3(2).Checked = True Then
                    vChk3(2).Checked = False
                End If
            Else
                
            End If
        Case 2
            If vChk3(2).Checked = True Then
                If vChk3(1).Checked = True Then
                    vChk3(1).Checked = False
                End If
            Else
                
            End If
        Case 3
            If vChk3(3).Checked = True Then
                'N/A. Code in frmNew
            Else
                'N/A. Code in frmNew
            End If
    End Select
End Sub

Public Sub loadSett()
Dim tempSet As String
Dim tempChar As String
Dim i As Integer
    tempSet = GetSetting("PGateway", "Main", "Load", "Error")
    protectPass = GetSetting("PGateway", "Main", "Pass", "Error")
    If InStr(1, tempSet, "Yes") Then
'-----------------------------------------------
        For i = 0 To vChk.UBound
            tempChar = Mid(GetSetting("PGateway", "Main", "vChk", "Error"), i + 1, 1)
            If tempChar = "1" Then
                vChk(i).Checked = True
                Call vChk_ButtonClicked(i, False)
            Else
                vChk(i).Checked = False
            End If
        Next i
'-----------------------------------------------
        For i = 0 To vChk1.UBound
            tempChar = Mid(GetSetting("PGateway", "Main", "vChk1", "Error"), i + 1, 1)
            If tempChar = "1" Then
                vChk1(i).Checked = True
                Call vChk1_ButtonClicked(i, True)
            Else
                vChk1(i).Checked = False
            End If
        Next i
'-----------------------------------------------
        For i = 0 To vChk3.UBound
            tempChar = Mid(GetSetting("PGateway", "Main", "vChk3", "Error"), i + 1, 1)
            If tempChar = "1" Then
                vChk3(i).Checked = True
                Call vChk3_ButtonClicked(i, True)
            Else
                vChk3(i).Checked = False
            End If
        Next i
'-----------------------------------------------
    End If
End Sub

Public Sub setSett()
    SaveSetting "PGateway", "Main", "Load", "Yes"
    optString = ""
    For i = 0 To vChk.UBound
        If vChk(i).Checked = True Then
            optString = optString & 1
        Else
            optString = optString & 0
        End If
    Next i
    SaveSetting "PGateway", "Main", "vChk", optString
    optString = ""
    For i = 0 To vChk1.UBound
        If vChk1(i).Checked = True Then
            optString = optString & 1
        Else
            optString = optString & 0
        End If
    Next i
    SaveSetting "PGateway", "Main", "vChk1", optString
    optString = ""
    For i = 0 To vChk3.UBound
        If vChk3(i).Checked = True Then
            optString = optString & 1
        Else
            optString = optString & 0
        End If
    Next i
    SaveSetting "PGateway", "Main", "vChk3", optString
    SaveSetting "PGateway", "Main", "Pass", protectPass
End Sub
