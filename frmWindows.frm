VERSION 5.00
Begin VB.Form frmWindows 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Associated Windows"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form2"
   ScaleHeight     =   4770
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ProcessGateway.UserControl2 UserControl21 
      Height          =   4815
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8493
      Begin ProcessGateway.UserControl7 uc7Exit 
         Height          =   360
         Left            =   4200
         TabIndex        =   5
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         Hold_Caption    =   "Exit"
      End
      Begin VB.ListBox lstWindows 
         Height          =   2595
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   5535
      End
      Begin VB.Label lblpID 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   3480
         TabIndex        =   6
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblpName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   3480
         TabIndex        =   4
         Top             =   120
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   4920
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblMain 
         BackStyle       =   0  'Transparent
         Caption         =   "Associated Windows for:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   3255
      End
   End
   Begin ProcessGateway.UserControl3 UserControl31 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   9128
   End
End
Attribute VB_Name = "frmWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    uc7Exit.SubClassMe
    lblpName.Caption = winName
    lblpID.Caption = winID
    If wndNames(0) = "" Then
        lstWindows.AddItem "No Associated Windows"
    End If
    For i = 0 To UBound(wndNames)
        If wndNames(i) = "" Then
            Exit For
        Else
            lstWindows.AddItem wndNames(i)
        End If
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    uc7Exit.UnSubClassMe
End Sub

Private Sub uc7Exit_Clicked()
    Unload Me
End Sub
