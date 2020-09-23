VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "st0rage"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   9180
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Index           =   1
      Left            =   4680
      ScaleHeight     =   2505
      ScaleWidth      =   2100
      TabIndex        =   35
      Top             =   3840
      Visible         =   0   'False
      Width           =   2100
      Begin VB.TextBox txtGen 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   315
         TabIndex        =   41
         Top             =   2025
         Width           =   1425
      End
      Begin VB.CheckBox ckLCase 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H006E5755&
         Height          =   225
         Left            =   465
         MaskColor       =   &H00808080&
         TabIndex        =   40
         Top             =   570
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.CheckBox ckUCase 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H006E5755&
         Height          =   225
         Left            =   465
         MaskColor       =   &H00808080&
         TabIndex        =   39
         Top             =   855
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.CheckBox ckNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H006E5755&
         Height          =   225
         Left            =   465
         MaskColor       =   &H00808080&
         TabIndex        =   38
         Top             =   1140
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.TextBox txtLen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   37
         Text            =   "10"
         Top             =   195
         Width           =   525
      End
      Begin VB.CheckBox ckSpec 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H006E5755&
         Height          =   225
         Left            =   465
         MaskColor       =   &H00808080&
         TabIndex        =   36
         Top             =   1425
         Width           =   210
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lowercase"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   750
         TabIndex        =   47
         Top             =   570
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "uppercase"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   750
         TabIndex        =   46
         Top             =   855
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "numbers"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   750
         TabIndex        =   45
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "length"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   450
         TabIndex        =   44
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lblGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "generate"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   315
         TabIndex        =   43
         Top             =   1725
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "specials"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   750
         TabIndex        =   42
         Top             =   1425
         Width           =   975
      End
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Index           =   6
      Left            =   285
      Picture         =   "frmMain.frx":2DBEC
      ScaleHeight     =   2505
      ScaleWidth      =   2100
      TabIndex        =   31
      Top             =   6435
      Visible         =   0   'False
      Width           =   2100
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   345
         PasswordChar    =   "*"
         TabIndex        =   32
         Top             =   900
         Width           =   1425
      End
      Begin VB.Label lblEntry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "enter"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   1500
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "enter password"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   10
         Left            =   375
         TabIndex        =   33
         Top             =   585
         Width           =   1515
      End
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Index           =   5
      Left            =   6615
      Picture         =   "frmMain.frx":3EE2A
      ScaleHeight     =   2505
      ScaleWidth      =   2100
      TabIndex        =   28
      Top             =   1140
      Visible         =   0   'False
      Width           =   2100
      Begin VB.TextBox txtAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1920
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   270
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label lblCredits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "credits"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   735
         TabIndex        =   50
         Top             =   2175
         Width           =   630
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "[code]  rolex"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   13
         Left            =   300
         TabIndex        =   49
         Top             =   795
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "password fox"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   15
         Left            =   300
         TabIndex        =   48
         Top             =   390
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "help"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   90
         TabIndex        =   30
         Top             =   2175
         Width           =   615
      End
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Index           =   3
      Left            =   4320
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   24
      Top             =   1110
      Visible         =   0   'False
      Width           =   2100
      Begin VB.TextBox txtExtract 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   375
         TabIndex        =   26
         Top             =   1845
         Width           =   1470
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Height          =   495
         Index           =   1
         Left            =   150
         Top             =   900
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   390
         TabIndex        =   27
         Top             =   1590
         Width           =   975
      End
      Begin VB.Image imgTarget 
         Height          =   480
         Left            =   300
         Picture         =   "frmMain.frx":50068
         Top             =   990
         Width           =   480
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "drag the arrow over a windows textbox to attempt to extract the password"
         ForeColor       =   &H00FFFFFF&
         Height          =   1050
         Index           =   11
         Left            =   285
         TabIndex        =   25
         Top             =   315
         Width           =   1740
      End
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Index           =   2
      Left            =   6795
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   2100
      Begin VB.PictureBox picPercent 
         BackColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   315
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   21
         Top             =   1665
         Width           =   15
      End
      Begin VB.TextBox txtQual 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   285
         TabIndex        =   19
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "rating:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   270
         TabIndex        =   23
         Top             =   1335
         Width           =   660
      End
      Begin VB.Label lblQual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   870
         TabIndex        =   22
         Top             =   1335
         Width           =   960
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Height          =   240
         Index           =   0
         Left            =   270
         Top             =   1620
         Width           =   1590
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "check password strength"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   8
         Left            =   375
         TabIndex        =   20
         Top             =   405
         Width           =   1410
      End
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Index           =   4
      Left            =   2490
      ScaleHeight     =   2505
      ScaleWidth      =   2100
      TabIndex        =   12
      Top             =   3855
      Visible         =   0   'False
      Width           =   2100
      Begin VB.CheckBox ckSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H006E5755&
         Height          =   225
         Left            =   120
         MaskColor       =   &H00808080&
         TabIndex        =   51
         Top             =   1200
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.TextBox txtStart 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   390
         TabIndex        =   15
         Text            =   "password"
         Top             =   705
         Width           =   1425
      End
      Begin VB.CheckBox ckStart 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H006E5755&
         Height          =   225
         Left            =   105
         MaskColor       =   &H00808080&
         TabIndex        =   14
         Top             =   270
         Width           =   210
      End
      Begin VB.Label lblSave 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "save data"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   390
         TabIndex        =   53
         Top             =   1800
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "autosave"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   16
         Left            =   480
         TabIndex        =   52
         Top             =   1215
         Width           =   885
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "require password on startup"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   3
         Left            =   420
         TabIndex        =   13
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Index           =   0
      Left            =   300
      ScaleHeight     =   2505
      ScaleWidth      =   2100
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   2100
      Begin VB.ListBox lstType 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1200
         ItemData        =   "frmMain.frx":526A2
         Left            =   210
         List            =   "frmMain.frx":526A4
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         Width           =   1710
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   780
         TabIndex        =   8
         Text            =   "name"
         Top             =   1800
         Width           =   1155
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   780
         TabIndex        =   7
         Text            =   "pass"
         Top             =   2130
         Width           =   1155
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H006E5755&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Text            =   "description"
         Top             =   1470
         Width           =   1725
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   1815
         Width           =   525
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "pass"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   225
         TabIndex        =   10
         Top             =   2130
         Width           =   525
      End
   End
   Begin VB.Image imgNull 
      Height          =   15
      Left            =   4815
      Picture         =   "frmMain.frx":526A6
      Top             =   720
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgArrow 
      Height          =   480
      Left            =   4515
      Picture         =   "frmMain.frx":52AEC
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H006E5755&
      BackStyle       =   0  'Transparent
      Caption         =   "exit"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2835
      TabIndex        =   18
      Top             =   2505
      Width           =   975
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H006E5755&
      BackStyle       =   0  'Transparent
      Caption         =   "info"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2835
      TabIndex        =   17
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H006E5755&
      BackStyle       =   0  'Transparent
      Caption         =   "options"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2835
      TabIndex        =   4
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H006E5755&
      BackStyle       =   0  'Transparent
      Caption         =   "recovery"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2835
      TabIndex        =   3
      Top             =   1305
      Width           =   975
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H006E5755&
      BackStyle       =   0  'Transparent
      Caption         =   "security"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2835
      TabIndex        =   2
      Top             =   1005
      Width           =   975
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H006E5755&
      BackStyle       =   0  'Transparent
      Caption         =   "generator"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2835
      TabIndex        =   1
      Top             =   705
      Width           =   975
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H006E5755&
      BackStyle       =   0  'Transparent
      Caption         =   "storage"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2835
      TabIndex        =   0
      Top             =   405
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Stats
    sDesc As String
    sName As String
    sPass As String
End Type
Private bStart As Boolean
Private sStart As String

Private Stat(100) As Stats
Private iList As Integer
Private bTarget As Boolean

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10
    





Private Sub ckStart_Click()
    If ckStart.Value = 0 Then
        txtStart.Enabled = False
    Else
        txtStart.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Dim iCount As Integer
    
    For iCount = 0 To 6
        picMenu(iCount).Left = 615
        picMenu(iCount).Top = 317
    Next iCount
    For iCount = 0 To 4
        picMenu(iCount).Picture = picMenu(6).Picture
    Next iCount
        

    FormRegion Me, vbWhite
    LoadPWData
    
    For iCount = 0 To lstType.ListCount
        If lstType.List(iCount) = "" Then
            lstType.ListIndex = iCount
            txtDesc.SelStart = Len(txtDesc.Text)
            Exit Sub
        End If
    Next iCount

    

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me.hwnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iCount As Integer
    For iCount = 0 To 6
        lblMenu(iCount).BackStyle = 0
    Next iCount
End Sub



Private Sub imgTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bTarget = True
    imgTarget.Picture = imgNull.Picture
    
    Me.MousePointer = 99
    Me.MouseIcon = imgArrow.Picture
    txtExtract.Text = ""
    
End Sub



Private Sub imgTarget_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TargetLen As Long
    Dim TempString As String
    Dim hwnd As Long
    
    bTarget = False
    
    imgTarget.Picture = imgArrow.Picture
    Me.MousePointer = 0


    Call GetCursorPos(CursorPosition)
    hwnd = WindowFromPoint(CursorPosition.X, CursorPosition.Y)
    hwnd = GetTopLevelParent(hwnd)
    TargetLen& = SendMessage(hwnd&, WM_GETTEXTLENGTH, 0&, 0&)
    TempString$ = String(TargetLen&, 0&)
    Call SendMessageByString(hwnd&, WM_GETTEXT, TargetLen& + 1, TempString$)
    txtExtract.Text = TempString$

End Sub













Private Sub lblCredits_Click()
    SoundPlay
    txtAbout.Visible = False
    If lblInfo(13).Visible = False Then
        lblInfo(13).Visible = True
        lblInfo(15).Visible = True
    Else
        lblInfo(13).Visible = False
        lblInfo(15).Visible = False
    End If
End Sub

Private Sub lblEntry_Click()
    SoundPlay
    If txtEntry.Text = txtStart.Text Then picMenu(6).Visible = False
End Sub

Private Sub lblEntry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEntry.BackStyle = 1
End Sub

Private Sub lblGen_Click()
    SoundPlay
    txtGen.Text = PWGenerate(Val(txtLen.Text))
End Sub

Private Sub lblGen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblGen.BackStyle = 1
End Sub

Private Sub lblHelp_Click()
    SoundPlay
    LoadHelpText
    lblInfo(13).Visible = False
    lblInfo(15).Visible = False
    If txtAbout.Visible = True Then
        txtAbout.Visible = False
    Else
        txtAbout.Visible = True
    End If
    
End Sub







Private Sub lblMenu_Click(Index As Integer)
    Dim iCount As Integer
    
    SoundPlay
    
    For iCount = 0 To 5
        picMenu(iCount).Visible = False
    Next iCount
    
    If Index <> 6 Then
        picMenu(Index).Visible = True
    Else
        If ckSave.Value = 1 Then SavePWData
        End
    End If
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iCount As Integer
    For iCount = 0 To 6
        lblMenu(iCount).BackStyle = 0
    Next iCount
    lblMenu(Index).BackStyle = 1
End Sub

Private Sub lblSave_Click()
    SoundPlay
    SavePWData
End Sub

Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSave.BackStyle = 1
End Sub



Private Sub lstType_Click()
    Dim iCount As Integer
    

        txtDesc.Text = Stat(lstType.ListIndex).sDesc
        txtName.Text = Stat(lstType.ListIndex).sName
        txtPass.Text = Stat(lstType.ListIndex).sPass

End Sub

Private Sub lstType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iList = lstType.ListIndex
End Sub



Private Sub picMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblGen.BackStyle = 0
    lblEntry.BackStyle = 0
    lblSave.BackStyle = 0
End Sub



Private Sub txtDesc_Click()
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc.Text)
End Sub

Private Sub txtDesc_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstType.ListIndex <> -1 Then
        Stat(lstType.ListIndex).sDesc = txtDesc.Text
        lstType.List(lstType.ListIndex) = txtDesc.Text
    End If
End Sub

Private Sub txtDesc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstType.ListIndex <> -1 Then
        Stat(lstType.ListIndex).sDesc = txtDesc.Text
        lstType.List(lstType.ListIndex) = txtDesc.Text
    End If

End Sub

Private Sub txtGen_Click()
    txtGen.SelStart = 0
    txtGen.SelLength = Len(txtGen.Text)
    Clipboard.SetText txtGen.Text
End Sub

Private Sub txtLen_Change()
    If IsNumeric(txtLen.Text) = False And txtLen.Text <> "" Then txtLen.Text = "10"
End Sub

Private Sub txtLen_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Is < 42
        Case 48 To 57
        Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtName_Click()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstType.ListIndex <> -1 Then
        Stat(lstType.ListIndex).sName = txtName.Text
    End If
End Sub

Private Sub txtPass_Click()
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
End Sub
Private Sub LoadPWData()
    Dim MyString As String
    Dim iTemp As Integer
    On Error GoTo eHandle
    Open App.Path & "\st0rage.dat" For Input As #1
    
    
    
    Input #1, MyString$
    ckStart.Value = Val(DeCrypt(MyString))
    
    If ckStart.Value = 1 Then
        picMenu(6).Visible = True
        txtStart.Enabled = True
    End If
    
    Input #1, MyString$
    txtStart.Text = DeCrypt(MyString)
    
    Input #1, MyString$
    ckSave.Value = Val(DeCrypt(MyString))

    
    Do
        
        Input #1, MyString$
        DoEvents
        Stat(iTemp).sDesc = DeCrypt(MyString)
        lstType.AddItem (Stat(iTemp).sDesc)
        
        Input #1, MyString$
        DoEvents
        Stat(iTemp).sName = DeCrypt(MyString)
        
        Input #1, MyString$
        DoEvents
        Stat(iTemp).sPass = DeCrypt(MyString)
        
        iTemp = iTemp + 1
    Loop
    
    Close #1
    
eHandle:
    Close #1
    Dim iCount As Integer
    For iCount = iTemp To 100
        With Stat(iCount)
            .sDesc = ""
            .sName = ""
            .sPass = ""
        End With
        lstType.AddItem ""
    Next iCount
    Load frmMain
    frmMain.Show
    
End Sub
Private Sub SavePWData()
    Dim SaveList As Integer

    
    'On Error Resume Next
    Open App.Path & "\st0rage.dat" For Output As #1
    
    

    Print #1, EnCrypt(Str(ckStart.Value))
    Print #1, EnCrypt(txtStart.Text)
    
    Print #1, EnCrypt(Str(ckSave.Value))

    For SaveList = 0 To lstType.ListCount - 2
        Print #1, EnCrypt(Stat(SaveList).sDesc)
        Print #1, EnCrypt(Stat(SaveList).sName)
        Print #1, EnCrypt(Stat(SaveList).sPass)
    Next SaveList
    
    Close #1
    Exit Sub


End Sub

Private Sub ClearStats()
    Dim iCount As Integer
    For iCount = 0 To 100
        With Stat(iCount)
            .sDesc = ""
            .sName = ""
            .sPass = ""
        End With
    Next iCount
End Sub



Private Sub txtPass_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstType.ListIndex <> -1 Then
        Stat(lstType.ListIndex).sPass = txtPass.Text
    End If
End Sub


Private Sub SoundPlay()
    Dim sFile As String
    Dim iFlag As Integer
    
    sFile = App.Path & "\menu.wav"
    iFlag = SND_ASYNC Or SND_NODEFAULT
    
    sndPlaySound sFile, iFlag
    
End Sub
Function PWGenerate(Length As Long) As String
    Dim sTemp As String
    Dim iCount As Integer
    Dim sChars As String

    If ckLCase.Value = 1 Then sChars = sChars & "abcdefghijklmnopqrstuvwxyz"
    If ckUCase.Value = 1 Then sChars = sChars & "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If ckNum.Value = 1 Then sChars = sChars & "01234567890123456789"
    If ckSpec.Value = 1 Then sChars = sChars & "!@#$%^&*()-_+=[]{}|\:;'?<>,./~`"
    
    For iCount = 1 To Length
        Randomize
        sTemp = sTemp & Mid(sChars, (Rnd * (Len(sChars) - 1)) + 1, 1)
    Next iCount
    
    PWGenerate = sTemp
End Function




Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstType.ListIndex <> -1 Then
        Stat(lstType.ListIndex).sPass = txtPass.Text
    End If
End Sub

Private Sub txtQual_Change()
    Dim iQual As Integer
    
    iQual = KeyQuality(txtQual.Text)
    
    If iQual >= 0 And iQual <= 19 Then
        lblQual.Caption = "poor"
    ElseIf iQual >= 20 And iQual <= 39 Then
        lblQual.Caption = "average"
    ElseIf iQual >= 40 And iQual <= 59 Then
        lblQual.Caption = "good"
    ElseIf iQual >= 60 And iQual <= 79 Then
        lblQual.Caption = "very good"
    ElseIf iQual >= 80 And iQual <= 98 Then
        lblQual.Caption = "excellent"
    ElseIf iQual >= 99 And iQual <= 100 Then
        If lblQual.Caption <> "flawless" Then SoundPlay
        lblQual.Caption = "flawless"
    End If
    
    picPercent.Width = iQual
End Sub


Public Sub LoadHelpText()

txtAbout.Text = ""


txtAbout.Text = txtAbout.Text & "•¤[ storage ]¤•" & vbCrLf
txtAbout.Text = txtAbout.Text & "this is your main password database.  to create a new entry, click on a blank line in the listbox.  on the first textbox add your description/url/etc.  then fill in your username and password in the appropriate textboxes. remember: ctrl + v is paste and ctrl + c is copy.  when you exit the program, your info will automatically be encrypted and saved (if autosave is on) or manually when you click 'save data' in options." & vbCrLf & vbCrLf

txtAbout.Text = txtAbout.Text & "•¤[ generator ]¤•" & vbCrLf
txtAbout.Text = txtAbout.Text & "this will automatically generate a random password, which can be customized by specifying the length and what kind of characters it has." & vbCrLf & vbCrLf

txtAbout.Text = txtAbout.Text & "•¤[ security ]¤•" & vbCrLf
txtAbout.Text = txtAbout.Text & "this function will check how hard a password is to crack, based on many variables.  my suggestion is to atleast make sure your password is 'excellent' or atleast 'very good'." & vbCrLf & vbCrLf

txtAbout.Text = txtAbout.Text & "•¤[ recovery ]¤•" & vbCrLf
txtAbout.Text = txtAbout.Text & "this tool is useful for textboxes that have passwords stored in them.  the passwords are usually hidden by ****** symbols.  drag the arrow icon on top of the textbox to attempt to reveal the hidden password." & vbCrLf & vbCrLf

txtAbout.Text = txtAbout.Text & "•¤[ options ]¤•" & vbCrLf
txtAbout.Text = txtAbout.Text & "to password protect this program, enable the checkbox and type in a password.  if autosave is off, make sure you click 'save data' to save all of your entries and settings." & vbCrLf & vbCrLf

txtAbout.Text = txtAbout.Text & "•¤[ rolex ]¤•" & vbCrLf & vbCrLf







End Sub











