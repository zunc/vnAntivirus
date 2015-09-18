VERSION 5.00
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmMas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MsgBox"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmMas.frx":0000
      Top             =   0
      Width           =   3735
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   275
      Index           =   2
      Left            =   0
      Picture         =   "frmMas.frx":0006
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   250
      Begin VB.Label lblR 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   275
      Index           =   1
      Left            =   0
      Picture         =   "frmMas.frx":0348
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   250
      Begin VB.Label lblP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   275
      Index           =   0
      Left            =   0
      Picture         =   "frmMas.frx":068A
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      Begin VB.Label lblC 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
   End
   Begin CtrUnicodeVN.LabelUni lblMas 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   1440
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   1440
   End
   Begin CtrUnicodeVN.LabelUni lblMas 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblMas 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblMas 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      SetKieugoTV     =   1
      Appearance      =   0
      Alignment       =   2
      ForeColor       =   255
   End
   Begin CtrUnicodeVN.LabelUni lblPath 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Path :"
      Appearance      =   0
   End
   Begin CtrUnicodeVN.ButtonUni cmdCancel 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmMas.frx":09CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      SetKieugoTV     =   1
   End
   Begin CtrUnicodeVN.ButtonUni cmdResume 
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Resume process"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmMas.frx":09E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      SetKieugoTV     =   1
   End
   Begin VB.Label lblID 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblCaption 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "frmMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SPI_GETWORKAREA As Long = 48&
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long

Private m_iChangeSpeed    As Long
Private m_iCounter        As Long
Private m_iScrnBottom     As Long
Private m_bOnTop          As Boolean
Private m_iWindowCount    As Long
Private m_bManualClose    As Boolean
Private m_bCodeClose      As Boolean
Private m_bFade           As Boolean
Private m_iOSver          As Byte
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim tg As Byte
'Dim formCap As String
'Dim mas As String
Dim co As Byte
'Dim com As String

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdResume_Click()
    SuspendResumeProcess Val(lblID.Caption), False
End Sub
Private Sub Form_Load()
    Me.Caption = lblCaption.Caption
    'Me.BackColor = RGB(255, 255, 223)
  Dim rc         As RECT
  Dim scrnRight  As Long
  Dim OSV        As OSVersionInfo
        OSV.OSVSize = Len(OSV)
        
    If GetVersionEx(OSV) = 1 Then
        If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then m_iOSver = 1 '/* Win 98/ME
        If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then m_iOSver = 2  '/* Win 2000/XP
    End If
    
    '/* Get Screen and TaskBar size
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, rc, 0&)
    
    '/* Screen Height - Taskbar Height (if is is located at the bottom of the screen)
    m_iScrnBottom = rc.Bottom * Screen.TwipsPerPixelY
    
    '/* Is the taskbar is located on the right side of the screen? (scrnRight < Screen.width)
    scrnRight = (rc.Right * Screen.TwipsPerPixelX)
    
    '/* Locate Form to bottom right and set default size
    Me.Move scrnRight - Me.Width, m_iScrnBottom, lblMas(3).Left + lblMas(3).Width + 100, cmdCancel.Top + 700
    
    'Me.Move scrnRight - Me.Width, m_iScrnBottom - Me.Height, txtMas.Left + txtMas.Width + 100, cmdScan.Top + 800
    
    Timer2.Enabled = True
    tg = 10
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub lblMas_Click(Index As Integer)
    If Timer.Enabled = True Then
        Timer.Enabled = False
    Else
        Timer.Enabled = True
    End If
End Sub

Private Sub Timer_Timer()
If tg - co > 0 Then
    co = co + 1
    Me.Caption = lblCaption.Caption & " (" & tg - co & ")"
Else
    Timer1.Enabled = True
End If
End Sub
Private Sub Timer1_Timer()
    If Me.Top > m_iScrnBottom Then
        Unload Me
    Else
        Me.Move Me.Left, Me.Top + 25, lblMas(3).Left + lblMas(3).Width + 100, cmdCancel.Top + 700
    End If
End Sub
Private Sub Timer2_Timer()
    If Me.Top < m_iScrnBottom - Me.Height Then
        Timer.Enabled = True
        Timer2.Enabled = False
    Else
        Me.Caption = lblCaption.Caption
        Me.Move Me.Left, Me.Top - 25, lblMas(3).Left + lblMas(3).Width + 100, cmdCancel.Top + 700
    End If
End Sub
Private Sub txtPath_Click()
    Timer.Enabled = False
    With txtPath
    .SelStart = 0
    .SelLength = Len(txtPath)
    End With
End Sub
