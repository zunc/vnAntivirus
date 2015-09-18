VERSION 5.00
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmMas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MsgBox"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CtrUnicodeVN.LabelUni lblMes 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   1200
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   1200
   End
   Begin CtrUnicodeVN.ButtonUni cmdCancel 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Hu3y"
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
      MICON           =   "frmMas.frx":0000
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
   Begin CtrUnicodeVN.ButtonUni cmdScan 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Que1t"
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
      MICON           =   "frmMas.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      SetKieugoTV     =   1
   End
   Begin VB.Label lblCaption 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblPath 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "frmMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

'Code this form from PSC
Private Const SPI_GETWORKAREA As Long = 48&
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
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

Private m_iChangeSpeed    As Long         '/* The window's display speed
Private m_iCounter        As Long         '/* Display time in milliseconds
Private m_iScrnBottom     As Long         '/* Height of the screen - taskbar (if it is on the bottom)
Private m_bOnTop          As Boolean      '/* Form Z-Order Flag
Private m_iWindowCount    As Long         '/* Screen stop position multiplier (displaying more then 1 at a time)
Private m_bManualClose    As Boolean      '/* Manual close Flag
Private m_bCodeClose      As Boolean      '/* Prevent user close option
Private m_bFade           As Boolean      '/* Fade or move Flag
Private m_iOSver          As Byte         '/* OS 1=Win98/ME; 2=Win2000/XP
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim tg As Byte
'Dim formCap As String
'Dim mas As String
Dim co As Byte
'Dim com As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdScan_Click()
    frmPht.Show
    frmPht.txtPath = lblPath.Caption
End Sub
Private Sub Form_Load()
    Language Me
    Me.Caption = lblCaption.Caption
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
    Me.Move scrnRight - Me.Width, m_iScrnBottom, lblMes.Left + lblMes.Width + 100, cmdScan.Top + 700
    
    'Me.Move scrnRight - Me.Width, m_iScrnBottom - Me.Height, txtMas.Left + txtMas.Width + 100, cmdScan.Top + 800
    
    Timer2.Enabled = True
    tg = 10
    
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
        Me.Move Me.Left, Me.Top + 25, lblMes.Left + lblMes.Width + 100, cmdScan.Top + 700
    End If
End Sub

Private Sub Timer2_Timer()
    If Me.Top < m_iScrnBottom - Me.Height Then
        Timer.Enabled = True
        Timer2.Enabled = False
    Else
        Me.Caption = lblCaption.Caption
        Me.Move Me.Left, Me.Top - 25, lblMes.Left + lblMes.Width + 100, cmdScan.Top + 700
    End If
End Sub
