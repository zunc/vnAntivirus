VERSION 5.00
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   Begin CtrUnicodeVN.FrameUni frmMain 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7646
      Caption         =   "Main"
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
      Begin CtrUnicodeVN.ButtonUni cmdScanSys 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Que1t he65 tho61ng"
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
         MICON           =   "frmMain.frx":08CA
         PICN            =   "frmMain.frx":08E6
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
      Begin CtrUnicodeVN.ButtonUni cmdPro 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Tie61n tri2nh (Chu7a hoa2n thie65n)"
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
         MICON           =   "frmMain.frx":0E80
         PICN            =   "frmMain.frx":0E9C
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
      Begin CtrUnicodeVN.ButtonUni cmdSta 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Kho73i d9o65ng"
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
         MICON           =   "frmMain.frx":1436
         PICN            =   "frmMain.frx":1452
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
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Que1t vo71i ma64u"
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
         MICON           =   "frmMain.frx":19EC
         PICN            =   "frmMain.frx":1A08
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
      Begin CtrUnicodeVN.ButtonUni cmdOpt 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Tu2y cho5n"
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
         MICON           =   "frmMain.frx":1FA2
         PICN            =   "frmMain.frx":1FBE
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
      Begin CtrUnicodeVN.ButtonUni cmdAu 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Ta1c gia3"
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
         MICON           =   "frmMain.frx":2558
         PICN            =   "frmMain.frx":2574
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
      Begin CtrUnicodeVN.ButtonUni cmdHide 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "A63n"
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
         MICON           =   "frmMain.frx":2B0E
         PICN            =   "frmMain.frx":2B2A
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
      Begin CtrUnicodeVN.ButtonUni cmdUp 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Du74 lie65u worm"
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
         MICON           =   "frmMain.frx":2B46
         PICN            =   "frmMain.frx":2B62
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
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Private Sub cmdAu_Click()
    frmAbout.Show
    Unload Me
End Sub
Private Sub cmdHide_Click()
    frmMnu.Show
    frmMnu.Hide
    Unload Me
End Sub
Private Sub cmdOpt_Click()
    frmOpt.Show
    Unload Me
End Sub
Private Sub cmdPro_Click()
    frmPro.Show
    Unload Me
End Sub
Private Sub cmdScan_Click()
    frmPht.Show
    Unload Me
End Sub
Private Sub cmdScanSys_Click()
Dim sOutPut
    sOutPut = ""
    sOutPut = GetFolder(Me.hwnd, "Scan Path : ", WindowsDir)
    If sOutPut <> "" Then
        PathWScan = sOutPut
        frmScan.Show
    Else
        ThongBao "vnAntiVirus", GetStr("MesSe")
    End If
End Sub
Private Sub cmdSta_Click()
    frmSta.Show
    Unload Me
End Sub
Private Sub cmdUp_Click()
    frmDat.Show
    Unload Me
End Sub
Private Sub Form_Load()

    App.TaskVisible = False
    PathApp = App.Path
    If Right(PathApp, 1) = "\" Then PathApp = Left(Path, Len(PathApp) - 1)
If FileExists(PathApp & "\Data.ini") = False Then
    ThongBao "vnAntiVirus", GetStr("MesNS")
    Me.Show
Else
    GetOpt
    If ichkShow = True Then
        Me.Show
    Else
        Me.Hide
    End If
    If ichkSystemTray = True Then Load frmMnu
End If

If (FileExists(PathApp & "\Language\VietNam.lng") = False) Or (FileExists(PathApp & "\Language\EngLish.lng") = False) Then
    bLang = False
    ThongBao "vnAntiVirus", GetStr("MesFL")
Else
    bLang = True
    Language Me
    Language frmMnu
End If

'Tinh nang Monitor hien nay lam viec chua on dinh nen tam thoi chua co mat
'If LoadMon = False Then
'    If FileExists(PathApp & "\Mon\Mon.exe") = True Then
'        Shell PathApp & "\Mon\Mon.exe " & PathDec
'        LoadMon = True
'    Else
'        ThongBao "vnAntiVirus", GetStr("MesNF") & " " & PathApp & "\Mon\Mon.exe"
'    End If
'End If
    SeeSta = False
End Sub
