VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPht 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan all system"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPht.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPathF 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Text            =   "C:\Windows"
      Top             =   480
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CtrUnicodeVN.ButtonUni cmdKill 
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Tie6u die65t"
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
      MICON           =   "frmPht.frx":058A
      PICN            =   "frmPht.frx":05A6
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
   Begin CtrUnicodeVN.FrameUni frmPath 
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
      Caption         =   "File trong he65 tho61ng"
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
      Begin VB.PictureBox picW 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1665
         ScaleWidth      =   5865
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   5895
         Begin vnAntivirus.ucFirefoxWait ffw 
            Height          =   240
            Left            =   1920
            TabIndex        =   18
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
         End
         Begin CtrUnicodeVN.LabelUni lblWait 
            Height          =   255
            Left            =   2160
            TabIndex        =   17
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Please wait..."
            SetKieugoTV     =   1
            Appearance      =   0
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ima"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnAvant"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tªn"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "§­êng dÉn"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "KÝch th­íc"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ImageList ima 
         Left            =   0
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin CtrUnicodeVN.FrameUni frmPro 
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2990
      Caption         =   "Tie61n tri2nh"
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
      Begin MSComctlLib.ListView LV1 
         Height          =   1305
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2302
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ima1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnAvant"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tªn"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "§­êng dÉn"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ChØ sè"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ImageList ima1 
         Left            =   0
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   300
   End
   Begin CtrUnicodeVN.ButtonUni cmdScan 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmPht.frx":0B40
      PICN            =   "frmPht.frx":0B5C
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
   Begin CtrUnicodeVN.ButtonUni cmdBro 
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "..."
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
      MICON           =   "frmPht.frx":10F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "D:\Worm\Sample\RealWorm.exe"
      Top             =   120
      Width           =   4695
   End
   Begin CtrUnicodeVN.LabelUni lblPath 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
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
      Caption         =   "D9u7o72ng da64n:"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.FrameUni frmRes 
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   4680
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      Caption         =   "Kho73i d9o65ng"
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
      Begin MSComctlLib.ListView LV2 
         Height          =   1545
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ima2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnAvant"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tªn khãa"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "GÝa trÞ"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "§­êng dÉn khãa"
            Object.Width           =   12347
         EndProperty
      End
      Begin MSComctlLib.ImageList ima2 
         Left            =   0
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin CtrUnicodeVN.ButtonUni cmdBrowFolder 
      Height          =   285
      Left            =   5640
      TabIndex        =   12
      Top             =   480
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "..."
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
      MICON           =   "frmPht.frx":1112
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CtrUnicodeVN.LabelUni lblFolder 
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   480
      Width           =   735
      _ExtentX        =   1296
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
      Caption         =   "Thu7 mu5c :"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.ButtonUni cmdBack 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Tro73 la5i"
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
      MICON           =   "frmPht.frx":112E
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
Attribute VB_Name = "frmPht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Option Explicit
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1
'---Tim dung luong------
Const GENERIC_READ = &H80000000
Const FILE_SHARE_READ = &H1
Const OPEN_EXISTING = 3
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 Dim CRC As String
'-----------------------------------------
'Private m_CRC As clsCRC
'-----------------------------------------
Private Scaning As Boolean
Private Size As Double

Public Function DungLuong(DuongDan As String) As Long
Dim hFile As Long, nSize As Currency
    hFile = CreateFile(DuongDan, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    GetFileSizeEx hFile, nSize
    CloseHandle hFile
DungLuong = nSize * 10000
End Function

Private Sub cmdBack_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdBro_Click()
    cd.DialogTitle = "Choose a file ..."
    cd.Filter = "Protable files (*.pif;*.exe)|*.exe;*.pif|All Files (*.*)|*.*"
    Dim Path As String
    cd.ShowOpen
    If cd.Filename <> "" Then txtPath.Text = cd.Filename
End Sub

Private Sub cmdBrowFolder_Click()

Dim sOutPut
    sOutPut = ""
    sOutPut = GetFolder(Me.hWnd, "Scan Path : ", txtPathF.Text)
    If sOutPut <> "" Then
        txtPathF.Text = sOutPut
    Else
        If Len(txtPath.Text) <> 0 Then
            Else
            ThongBao "vnAntiVirus", GetStr("MesSe")
    End If
    End If
End Sub

Private Sub cmdKill_Click()

Dim i As Byte
'Kill process
For i = 1 To LV1.ListItems.Count
    If LV1.ListItems.Count >= i Then
        If LV1.ListItems(i).Checked = True Then
            SuspendResumeProcess CLng(LV1.ListItems(i).SubItems(2)), True
            KillProcessById LV1.ListItems(i).SubItems(2)
            LV1.ListItems.Remove (i)
            i = i - 1
        End If
    End If
Next
ThongBao "vnAntivirus", GetStr("MesKPro")

'Xoa file trong he thong
For i = 1 To LV.ListItems.Count
    If LV.ListItems.Count >= i Then
        If LV.ListItems(i).Checked = True Then
            If FileExists(LV.ListItems(i).SubItems(1)) = True Then
                XoaFile LV.ListItems(i).SubItems(1)
                
                If FileExists(LV.ListItems(i).SubItems(1)) = True Then
                    'Chu y : Dong code ben duoi dac biet nguy hiem neu khong tieu diet hoan toan cac process
                    'chuong trinh se roi vao vong lap vo tan neu ko can than
                    i = i - 1
                Else
                    LV.ListItems.Remove (i)
                    i = i - 1
                End If
                
            End If
        End If
    End If
Next
ThongBao "vnAntivirus", GetStr("MesKFile")
'Xoa key trong regedit
For i = 1 To LV2.ListItems.Count
    If LV2.ListItems.Count >= i Then
        If LV2.ListItems(i).Checked = True Then
            If Left(LV2.ListItems(i).ListSubItems(2), 18) = "HKEY_LOCAL_MACHINE" Then
                DelSetting HKEY_LOCAL_MACHINE, Right(LV2.ListItems(i).ListSubItems(2), Len(LV2.ListItems(i).ListSubItems(2)) - 19), LV2.ListItems(i).Text
            ElseIf Left(LV2.ListItems(i).ListSubItems(2), 17) = "HKEY_CURRENT_USER" Then
                DelSetting HKEY_CURRENT_USER, Right(LV2.ListItems(i).ListSubItems(2), Len(LV2.ListItems(i).ListSubItems(2)) - 18), LV2.ListItems(i).Text
            End If
                LV2.ListItems.Remove (i)
                i = i - 1
        End If
    End If
Next

    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", WindowsDir & "\system32\userinit.exe"
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", WindowsDir & "\explorer.exe"
ThongBao "vnAntivirus", GetStr("MesKRegedit")
End Sub
Private Sub cmdScan_Click()
If Scaning = False Then
    If FileExists(txtPath.Text) = True Then
        Scaning = True
        cmdScan.Caption = GetStr("MesScanT")
        CRC = Hex$(m_CRC.CalculateFile(txtPath.Text))
        'Debug.Print CRC
'Tien hanh quet trong tat ca cac tien trinh (Process)
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
  Dim exename As String
  Dim ID As Long
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)

    ThietLap LV1, ima1, Pic
   While theloop <> 0

      ID = proc.th32ProcessID
      theloop = ProcessNext(snap, proc)
      If FileExists(ProcessPathByPID(proc.th32ProcessID)) = True Then
            If CRC = Hex$(m_CRC.CalculateFile(ProcessPathByPID(proc.th32ProcessID))) Then
                    Dim lsv As ListItem
                    Set lsv = LV1.ListItems.Add()
                    lsv.Text = proc.szExeFile
                    lsv.SubItems(1) = ProcessPathByPID(proc.th32ProcessID)
                    lsv.SubItems(2) = proc.th32ProcessID
                    lsv.Checked = True
            End If
        End If
        
   Wend
   CloseHandle snap
      
      If LV1.ListItems.Count <> 0 Then GetIcons LV1, ima1, Pic
ThongBao "vnAntivirus", GetStr("MesLPro")

'Tien hanh tim cac file trong he thong
        ThietLap LV, Ima, Pic
        picW.Visible = True
        ffw.PlayWait
        Size = DungLuong(txtPath.Text)
    Set SP = New cScanPath
        LV.ListItems.Clear
        
        With SP
            .Archive = True
            .Compressed = True
            .Hidden = True
            .Normal = True
            .ReadOnly = True
            .System = True
            
            .Filter = "*.exe;*.pif"
            
            .StartScan txtPathF.Text, True, True
        End With
        

        If LV.ListItems.Count <> 0 Then GetIcons LV, Ima, Pic
        picW.Visible = False
ThongBao "vnAntivirus", GetStr("MesLFile")

'--------------------------------------------------
'Tien hanh quet tat ca cac key khoi dong trong Regedit
   
    ThietLap LV2, ima2, Pic
    
    Dim PathInit As String
    Dim PathExp As String
    Dim t As Byte
    Dim t1 As Byte
    
    Dim lsv1 As ListItem
    
    PathInit = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "Userinit")
    'Xu lu key Userinit
        If CRC = Hex$(m_CRC.CalculateFile(PathInit)) Then
                    Set lsv1 = LV2.ListItems.Add()
                    lsv1.Text = "Userinit"
                    lsv1.SubItems(1) = PathInit
                    lsv1.SubItems(2) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\Userinit"
                    lsv1.Checked = True
        End If
        'Xu ly key Explorer

        Dim tmpStr As String
        PathExp = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "Shell")
        Do While InStr(1, PathExp, ".exe") Or InStr(1, PathExp, ".pif")
            t = InStr(1, PathExp, ".", vbBinaryCompare)
            tmpStr = Left(PathExp, t + 3)
            tmpStr = ChuoiGiaTri(tmpStr)
                
                If CRC = Hex$(m_CRC.CalculateFile(tmpStr)) Then
                    Set lsv1 = LV2.ListItems.Add()
                    lsv1.Text = "Shell"
                    lsv1.SubItems(1) = tmpStr
                    lsv1.SubItems(2) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\Shell"
                    lsv1.Checked = True
                End If
            If Len(PathExp) >= t + 4 Then
                PathExp = Right(PathExp, Len(PathExp) - t - 4)
            Else
                PathExp = ""
            End If
        Loop
        
    getVal a
    getVal b
    
    If LV2.ListItems.Count <> 0 Then GetIcons LV2, ima2, Pic
    ThongBao "vnAntivirus", GetStr("MesLRegedit")
    cmdScan.Caption = GetStr("MesScanF")
    Scaning = False
    Else
    Scaning = False
    ThongBao "vnAntivirus", GetStr("MesNF")
    End If
Else
    cmdScan.Caption = GetStr("MesScanF")
    SP.StopScan
    DoEvents
    Scaning = False
    ThongBao "vnAntivirus", GetStr("MesStopScan")
End If

End Sub

Private Sub Form_Load()
        Language Me
        Scaning = False
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
End Sub
Private Sub SP_DirMatch(Directory As String, Path As String)
    'This Event fires for each Folder found
End Sub

Private Sub SP_FileMatch(Filename As String, Path As String)
    'This Event fires for each File found
    'lst.AddItem Path & Filename
If Size = DungLuong(Path & Filename) Then
    With LV
        If CRC = Hex$(m_CRC.CalculateFile(Path & Filename)) Then
                Dim lsv As ListItem
                    Set lsv = LV.ListItems.Add()
                    lsv.Text = Filename
                    lsv.SubItems(1) = Path & Filename
                    lsv.SubItems(2) = DungLuong(Path & Filename)
                    lsv.Checked = True
        End If
    End With
End If
End Sub
Private Sub getVal(START As Key)
'Sorry cac ban nhe, Dung Coi nhac qua nen copy y chang lai doan code nay
Dim Cnt As Long, Buf As String, Buf2 As String, retdata As Long, typ As Long
    'List1.Clear
    Dim KeyName As String
    Dim KeyPath As String
    Buf = Space(BUFFER_SIZE)
    Buf2 = Space(BUFFER_SIZE)
    Ret = BUFFER_SIZE
    retdata = BUFFER_SIZE

    Cnt = 0
    RegOpenKeyEx START, Pathkey, 0, KEY_ALL_ACCESS, Result
    While RegEnumValue(Result, Cnt, Buf, Ret, 0, typ, ByVal Buf2, retdata) <> ERROR_NO_MORE_ITEMS
        If typ = REG_DWORD Then
            KeyName = Left(Buf, Ret)
            If Trim(Buf2) <> "" Then KeyPath = ChuoiGiaTri(Left(Asc(Buf2), retdata - 1))
        Else
            KeyName = Left(Buf, Ret)
            If Trim(Buf2) <> "" Then KeyPath = ChuoiGiaTri(Left(Buf2, retdata - 1))
        End If
        
        If CRC = Hex$(m_CRC.CalculateFile(KeyPath)) Then
        
            
                'Pic.Cls
                'GetIcon KeyPath, Pic
                'ima2.ListImages.Add LV2.ListItems.Count + 1, , Pic.Image
                Dim lsv As ListItem
                      Set lsv = LV2.ListItems.Add()
                      lsv.Text = KeyName
                      lsv.SubItems(1) = KeyPath
                        If START = a Then
                            lsv.SubItems(2) = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                        ElseIf START = b Then
                            lsv.SubItems(2) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                        End If
                            lsv.Checked = True
            End If
        Cnt = Cnt + 1
        Buf = Space(BUFFER_SIZE)
        Buf2 = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        retdata = BUFFER_SIZE
    Wend
    RegCloseKey Result
End Sub
