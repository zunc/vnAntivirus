VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   2040
      Top             =   4080
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   2040
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   300
   End
   Begin MSComctlLib.ImageList ima 
      Left            =   1320
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3945
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6959
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
         Text            =   "Tªn tiÕn tr×nh"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "§­êng dÉn tiÕn tr×nh"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ChØ sè"
         Object.Width           =   1676
      EndProperty
   End
   Begin CtrUnicodeVN.ButtonUni cmdRe 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "La2m tu7o7i"
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
      MICON           =   "frmPro.frx":058A
      PICN            =   "frmPro.frx":05A6
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
   Begin CtrUnicodeVN.ButtonUni cmdBack 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
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
      MICON           =   "frmPro.frx":0B40
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
      Left            =   3720
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Que1t tie61n tri2nh"
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
      MICON           =   "frmPro.frx":0B5C
      PICN            =   "frmPro.frx":0B78
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
Attribute VB_Name = "frmPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Private Sub cmdBack_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdRe_Click()
    GetProcess LV, ima, pic
End Sub

Private Sub cmdScan_Click()
Dim i As Integer
With frmMnu
    For i = 1 To LV.ListItems.Count
        If FileExists(LV.ListItems(i).SubItems(1)) = True Then
            ScanFile LV.ListItems(i).SubItems(1), True, True, True, .ima, .pic, .pic1
            frmDetect.GetIDProcess
        End If
    Next
End With
    ThongBao "vnAntiVirus", GetStr("MesComScan")
End Sub
Private Sub Form_Load()
    Language Me
    GetProcess LV, ima, pic
End Sub
Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu frmMnu.mnuc0
End Sub
Private Sub Timer_Timer()
  Dim i As Integer
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
    i = 0
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
   While theloop <> 0

      theloop = ProcessNext(snap, proc)
      If ProcessPathByPID(proc.th32ProcessID) <> "SYSTEM" Then
              i = i + 1
              If i > LV.ListItems.Count Then GoTo KetThuc
            If LV.ListItems(i).SubItems(1) <> ProcessPathByPID(proc.th32ProcessID) Then GoTo KetThuc
      End If
   Wend
   CloseHandle snap
Exit Sub
KetThuc:
    GetProcess LV, ima, pic
End Sub
