VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmDetect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detect"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   2325
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDetect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstDat 
      Height          =   1740
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   4335
   End
   Begin CtrUnicodeVN.ButtonUni cmdKill 
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3000
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
      MICON           =   "frmDetect.frx":058A
      PICN            =   "frmDetect.frx":05A6
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
   Begin CtrUnicodeVN.FrameUni frmDetect 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5106
      Caption         =   "File detect :"
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
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   2640
         Visible         =   0   'False
         Width           =   300
      End
      Begin MSComctlLib.ListView LV 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4471
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "§èi t­îng"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NhËn d¹ng"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "§­êng dÉn"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "TiÕn tr×nh"
            Object.Width           =   1764
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
   Begin CtrUnicodeVN.ButtonUni cmdBack 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
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
      MICON           =   "frmDetect.frx":0B40
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
Attribute VB_Name = "frmDetect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Sub cmdBack_Click()
    frmMain.Show
    Unload Me
End Sub
Private Sub cmdKill_Click()
Dim i As Integer
For i = 1 To LV.ListItems.Count
    If LV.ListItems.Count >= i Then
        If LV.ListItems(i).Checked = True Then
            If IsNumeric(LV.ListItems(i).SubItems(3)) = True Then
                SuspendResumeProcess CLng(LV.ListItems(i).SubItems(3)), True
                KillProcessById LV.ListItems(i).SubItems(3)
                DoEvents
            End If
        End If
    End If
Next

For i = 1 To LV.ListItems.Count
    If LV.ListItems.Count >= i Then
        If LV.ListItems(i).Checked = True Then
            If (LV.ListItems(i).Text = GetStr("DecVirus")) Or (LV.ListItems(i).Text = GetStr("DecFile")) Then
                XoaFile LV.ListItems(i).SubItems(2)
            ElseIf LV.ListItems(i).Text = GetStr("DecVir") Then
                CleanVirus LV.ListItems(i).SubItems(2), lstDat.List(i - 1)
                'Debug.Print lstDat.List(i - 1)
            End If
                DoEvents
                LV.ListItems.Remove (i)
                i = i - 1
        End If
        
    End If
Next

If SeeSta = True Then frmSta.GetStartup
ThongBao "vnAntiVirus", GetStr("MesKDec")
End Sub
Private Sub Form_Load()
    Language Me
    tb = True
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    tb = False
End Sub
Public Sub GetIDProcess()
    If LV.ListItems.Count <> 0 Then
      Dim theloop As Long
      Dim proc As PROCESSENTRY32
      Dim snap As Long
      Dim ID As Long
      Dim PathID As String
       snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
        proc.dwSize = Len(proc)
       theloop = ProcessFirst(snap, proc)
       While theloop <> 0
          ID = proc.th32ProcessID
          theloop = ProcessNext(snap, proc)
          PathID = ProcessPathByPID(ID)
          'Debug.Print PathID
          If PathID <> "SYSTEM" Then
                For i = 1 To LV.ListItems.Count
                      If PathID = LV.ListItems(i).SubItems(2) Then LV.ListItems(i).SubItems(3) = ID
                Next
            End If
       Wend
       CloseHandle snap
    End If
End Sub
