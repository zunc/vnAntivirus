VERSION 5.00
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin vnAntivirus.XP_ProgressBar pro 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   6956042
      Orientation     =   1
      Scrolling       =   3
   End
   Begin VB.ListBox lstStr 
      Height          =   1035
      Left            =   1920
      TabIndex        =   18
      Top             =   3240
      Width           =   855
   End
   Begin VB.ListBox lstSDec 
      Height          =   1035
      Left            =   2760
      TabIndex        =   17
      Top             =   3240
      Width           =   855
   End
   Begin CtrUnicodeVN.CheckBoxUni chkIndex 
      Height          =   240
      Left            =   0
      TabIndex        =   16
      Top             =   1800
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   423
      Value           =   1
      Caption         =   "Ta5o ba3n chi3 mu5c"
      Pic_UncheckedNormal=   "frmScan.frx":058A
      Pic_CheckedNormal=   "frmScan.frx":07E4
      Pic_MixedNormal =   "frmScan.frx":0A3E
      Pic_UncheckedDisabled=   "frmScan.frx":0C98
      Pic_CheckedDisabled=   "frmScan.frx":0EF2
      Pic_MixedDisabled=   "frmScan.frx":114C
      Pic_UncheckedOver=   "frmScan.frx":13A6
      Pic_CheckedOver =   "frmScan.frx":1600
      Pic_MixedOver   =   "frmScan.frx":185A
      Pic_UncheckedDown=   "frmScan.frx":1AB4
      Pic_CheckedDown =   "frmScan.frx":1D0E
      Pic_MixedDown   =   "frmScan.frx":1F68
      SetKieugoTV     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CtrUnicodeVN.CheckBoxUni chkCWI 
      Height          =   240
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   423
      Value           =   0
      Caption         =   "Que1t theo chi3 mu5c d9a4 sa81p xe61p"
      Pic_UncheckedNormal=   "frmScan.frx":21C2
      Pic_CheckedNormal=   "frmScan.frx":241C
      Pic_MixedNormal =   "frmScan.frx":2676
      Pic_UncheckedDisabled=   "frmScan.frx":28D0
      Pic_CheckedDisabled=   "frmScan.frx":2B2A
      Pic_MixedDisabled=   "frmScan.frx":2D84
      Pic_UncheckedOver=   "frmScan.frx":2FDE
      Pic_CheckedOver =   "frmScan.frx":3238
      Pic_MixedOver   =   "frmScan.frx":3492
      Pic_UncheckedDown=   "frmScan.frx":36EC
      Pic_CheckedDown =   "frmScan.frx":3946
      Pic_MixedDown   =   "frmScan.frx":3BA0
      SetKieugoTV     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   1080
      Width           =   4575
   End
   Begin CtrUnicodeVN.LabelUni lblPathScan 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Thu7 mu5c que1t :"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   3240
   End
   Begin VB.ListBox lstName 
      Height          =   1035
      Left            =   840
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.ListBox lstDat 
      Height          =   1035
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin CtrUnicodeVN.ButtonUni cmdCancel 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmScan.frx":3DFA
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
   Begin CtrUnicodeVN.LabelUni lblDQ 
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
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
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblNP 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "D9ang que1t :"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblDaQuet 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "D9a4 que1t:"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblPhH 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "D9a4 pha1t hie65n :"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblPH 
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "0"
      Appearance      =   0
   End
   Begin CtrUnicodeVN.ButtonUni cmdScan 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
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
      MICON           =   "frmScan.frx":3E16
      PICN            =   "frmScan.frx":3E32
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
   Begin CtrUnicodeVN.LabelUni lblTime 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Tho72i gian:"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblTG 
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblPathText 
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
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
   Begin vnAntivirus.ucFirefoxWait ffw 
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
End
Attribute VB_Name = "frmScan"
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
Public dq As Long
Public tg As Long
Dim FullPath As String
Dim Scaning As Boolean

Dim iIndex As Boolean
Dim iCWI As Boolean
Dim strTMP As String
Dim DemFile As Long
Dim i As Long
'Private m_CRC As clsCRC

Private Sub ThietLapForm()
    'pro.Color = RGB(226, 233, 246)
    'GetData App.Path & "\Dat\Sign.vnd", lstDat, lstName, True
    'GetData App.Path & "\Dat\SW\YourSign.vnd", lstDat, lstName, False
    'GetData App.Path & "\Dat\String.vnd", lstStr, lstSDec, True

'Note:    'vnd= VnAntivirus data ;-)
    ' dung co dich la "Viet Nam Dong" nhe, Dung Coi ko tham lam tien ...lam dau
    
End Sub

Private Sub chkCWI_Click()
    If chkCWI.Value = Checked Then chkIndex.Value = Unchecked
End Sub
Private Sub chkIndex_Click()
    If chkIndex.Value = Checked Then chkCWI.Value = Unchecked
End Sub
Private Sub cmdCancel_Click()
    frmMain.Show
    Unload Me
End Sub
Private Sub cmdScan_Click()
    'On Error Resume Next

If (Scaning = False) Then
    cmdCancel.Enabled = False
    'frmDetect.LV.Visible = False
    Scaning = True
    dq = 0
    ph = 0
    tg = 0
    DemFile = 0
    'MsgBox strTmp & "\Index.vnd"
    Timer.Enabled = True
    If chkIndex.Value = Checked Then iIndex = True
    If chkCWI.Value = Checked Then GoTo CWI
    ffw.PlayWait
    lblPH.Caption = "0 file"
    cmdScan.Caption = GetStr("MesScanT")
    cmdCancel.Enabled = False
    Scaning = True
    'dq = 0
    'ph = 0
    'tg = 0
    Timer.Enabled = True
    Set SP = New cScanPath
        
        If FileExists(PathApp & "\indexTmp.vnd") = True Then XoaFile PathApp & "\indexTmp.vnd"
        
        With SP
            .Archive = True
            .Compressed = True
            .Hidden = True
            .Normal = True
            .ReadOnly = True
            .System = True
            
            .Filter = "*.exe;*.pif;*.com;*.vbs;*.bat;*.asp;*.bin;*.chm;*.cpl;*.dll;*.eml;*.hta;*.htm;*.mht;*.ocx;*.url"
            '"*.exe;*.pif;*.com;*.vbs;*.bat;*.asp;*.bin;*.chm;*.cpl;*.dll;*.drv;*.eml;*.hta;*.htm*.drv;*.mht;*.mp3;*.ocx;*.sys;*.url"
            .StartScan PathWScan, True, True
        End With
        'Debug.Print "Okie 2"
        'MsgBox "Xong liet ke"
        'Okie, sau khi da tao bang chi muc file xong
        
    If Right(PathWScan, 1) = "\" Then PathWScan = Left(PathWScan, Len(PathWScan) - 1)
    If iIndex = True Then
        If FileExists(PathWScan & "\Index.vnd") = True Then XoaFile PathWScan & "\Index.vnd"
        FileCopy PathApp & "\indexTmp.vnd", PathWScan & "\Index.vnd"
        WriteINI PathWScan & "\Index.vnd", "Info", "File", CStr(DemFile)
    End If
    
ScanIndex PathApp & "\indexTmp.vnd"

        Timer.Enabled = False
        cmdCancel.Enabled = True
        Scaning = False
        ffw.StopWait
        cmdScan.Caption = GetStr("MesScanF")
        ThongBao "vnAntiVirus", GetStr("MesComScan")
        lblDQ.Caption = dq
Else
    SP.StopScan
    cmdScan.Caption = GetStr("MesScanF")
    cmdCancel.Enabled = True
    Scaning = False
    ffw.StopWait
    ThongBao "vnAntiVirus", GetStr("MesStoScan")
End If
Exit Sub
CWI:
    dq = 0
    ph = 0
    DemFile = Val(ReadINI(strTMP & "\Index.vnd", "Info", "File"))
    ScanIndex strTMP & "\Index.vnd"
    ThongBao "vnAntiVirus", GetStr("MesComScan")
End Sub
Private Sub Form_Load()
    Language Me
    Scaning = False
    lblPathText.Caption = PathWScan
    If Right(PathWScan, 1) = "\" Then
        strTMP = Left(PathWScan, Len(PathWScan) - 1)
    Else
        strTMP = PathWScan
    End If
    If FileExists(strTMP & "\index.vnd") = True Then
        chkCWI.Visible = True
    Else
        chkCWI.Visible = False
    End If
    iIndex = False
    iCWI = False
    ThietLapForm
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32

End Sub
Private Sub SP_FileMatch(Filename As String, Path As String)
'Dung luong tap tin duoc quet se nho hon 6000000 byte (Gan 6Mb)
If DungLuong(Path & Filename) > 6000000 Then GoTo KetThuc
    Dim KetQua As String
    FullPath = Path & Filename
    KetQua = Hex$(m_CRC.CalculateFile(FullPath))
        If Len(KetQua) < 8 Then
            Select Case Len(KetQua)
            'Dung coi nghi, viec xet truong hop se nhanh hon viec su dung For
                Case 7
                    KetQua = "0" & KetQua
                Case 6
                    KetQua = "00" & KetQua
                Case 5
                    KetQua = "000" & KetQua
                Case 4
                    KetQua = "0000" & KetQua
                   Case 3
                    KetQua = "00000" & KetQua
                Case 2
                    KetQua = "000000" & KetQua
                Case 4
                    KetQua = "0000000" & KetQua
            End Select
        End If
    AddToFile KetQua & "|" & Path & Filename, PathApp & "\indexTMP.vnd"
    DemFile = DemFile + 1
KetThuc:
End Sub
Private Sub Timer_Timer()
    'tg = tg + 1
    'lblTG.Caption = tg
    lblDQ.Caption = dq & " file"
    txtPath.Text = FullPath
End Sub
Public Sub ScanIndex(PathFileIndex As String)
    On Error Resume Next
Dim DatTmp As String
Dim CRCTmp As String
Dim strPath As String

Open PathFileIndex For Input As #2
    Do While Not EOF(2)
        Line Input #2, DatTmp
        If dq >= DemFile Then GoTo TheEnd
        'Dong lenh tren nham trach loi khi quet file index.vnd trong thu muc
        dq = dq + 1
        pro.Value = Int(dq / DemFile * pro.Max)
        CRCTmp = Split(DatTmp, "|", , vbBinaryCompare)(0)
        strPath = Split(DatTmp, "|", , vbBinaryCompare)(1)
    'Thong qua Test, Dung Coi xac dinh duoc rang, neu thuc hien Check Icon thi toc do quet se giam 1.5 lan
    'ScanFile Path & Filename, lstDat, lstName, lstStr, lstSDec, True, True, True, False, frmMnu.Ima, frmMnu.Pic, frmMnu.pic1

        If ScanCRCMain(CRCTmp, strPath) = True Then GoTo KetThucByCRC
        DoEvents
        
        'Tien hanh check virus qua chuoi String
        Dim BoDem As String
            Open strPath For Binary As #1
                BoDem = Space(LOF(1))
                Get #1, , BoDem
            Close #1
        For i = 0 To frmMnu.lstStr.ListCount - 1
            If InStr(1, BoDem, frmMnu.lstStr.List(i), vbBinaryCompare) <> 0 Then
            'Yeah, da nhan ra virus roi nhe
            'Nhan dang theo ky thuat nhan dang chuoi String
                Detect GetStr("DecFile"), frmMnu.lstSDec.List(i), strPath
                ph = ph + 1
                lblPH.Caption = ph
                GoTo KetThuc
            End If
        Next
            'Nhan dang virus thong qua chuoi string (Khong the xac dinh virus qua CRC)
        For i = 0 To frmMnu.lstSVir.ListCount - 1
            If InStr(1, BoDem, frmMnu.lstSVir.List(i), vbBinaryCompare) <> 0 Then
                Detect GetStr("DecVir"), frmMnu.lstVirNa.List(i), strPath, frmMnu.lstVirDat.List(i)
                ph = ph + 1
                lblPH.Caption = ph
                GoTo KetThuc
            End If
       Next
       
       BoDem = vbNullString

KetThuc:
    BoDem = vbNullString
KetThucByCRC:
    Loop
TheEnd:
Close #2
frmDetect.GetIDProcess


End Sub

