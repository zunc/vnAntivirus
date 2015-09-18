VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmSta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Starup"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmSta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin CtrUnicodeVN.ButtonUni cmdRe 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3840
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
      MICON           =   "frmSta.frx":058A
      PICN            =   "frmSta.frx":05A6
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
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   1440
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComctlLib.ImageList ima 
      Left            =   240
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3705
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
         Object.Width           =   12700
      EndProperty
   End
   Begin CtrUnicodeVN.ButtonUni cmdBack 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
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
      MICON           =   "frmSta.frx":0B40
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
      Left            =   3600
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Que1t kho73i d9o65ng"
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
      MICON           =   "frmSta.frx":0B5C
      PICN            =   "frmSta.frx":0B78
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
   Begin CtrUnicodeVN.ButtonUni cmdRepair 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Su74a chu74a key"
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
      MICON           =   "frmSta.frx":1112
      PICN            =   "frmSta.frx":112E
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
Attribute VB_Name = "frmSta"
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

Private Sub getVal(START As Key)
'Doan code nay co nguon goc tu PSC
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
            KeyPath = ChuoiGiaTri(Left(Asc(Buf2), retdata - 1))
        Else
            KeyName = Left(Buf, Ret)
            KeyPath = ChuoiGiaTri(Left(Buf2, retdata - 1))
        End If
        
            pic.Cls
            GetIcon KeyPath, pic
            ima.ListImages.Add LV.ListItems.Count + 1, , pic.Image
            Dim lsv As ListItem
                  Set lsv = LV.ListItems.Add(, , KeyName, , LV.ListItems.Count + 1)
                  lsv.SubItems(1) = KeyPath
                  
        Cnt = Cnt + 1
        Buf = Space(BUFFER_SIZE)
        Buf2 = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        retdata = BUFFER_SIZE
    Wend
    RegCloseKey Result
End Sub
Private Sub cmdBack_Click()
    frmMain.Show
    Unload Me
End Sub
Private Sub cmdRe_Click()
    GetStartup
End Sub

Private Sub cmdRepair_Click()
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", WindowsDir & "\system32\userinit.exe"
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", WindowsDir & "\explorer.exe"
    ThongBao "vnAntiVirus", GetStr("MesRCR")
    cmdRepair.Visible = False
    GetStartup
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
    cmdRepair.Visible = False
    GetStartup
    SeeSta = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SeeSta = False
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu frmMnu.mnub0
End Sub
Private Sub getValSta(START As Key)
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
                            'Debug.Print Buf
                KeyName = Left(Buf, Ret)
                If Trim(Buf2) <> "" Then KeyPath = ChuoiGiaTri(Left(Buf2, retdata - 1))
            End If
                Dim lsv As ListItem
                      Set lsv = LV.ListItems.Add()
                      lsv.Text = KeyName
                      lsv.SubItems(1) = KeyPath
                        If START = a Then
                            lsv.SubItems(2) = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                        ElseIf START = B Then
                            lsv.SubItems(2) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                        End If
                            lsv.Checked = True
        Cnt = Cnt + 1
        Buf = Space(BUFFER_SIZE)
        Buf2 = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        retdata = BUFFER_SIZE
    Wend
    RegCloseKey Result
End Sub

Public Sub GetStartup()
    ThietLap LV, ima, pic
    
        GetKeySta HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit"
        'Xu ly key Explorer
        GetKeySta HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
        'Xu ly cac key dac biet khac
        GetKeySta HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "Load"
        GetKeySta HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "Run"
        If (FileExists(LV.ListItems(1).SubItems(1)) = False) Or (FileExists(LV.ListItems(1).SubItems(1)) = False) Then
            ThongBao "Warning", GetStr("MesKDR")
            cmdRepair.Visible = True
        End If
    getValSta a
    getValSta B
    If LV.ListItems.Count <> 0 Then GetIcons LV, ima, pic

End Sub
Private Sub GetKeySta(hKey As Key, kPath As String, kName As String)
    Dim PathExp As String
    Dim t As Byte
    Dim t1 As Byte
    
        Dim tmpStr As String
        PathExp = GetString(hKey, kPath, kName)
        Do While InStr(1, PathExp, ".exe") Or InStr(1, PathExp, ".pif") Or InStr(1, PathExp, ".htm")
            t = InStr(1, PathExp, ".", vbBinaryCompare)
            tmpStr = Left(PathExp, t + 3)
            tmpStr = ChuoiGiaTri(tmpStr)
                    Set lsv1 = LV.ListItems.Add()
                    lsv1.Text = kName
                    lsv1.SubItems(1) = tmpStr
                    If hKey = B Then
                        lsv1.SubItems(2) = "HKEY_LOCAL_MACHINE\" & kPath ' & kName
                    ElseIf hKey = a Then
                        lsv1.SubItems(2) = "HKEY_CURRENT_USER\" & kPath ' & kName
                    End If
                    lsv1.Checked = True
            If Len(PathExp) >= t + 4 Then
                PathExp = Right(PathExp, Len(PathExp) - t - 4)
            Else
                PathExp = ""
            End If
        Loop
End Sub
