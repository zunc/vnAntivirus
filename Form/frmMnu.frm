VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMnu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstPro 
      Height          =   450
      Left            =   0
      TabIndex        =   19
      Top             =   2280
      Width           =   495
   End
   Begin VB.ListBox lstVirDat 
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   2040
      Width           =   615
   End
   Begin VB.ListBox lstVirNa 
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   2040
      Width           =   615
   End
   Begin VB.ListBox lstSVir 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   735
   End
   Begin VB.ListBox lstTMP 
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   1560
      Width           =   615
   End
   Begin VB.ListBox lstIE 
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   1560
      Width           =   615
   End
   Begin VB.ListBox lstSDec 
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1560
      Width           =   375
   End
   Begin VB.ListBox lstStr 
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   1560
      Width           =   375
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   0
      Pattern         =   "*.exe;*.pif"
      System          =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
   Begin VB.ListBox lstDat 
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.ListBox lstNa 
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.FileListBox File 
      Height          =   285
      Left            =   1800
      Pattern         =   "*.ico"
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picLoad 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   1800
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   360
      Width           =   300
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2280
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2640
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   0
      Width           =   300
   End
   Begin VB.ListBox lstFile 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.ListBox lstStaTMP 
      Height          =   840
      Left            =   6360
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.ListBox lstSta 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.Timer Timer 
      Interval        =   1500
      Left            =   1320
      Top             =   120
   End
   Begin CtrUnicodeVN.MenuXpUni Menu 
      Left            =   120
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   13
      Bmp:1           =   "frmMnu.frx":0000
      Key:1           =   "#mnua1"
      Bmp:2           =   "frmMnu.frx":0428
      Key:2           =   "#mnua2"
      Bmp:3           =   "frmMnu.frx":0850
      Key:3           =   "#mnua3"
      Bmp:4           =   "frmMnu.frx":0C78
      Key:4           =   "#mnua4"
      Bmp:5           =   "frmMnu.frx":10A0
      Key:5           =   "#mnuaExit"
      Bmp:6           =   "frmMnu.frx":14C8
      Key:6           =   "#mnub2"
      Bmp:7           =   "frmMnu.frx":18F0
      Key:7           =   "#mnub1"
      Bmp:8           =   "frmMnu.frx":1D18
      Key:8           =   "#mnub3"
      Bmp:9           =   "frmMnu.frx":2140
      Key:9           =   "#mnuc2"
      Bmp:10          =   "frmMnu.frx":2568
      Key:10          =   "#mnud1"
      Bmp:11          =   "frmMnu.frx":2990
      Key:11          =   "#mnue1"
      Bmp:12          =   "frmMnu.frx":2DB8
      Key:12          =   "#mnuc1"
      Bmp:13          =   "frmMnu.frx":31E0
      Key:13          =   "#mnuc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UserSelectedMenuBackColour=   13040639
      UserSelectedMenuBorderColour=   0
      UserTopMenuBackColour=   16761765
      UserTopMenuSelectedColour=   16769990
      UserMenuBorderColour=   5910544
      UserCheckBackColour=   8108783
      UserCheckBorderColour=   5910544
      UserGradientOne =   16777215
      UserGradientTwo =   16761765
      UserSelectedItemForeColour=   0
      SetKieugoTV     =   1
      UseMenuSystemVn =   0   'False
   End
   Begin vnAntivirus.cSysTray cSysTray1 
      Height          =   510
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmMnu.frx":3608
      TrayTip         =   "vnAntiVirus 0.5 (Alpha)"
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ima 
      Left            =   3360
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMnu.frx":3BA2
            Key             =   "RealWorm"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMnu.frx":3EF4
            Key             =   "WinFile"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Menu mnua0 
      Caption         =   "Main"
      Begin VB.Menu mnua1 
         Caption         =   "Tie61n tri2nh"
      End
      Begin VB.Menu mnua2 
         Caption         =   "Kho73i d9o65ng"
      End
      Begin VB.Menu mnuaSpa 
         Caption         =   "-"
      End
      Begin VB.Menu mnua3 
         Caption         =   "Hie65n"
      End
      Begin VB.Menu mnua4 
         Caption         =   "Ta1c gia3"
      End
      Begin VB.Menu mnuaSpa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaExit 
         Caption         =   "Thoa1t"
      End
   End
   Begin VB.Menu mnub0 
      Caption         =   "Regedit"
      Begin VB.Menu mnub1 
         Caption         =   "Ta5o kho1a"
      End
      Begin VB.Menu mnub2 
         Caption         =   "Xo1a kho1a"
      End
      Begin VB.Menu mnubSpa 
         Caption         =   "-"
      End
      Begin VB.Menu mnub3 
         Caption         =   "Que1t he65 tho61ng"
      End
   End
   Begin VB.Menu mnuc0 
      Caption         =   "Process"
      Begin VB.Menu mnuc1 
         Caption         =   "Die65t tie61n tri2nh"
      End
      Begin VB.Menu mnuc2 
         Caption         =   "App mo71i"
      End
      Begin VB.Menu mnucSpa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuc3 
         Caption         =   "Que1t he65 tho61ng"
      End
   End
   Begin VB.Menu mnud0 
      Caption         =   "Add worm"
      Begin VB.Menu mnud1 
         Caption         =   "Xo1a"
      End
   End
End
Attribute VB_Name = "frmMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

'Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1

Dim iewindow As InternetExplorer
Private currentwindows     As New ShellWindows
Dim qc As Boolean
Dim ccTMP As String
'Private m_CRC As clsCRC
Private Sub cSysTray1_MouseDblClick(Button As Integer, ID As Long)
    frmMain.Show
End Sub
Private Sub cSysTray1_MouseDown(Button As Integer, ID As Long)
    PopupMenu frmMnu.mnua0
End Sub
Private Sub cSysTray1_MouseUp(Button As Integer, ID As Long)
    PopupMenu frmMnu.mnua0
End Sub
Private Sub Form_Load()
    Language Me
    tb = False
    Me.Hide
    ThietLapForm
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
End Sub
Private Sub Form_Unload(Cancel As Integer)
    cSysTray1.InTray = False
End Sub
Private Sub mnua1_Click()
    frmPro.Show
End Sub
Private Sub mnua2_Click()
    frmSta.Show
End Sub
Private Sub mnua3_Click()
    frmMain.Show
End Sub

Private Sub mnua4_Click()
    frmAbout.Show
End Sub
Private Sub mnuaExit_Click()
    Dim i As Form
    For Each i In Forms
        Unload i
    Next
End Sub
Private Sub mnub1_Click()
    frmAddKey.Show
    Unload frmSta
End Sub

Private Sub mnub2_Click()

With frmSta
If Left(.LV.SelectedItem.ListSubItems(2), Len("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon")) <> "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" Then
    If Left(.LV.SelectedItem.ListSubItems(2), 18) = "HKEY_LOCAL_MACHINE" Then
        DelSetting HKEY_LOCAL_MACHINE, Right(.LV.SelectedItem.ListSubItems(2), Len(.LV.SelectedItem.ListSubItems(2)) - 19), .LV.SelectedItem.Text
    ElseIf Left(.LV.SelectedItem.ListSubItems(2), 17) = "HKEY_CURRENT_USER" Then
        DelSetting HKEY_CURRENT_USER, Right(.LV.SelectedItem.ListSubItems(2), Len(.LV.SelectedItem.ListSubItems(2)) - 18), .LV.SelectedItem.Text
    End If
        .GetStartup
Else
    ThongBao "vnAntiVirus", GetStr("MesCantDelKey")
End If
End With
End Sub
Private Sub mnub3_Click()
    frmPht.Show
    frmPht.txtPath = frmSta.LV.SelectedItem.ListSubItems(1)
    Unload frmSta
End Sub
Private Sub mnuc1_Click()
    KillProcessById frmPro.LV.SelectedItem.ListSubItems(2)
    DoEvents
    GetProcess frmPro.LV, frmPro.Ima, frmPro.Pic
End Sub

Private Sub mnuc2_Click()
    frmNewApp.Show
End Sub
Private Sub mnuc3_Click()
    frmPht.Show
    frmPht.txtPath = frmPro.LV.SelectedItem.ListSubItems(1)
    Unload frmSta
End Sub
Private Sub mnud1_Click()
If frmDat.LVI.ListItems.Count <> 0 Then
    XoaFile PathApp & "\Dat\Icon\" & frmDat.LVI.SelectedItem.Text & ".ico"
    frmDat.GetInfo
End If
End Sub
Private Sub getSta(START As Key)
'Sorry cac ban nhe, Dung Coi nhac qua nen copy y chang lai doan code nay
Dim Cnt As Long, Buf As String, Buf2 As String, retdata As Long, typ As Long

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
            KeyPath = Left(Asc(Buf2), retdata - 1)
        Else
            KeyName = Left(Buf, Ret)
            KeyPath = Left(Buf2, retdata - 1)
        End If
                    
                        If START = a Then
                            lstSta.AddItem "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" & KeyName & "|" & KeyPath
                        ElseIf START = B Then
                            lstSta.AddItem "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" & KeyName & "|" & KeyPath
                        End If
                        
        Cnt = Cnt + 1
        Buf = Space(BUFFER_SIZE)
        Buf2 = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        retdata = BUFFER_SIZE
    Wend
    RegCloseKey Result
End Sub
Private Function Find(Path As String, LstListBox As ListBox) As String
    Dim i As Byte
    Find = "0"
    For i = 0 To LstListBox.ListCount - 1
        If Path = LstListBox.List(i) Then
            Find = "OK"
        ElseIf (Path <> LstListBox.List(i)) And (Split(Path, "|", , vbBinaryCompare)(0) = Split(LstListBox.List(i), "|", , vbBinaryCompare)(0)) Then
            Find = "kc|" & Path & "|" & LstListBox.List(i)
        Else
            
        End If
    Next
End Function
Private Sub SysInfo_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
    If DeviceType <> 2 Then Exit Sub
    If DeviceType = 2 Then
        If ichkUSB = True Then
            CheckThuMuc GetUSB(DeviceID) & ":"
            
            Set SP = New cScanPath
            With SP
                .Archive = True
                .Compressed = True
                .Hidden = True
                .Normal = True
                .ReadOnly = True
                .System = True
            
                .Filter = "*.exe;*.pif"
            
                .StartScan GetUSB(DeviceID) & ":", True, True
            End With
            ThongBao "vnAV", GetStr("MesSUSB")
            SP.StopScan
            Exit Sub
        End If
    End If

End Sub
Private Function GetUSB(dev_id As Long) As String
    drives = 1
        For i = 0 To 25
            If drives = dev_id Then
                GetUSB = Chr(i + 65)
                Exit Function
            End If
        drives = drives * 2
        Next i
End Function
Private Function CheckSize(Size As Double, ListBoxSize As ListBox) As Integer
Dim i As Integer
CheckSize = 0
For i = 0 To ListBoxSize.ListCount - 1
    If Size = CDbl(ListBoxSize.List(i)) Then CheckSize = i + 1
Next
End Function
Private Sub SP_FileMatch(Filename As String, Path As String)
    Dim tmp As String
    tmp = IIf(Len(Path) = 2, Path & "\", Path)
    'Debug.Print tmp & Filename
    ScanFile tmp & Filename, ichkScanI, ichkAutoIT, True, Ima, Pic, pic1
    'Dim strTMP As String
   '
   ' If ichkScanI = True Then
   '     strTMP = SoSanhImage(Path & Filename, pic, pic1, ima)
   '     If strTMP <> "0" Then Detect Filename, "Worm " & strTMP, Path & Filename
   ' End If
   '
   ' If ichkAutoIT = True Then
   '         Dim BoDem As String
   '         BoDem = ""
   '         Open Path & Filename For Binary As #1
   '             BoDem = Space(LOF(1))
   '             Get #1, , BoDem
   '         Close #1
            'Dong lenh ben duoi kha la au, tuy nhien co the "Chap nhan duoc" (hi hi)
            'Neu ban muon hieu tai sao thi hay thu phan tich 1 file viet tren AutoIT nhe
   '         If InStr(1, BoDem, "AutoIt", vbBinaryCompare) <> 0 Then Detect Filename, "File write on AutoIT ", Path & Filename
   ' End If
   '             BoDem = ""
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Can phai sua chua
    'ScanFile Path & Filename, lstDat, lstNa
End Sub
Private Sub Timer_Timer()
'Check thuong truc thu muc duoc duyet

'Sorry nhe, vi trinh do thap kem nen Dung Coi phai su dung rat nhieu Control
                   
Dim buffer, ValidData As String
Dim C As Collection
Dim currentlocation As String
Dim CRC As String
Dim k As Integer
lstTMP.Clear
    On Error Resume Next

    Timer.Enabled = False
    For Each iewindow In currentwindows
        
        DoEvents
        If iewindow.Busy Then
            GoTo busysignal
        End If

        currentlocation = iewindow.LocationURL
        lstTMP.AddItem currentlocation

        For k = 0 To lstIE.ListCount - 1
            If currentlocation = lstIE.List(k) Then GoTo KetThuc
        Next
        If ccTMP = currentlocation Then GoTo TheEnd
        ccTMP = currentlocation
        ValidData = InStr(1, buffer, iewindow.LocationName & "|" & iewindow.LocationURL & "|")
        If ValidData = 0 Then
        
            If Mid$(currentlocation, 1, 7) = "file://" Then
                 currentlocation = Replace(currentlocation, "file:///", "")
                 currentlocation = Replace(currentlocation, "%20", " ")
                 currentlocation = Replace(currentlocation, "/", "\")
                   'FullPathSearch currentlocation, c
                   Dim chErr As Integer
                   chErr = InStr(1, currentlocation, "\", vbBinaryCompare)
                   If (chErr > 3) Or (chErr = 0) Then GoTo KetThuc
                   Dim strTMP As String
                   File1.Path = currentlocation
                   File1.Refresh

                   'If FileTMP.ListCount <> File1.ListCount Then
                        'Thao tac nay cung thoa tac o tren ngam do ton bo nho cua PC
                        'Tuy nhien day chua phai la bien phap toi uu
                        'FileTMP.Path = currentlocation
                        'FileTMP.Refresh
                        
                        'Debug.Print currentlocation
                        
                        Dim i As Integer
                        Dim j As Integer
                    For i = 0 To File1.ListCount - 1
                        strTMP = File1.List(i)
                        Dim tmp As String
                        tmp = IIf(Len(currentlocation) = 3, Left(currentlocation, 2), currentlocation)
                        If DungLuong(tmp & "\" & File1.List(i)) > 3000000 Then GoTo KetThuc
                        
                        'Kiem tra chuoi CRC
                        If ScanCRCMain(Hex$(m_CRC.CalculateFile(currentlocation & "\" & File1.List(i))), currentlocation & "\" & File1.List(i)) = True Then GoTo KetThuc

                        'Kiem tra voi Icon
                        If UCase(Right(strTMP, 3)) = "EXE" Then
                             strTMP = SoSanhImage(tmp & "\" & File1.List(i), Pic, pic1, Ima)
                             'Kiem tra file thong qua Icon (Chi kiem tra file exe)
                             If strTMP <> "0" Then Detect GetStr("DecVirus"), "Virus : " & strTMP, tmp & "\" & File1.List(i): GoTo KetThuc
                         End If
                         
                                 'Kiem tra file thong qua cac chuoi String
                                     Dim BoDem As String
                                         Open currentlocation & "\" & File1.List(i) For Binary As #1
                                             BoDem = Space(LOF(1))
                                             Get #1, , BoDem
                                         Close #1
                                     For j = 0 To lstStr.ListCount - 1
                                         If InStr(1, BoDem, lstStr.List(j), vbBinaryCompare) <> 0 Then
                                             Detect GetStr("DecFile"), lstSDec.List(j), currentlocation & "\" & File1.List(i)
                                             Exit For
                                         End If
                                     Next
                                     'Nhan dang virus thong qua chuoi string (Khong the xac dinh virus qua CRC)
                                    For j = 0 To lstSVir.ListCount - 1
                                         If InStr(1, BoDem, lstSVir.List(j), vbBinaryCompare) <> 0 Then
                                             Detect GetStr("DecVir"), lstVirNa.List(j), currentlocation & "\" & File1.List(i), lstVirDat.List(j)
                                             Exit For
                                         End If
                                     Next
                                     
                                     BoDem = ""
KetThuc:
                        Next
                        
                   'End If
            End If
        End If

busysignal:
        
    Next
        lstIE.Clear
        For k = 0 To lstTMP.ListCount - 1
            lstIE.AddItem lstTMP.List(k)
        Next
                    
    Timer.Enabled = True
    On Error GoTo 0
TheEnd:
    Timer.Enabled = True
'Hi hi, de cai ten hay hay chut cho "Nghe si" cai nao
'Debug.Print currentlocation
End Sub
Private Sub ThietLapForm()
'Khoi tao cac thong tin chuan bi cho viec nhan dang worm
    File.Path = PathApp & "\Dat\Icon"
    File.Refresh
    Ima.ListImages.Clear
    Dim i
    For i = 0 To File.ListCount - 1
        picLoad.Cls
        picLoad.Picture = LoadPicture(PathApp & "\Dat\Icon" & "\" & File.List(i))
        Ima.ListImages.Add , Left(File.List(i), Len(File.List(i)) - 4), picLoad.Image
    Next
    'GetData App.Path & "\Dat\Sign.vnd", lstDat, lstNa, True
    'GetData App.Path & "\Dat\SW\YourSign.vnd", lstDat, lstNa, False
    GetData App.Path & "\Dat\String.vnd", lstStr, lstSDec, True
If XoaSach = True Then
    lstData.Clear
    lstNameV.Clear
End If

'Them CSDL virus vao vnAV
    lstSVir.Clear
    lstVirNa.Clear
    lstVirDat.Clear
    Open App.Path & "\Dat\Virus.vnd" For Input As #1
        Do While Not EOF(1)
            Line Input #1, InputData
            lstSVir.AddItem Split(InputData, "|", , vbBinaryCompare)(0)
            lstVirNa.AddItem Split(InputData, "|", , vbBinaryCompare)(1)
            lstVirDat.AddItem Split(InputData, "||", , vbBinaryCompare)(1)
        Loop
    Close #1
End Sub
