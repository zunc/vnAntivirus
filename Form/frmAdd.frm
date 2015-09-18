VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add data worm"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File 
      Height          =   510
      Left            =   1440
      Pattern         =   "*.ico"
      TabIndex        =   14
      Top             =   3480
      Width           =   615
   End
   Begin VB.PictureBox picCom 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   300
   End
   Begin MSComctlLib.ImageList ima 
      Left            =   120
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   480
      Width           =   3255
   End
   Begin CtrUnicodeVN.ButtonUni cmdCancel 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmAdd.frx":058A
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
      Left            =   3000
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Ca65p nha65t"
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
      MICON           =   "frmAdd.frx":05A6
      PICN            =   "frmAdd.frx":05C2
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
   Begin CtrUnicodeVN.FrameUni frmSet 
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
      Caption         =   "Ca61u hi2nh :"
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
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   275
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   275
      End
      Begin CtrUnicodeVN.LabelUni lblIcon 
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "Nha65n da5ng worm ba82ng ca1ch que1t bie63u tu7o75ng"
         SetKieugoTV     =   1
         Appearance      =   0
      End
      Begin CtrUnicodeVN.OptionUni optFile 
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Que1t theo ma4 nha65n da5ng rie6ng"
         SetKieugoTV     =   1
      End
      Begin CtrUnicodeVN.OptionUni optIcon 
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
         Caption         =   "Que1t vo71i bie63u tu7o75ng"
         SetKieugoTV     =   1
      End
      Begin CtrUnicodeVN.LabelUni lblFile 
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "Nha65n da5ng worm ba82ng ma4 rie6ng cu3a tu72ng file"
         SetKieugoTV     =   1
         Appearance      =   0
      End
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin CtrUnicodeVN.LabelUni lblPath 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
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
      Caption         =   "D9u7o72ng da64n:"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.ButtonUni cmdBro 
      Height          =   285
      Left            =   3960
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
      MICON           =   "frmAdd.frx":0B5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CtrUnicodeVN.LabelUni lblName 
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   480
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
      Caption         =   "Te6n worm:"
      SetKieugoTV     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Private Sub cmdBro_Click()
    cd.DialogTitle = "Choose a file ..."
    cd.Filter = "Portable files (*.pif;*.exe)|*.exe;*.pif|All Files (*.*)|*.*"
    Dim Path As String
    cd.ShowOpen
    If cd.Filename <> "" Then txtPath.Text = cd.Filename: GetIcon txtPath.Text, Pic: Pic.Visible = True
End Sub
Private Sub cmdCancel_Click()
    frmDat.Show
    Unload Me
End Sub
Private Sub cmdUp_Click()
Dim kq As String
If txtName.Text <> "" Then
    If optIcon.Value = True Then
        If FileExists(PathApp & "\Dat\Icon\" & txtName.Text & ".ico") = True Then
            ThongBao "vnAntiVirus", GetStr("MesTDL")
        Else
            kq = KiemTraIcon
            If kq = "" Then
                SavePicture Pic.Image, PathApp & "\Dat\Icon\" & txtName.Text & ".ico"
                frmDat.GetInfo
                ThongBao "vnAntiVirus", GetStr("MesComUI") & txtName.Text
            Else
                ThongBao "vnAntiVirus", GetStr("MesIAs") & kq
            End If
        End If
    ElseIf optFile.Value = True Then

        Dim tt As Boolean
        Dim tt1 As Boolean
        
        tt = False
        tt1 = False
        
        Dim tn As String
        Dim CRC As String
        CRC = Hex$(m_CRC.CalculateFile(txtPath.Text))
            If CRC <> "0" Then
                Dim InputData As String
                Open PathApp & "\Dat\Sign\" & Left(CRC, 2) & ".vnd" For Input As #1
                    Do While Not EOF(1)
                        Line Input #1, InputData
                        If CRC = Split(InputData, "|", , vbBinaryCompare)(0) Then tt = True
                        If txtName.Text = Split(InputData, "|", , vbBinaryCompare)(1) Then tt1 = True
                    Loop
                Close #1
                If tt = True Then ThongBao "vnAntiVirus", "Virus na2y d9a4 d9u7o75c ca65p nha65t": GoTo KetThuc1
                If tt1 = True Then ThongBao "vnAntiVirus", "Te6n virus bi5 tru2ng": GoTo KetThuc1
                
                AddToFile CRC & "|" & txtName.Text, PathApp & "\Dat\Sign\" & Left(CRC, 2) & ".vnd"
                                ThongBao "vnAntiVirus", GetStr("MesComUD") & " : " & txtName.Text
            Else
                ThongBao "vnAntiVirus", "Du74 lie65u nha65p va2o co1 lo64i"
KetThuc1:
            End If

    End If
'End If
Else
    ThongBao "vnAntiVirus", GetStr("MesNN")
End If

End Sub
Private Sub Form_Load()
    Language Me
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
End Sub
Private Sub txtPath_Change()
If FileExists(txtPath.Text) = True Then GetIcon txtPath.Text, Pic: Pic.Visible = True
End Sub
Private Function KiemTraIcon() As String
    KiemTraIcon = ""
    'Xu ly thao tac kiem tra Icon xem co ton tai trong data chua
    Ima.ListImages.Clear
    File.Path = PathApp & "\Dat\Icon"
    File.Refresh
    Dim i As Integer
    Dim re As Byte
    Dim kq As String
    kq = ""
    For i = 0 To File.ListCount - 1
        picCom.Cls
        picCom.Picture = LoadPicture(PathApp & "\Dat\Icon\" & File.List(i))
        PaP picCom, Pic, Pic.Width, Pic.Height, 15, re
        If re = 100 Then KiemTraIcon = Left(File.List(i), Len(File.List(i)) - 4)
    Next
End Function
