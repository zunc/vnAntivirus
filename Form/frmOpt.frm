VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmOpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin CtrUnicodeVN.FrameUni frmSta 
      Height          =   3135
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5530
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
      Begin CtrUnicodeVN.CheckBoxUni chkStartup 
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   423
         Value           =   1
         Caption         =   "La2m vie65c khi kho73i d9o65ng"
         Pic_UncheckedNormal=   "frmOpt.frx":058A
         Pic_CheckedNormal=   "frmOpt.frx":07E4
         Pic_MixedNormal =   "frmOpt.frx":0A3E
         Pic_UncheckedDisabled=   "frmOpt.frx":0C98
         Pic_CheckedDisabled=   "frmOpt.frx":0EF2
         Pic_MixedDisabled=   "frmOpt.frx":114C
         Pic_UncheckedOver=   "frmOpt.frx":13A6
         Pic_CheckedOver =   "frmOpt.frx":1600
         Pic_MixedOver   =   "frmOpt.frx":185A
         Pic_UncheckedDown=   "frmOpt.frx":1AB4
         Pic_CheckedDown =   "frmOpt.frx":1D0E
         Pic_MixedDown   =   "frmOpt.frx":1F68
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
      Begin CtrUnicodeVN.CheckBoxUni chkShow 
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   423
         Value           =   1
         Caption         =   "Hie65n chu7o7ng tri2nh"
         Pic_UncheckedNormal=   "frmOpt.frx":21C2
         Pic_CheckedNormal=   "frmOpt.frx":241C
         Pic_MixedNormal =   "frmOpt.frx":2676
         Pic_UncheckedDisabled=   "frmOpt.frx":28D0
         Pic_CheckedDisabled=   "frmOpt.frx":2B2A
         Pic_MixedDisabled=   "frmOpt.frx":2D84
         Pic_UncheckedOver=   "frmOpt.frx":2FDE
         Pic_CheckedOver =   "frmOpt.frx":3238
         Pic_MixedOver   =   "frmOpt.frx":3492
         Pic_UncheckedDown=   "frmOpt.frx":36EC
         Pic_CheckedDown =   "frmOpt.frx":3946
         Pic_MixedDown   =   "frmOpt.frx":3BA0
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
      Begin CtrUnicodeVN.CheckBoxUni chkSystemTray 
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   423
         Value           =   1
         Caption         =   "Hie65n tre6n khay he65 tho61ng"
         Pic_UncheckedNormal=   "frmOpt.frx":3DFA
         Pic_CheckedNormal=   "frmOpt.frx":4054
         Pic_MixedNormal =   "frmOpt.frx":42AE
         Pic_UncheckedDisabled=   "frmOpt.frx":4508
         Pic_CheckedDisabled=   "frmOpt.frx":4762
         Pic_MixedDisabled=   "frmOpt.frx":49BC
         Pic_UncheckedOver=   "frmOpt.frx":4C16
         Pic_CheckedOver =   "frmOpt.frx":4E70
         Pic_MixedOver   =   "frmOpt.frx":50CA
         Pic_UncheckedDown=   "frmOpt.frx":5324
         Pic_CheckedDown =   "frmOpt.frx":557E
         Pic_MixedDown   =   "frmOpt.frx":57D8
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
   End
   Begin MSComctlLib.ImageList ima 
      Left            =   1440
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpt.frx":5A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpt.frx":6684
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpt.frx":72D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CtrUnicodeVN.ButtonUni cmdOk 
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "D9o62ng y1"
      ENAB            =   0   'False
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
      MICON           =   "frmOpt.frx":7F28
      PICN            =   "frmOpt.frx":7F44
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
   Begin CtrUnicodeVN.FrameUni frmAutoScan 
      Height          =   3135
      Left            =   1440
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5530
      Caption         =   "Tu75 d9o65ng que1t"
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
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   16
         Text            =   "C:\Windows"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin CtrUnicodeVN.CheckBoxUni chkUSB 
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   423
         Value           =   1
         Caption         =   "Que1t tre6n USB"
         Pic_UncheckedNormal=   "frmOpt.frx":84DE
         Pic_CheckedNormal=   "frmOpt.frx":8738
         Pic_MixedNormal =   "frmOpt.frx":8992
         Pic_UncheckedDisabled=   "frmOpt.frx":8BEC
         Pic_CheckedDisabled=   "frmOpt.frx":8E46
         Pic_MixedDisabled=   "frmOpt.frx":90A0
         Pic_UncheckedOver=   "frmOpt.frx":92FA
         Pic_CheckedOver =   "frmOpt.frx":9554
         Pic_MixedOver   =   "frmOpt.frx":97AE
         Pic_UncheckedDown=   "frmOpt.frx":9A08
         Pic_CheckedDown =   "frmOpt.frx":9C62
         Pic_MixedDown   =   "frmOpt.frx":9EBC
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
      Begin CtrUnicodeVN.CheckBoxUni chkScanI 
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Value           =   1
         Caption         =   "Que1t bie63u tu7o75ng"
         Pic_UncheckedNormal=   "frmOpt.frx":A116
         Pic_CheckedNormal=   "frmOpt.frx":A370
         Pic_MixedNormal =   "frmOpt.frx":A5CA
         Pic_UncheckedDisabled=   "frmOpt.frx":A824
         Pic_CheckedDisabled=   "frmOpt.frx":AA7E
         Pic_MixedDisabled=   "frmOpt.frx":ACD8
         Pic_UncheckedOver=   "frmOpt.frx":AF32
         Pic_CheckedOver =   "frmOpt.frx":B18C
         Pic_MixedOver   =   "frmOpt.frx":B3E6
         Pic_UncheckedDown=   "frmOpt.frx":B640
         Pic_CheckedDown =   "frmOpt.frx":B89A
         Pic_MixedDown   =   "frmOpt.frx":BAF4
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
      Begin CtrUnicodeVN.CheckBoxUni chkAutoIT 
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   423
         Value           =   1
         Caption         =   "Que1t file vie61t tre6n AutoIT"
         Pic_UncheckedNormal=   "frmOpt.frx":BD4E
         Pic_CheckedNormal=   "frmOpt.frx":BFA8
         Pic_MixedNormal =   "frmOpt.frx":C202
         Pic_UncheckedDisabled=   "frmOpt.frx":C45C
         Pic_CheckedDisabled=   "frmOpt.frx":C6B6
         Pic_MixedDisabled=   "frmOpt.frx":C910
         Pic_UncheckedOver=   "frmOpt.frx":CB6A
         Pic_CheckedOver =   "frmOpt.frx":CDC4
         Pic_MixedOver   =   "frmOpt.frx":D01E
         Pic_UncheckedDown=   "frmOpt.frx":D278
         Pic_CheckedDown =   "frmOpt.frx":D4D2
         Pic_MixedDown   =   "frmOpt.frx":D72C
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
      Begin CtrUnicodeVN.CheckBoxUni chkSam 
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   423
         Value           =   1
         Caption         =   "Que1t theo ma64u"
         Pic_UncheckedNormal=   "frmOpt.frx":D986
         Pic_CheckedNormal=   "frmOpt.frx":DBE0
         Pic_MixedNormal =   "frmOpt.frx":DE3A
         Pic_UncheckedDisabled=   "frmOpt.frx":E094
         Pic_CheckedDisabled=   "frmOpt.frx":E2EE
         Pic_MixedDisabled=   "frmOpt.frx":E548
         Pic_UncheckedOver=   "frmOpt.frx":E7A2
         Pic_CheckedOver =   "frmOpt.frx":E9FC
         Pic_MixedOver   =   "frmOpt.frx":EC56
         Pic_UncheckedDown=   "frmOpt.frx":EEB0
         Pic_CheckedDown =   "frmOpt.frx":F10A
         Pic_MixedDown   =   "frmOpt.frx":F364
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
      Begin CtrUnicodeVN.CheckBoxUni chkDec 
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   423
         Value           =   1
         Caption         =   "Nha65n da5ng the6m ta65p tin va2o he65 tho61ng"
         Pic_UncheckedNormal=   "frmOpt.frx":F5BE
         Pic_CheckedNormal=   "frmOpt.frx":F818
         Pic_MixedNormal =   "frmOpt.frx":FA72
         Pic_UncheckedDisabled=   "frmOpt.frx":FCCC
         Pic_CheckedDisabled=   "frmOpt.frx":FF26
         Pic_MixedDisabled=   "frmOpt.frx":10180
         Pic_UncheckedOver=   "frmOpt.frx":103DA
         Pic_CheckedOver =   "frmOpt.frx":10634
         Pic_MixedOver   =   "frmOpt.frx":1088E
         Pic_UncheckedDown=   "frmOpt.frx":10AE8
         Pic_CheckedDown =   "frmOpt.frx":10D42
         Pic_MixedDown   =   "frmOpt.frx":10F9C
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
      Begin CtrUnicodeVN.ButtonUni cmdBrowFolder 
         Height          =   285
         Left            =   2760
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
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
         MICON           =   "frmOpt.frx":111F6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin CtrUnicodeVN.FrameUni frmLang 
      Height          =   3135
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5530
      Caption         =   "Ngo6n ngu74"
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
      Begin CtrUnicodeVN.OptionUni optEng 
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "English"
         SetKieugoTV     =   1
      End
      Begin CtrUnicodeVN.OptionUni optVie 
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "Tie61ng Vie65t"
         SetKieugoTV     =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   5741
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ima"
      ForeColor       =   12582912
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnAvant"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin CtrUnicodeVN.ButtonUni cmdCancel 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmOpt.frx":11212
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
Attribute VB_Name = "frmOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source
Private Sub chkAutoIT_Click()
    cmdOk.Enabled = True
    chkUSB.Value = Checked
    If (chkAutoIT.Value = Unchecked) And (chkScanI.Value = Unchecked) And (chkSam.Value = Unchecked) Then chkUSB.Value = Unchecked
End Sub

Private Sub chkDec_Click()
    cmdOk.Enabled = True
    If chkDec.Value = Checked Then
        cmdBrowFolder.Enabled = True
        txtPath.Enabled = True
    Else
        cmdBrowFolder.Enabled = False
        txtPath.Enabled = False
    End If
End Sub
Private Sub chkSam_Click()
    cmdOk.Enabled = True
    chkUSB.Value = Checked
    If (chkAutoIT.Value = Unchecked) And (chkScanI.Value = Unchecked) And (chkSam.Value = Unchecked) Then chkUSB.Value = Unchecked
End Sub
Private Sub chkScanI_Click()
    cmdOk.Enabled = True
    chkUSB.Value = Checked
    If (chkAutoIT.Value = Unchecked) And (chkScanI.Value = Unchecked) And (chkSam.Value = Unchecked) Then chkUSB.Value = Unchecked
End Sub
Private Sub chkShow_Click()
    cmdOk.Enabled = True
    If (chkSystemTray.Value = Unchecked) And (chkShow.Value = Unchecked) Then chkSystemTray.Value = Checked
End Sub
Private Sub chkStartup_Click()
    cmdOk.Enabled = True
    If chkStartup.Value = Checked Then
        SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "vnAntivirus", PathApp & "\" & App.exename
    Else
        DelSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "vnAntivirus"
    End If
End Sub

Private Sub chkSystemTray_Click()
    cmdOk.Enabled = True
    If (chkSystemTray.Value = Unchecked) And (chkShow.Value = Unchecked) Then chkShow.Value = Checked
End Sub

Private Sub chkUSB_Click()
    cmdOk.Enabled = True
    If chkUSB.Value = Unchecked Then
        chkScanI.Enabled = False
        chkAutoIT.Enabled = False
        chkSam.Enabled = False
        chkScanI.Value = Unchecked
        chkAutoIT.Value = Unchecked
        chkSam.Value = Unchecked
    Else
        chkScanI.Enabled = True
        chkAutoIT.Enabled = True
        chkSam.Enabled = True
        chkScanI.Value = Checked
        chkAutoIT.Value = Checked
        chkSam.Value = Checked
    End If
End Sub

Private Sub cmdBrowFolder_Click()
Dim sOutPut
    sOutPut = ""
    sOutPut = GetFolder(Me.hwnd, "Scan Path : ", txtPath.Text)
    If sOutPut <> "" Then
        txtPath.Text = sOutPut
    Else
        If Len(txtPath.Text) <> 0 Then
            Else
            ThongBao "vnAntiVirus", GetStr("MesSe")
    End If
    End If
End Sub

Private Sub cmdCancel_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdOk_Click()

        cmdOk.Enabled = False
        ControlSet chkUSB, ichkUSB
        ControlSet chkScanI, ichkScanI
        ControlSet chkAutoIT, ichkAutoIT
        ControlSet chkSam, ichkSam
        ControlSet chkDec, ichkDec
        
        ControlSet chkShow, ichkShow
        ControlSet chkSystemTray, ichkSystemTray
        
        ioptVie = optVie.Value
        
    WriteINI PathApp & "\Data.ini", "Option", "chkUSB", QuyDoi1(ichkUSB)
    WriteINI PathApp & "\Data.ini", "Option", "chkScanI", QuyDoi1(ichkScanI)
    WriteINI PathApp & "\Data.ini", "Option", "chkAutoIT", QuyDoi1(ichkAutoIT)
    WriteINI PathApp & "\Data.ini", "Option", "chkSam", QuyDoi1(ichkSam)
    WriteINI PathApp & "\Data.ini", "Option", "chkDec", QuyDoi1(ichkDec)
    WriteINI PathApp & "\Data.ini", "Option", "PathDec", txtPath.Text

    WriteINI PathApp & "\Data.ini", "Option", "chkShow", QuyDoi1(ichkShow)
    WriteINI PathApp & "\Data.ini", "Option", "chkSystemTray", QuyDoi1(ichkSystemTray)
    WriteINI PathApp & "\Data.ini", "Option", "optVie", QuyDoi1(ioptVie)
        ResetLV
        Language Me
        Language frmMnu
    'frmMain.Show
    'frmPro.Show
End Sub

Private Sub Form_Load()

    Language Me
    'MsgBox ReadINI("D:\MySoft\LovePN\Language\EngLish.lng", "frmOpt", "chkUSB")
        SetControl chkUSB, ichkUSB
        SetControl chkScanI, ichkScanI
        SetControl chkAutoIT, ichkAutoIT
        SetControl chkSam, ichkSam
        SetControl chkDec, ichkDec
        ResetLV
        If ichkDec = True Then
            txtPath.Enabled = True
            cmdBrowFolder.Enabled = True
        Else
            txtPath.Enabled = False
            cmdBrowFolder.Enabled = False
        End If
        txtPath.Text = PathDec
        
        SetControl chkShow, ichkShow
        SetControl chkSystemTray, ichkSystemTray
        
        If GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "vnAntivirus") = PathApp & "\" & App.exename Then
            chkStartup.Value = Checked
        Else
            chkStartup.Value = Unchecked
        End If
        
        optVie.Value = ioptVie
        optEng.Value = Not (ioptVie)
        cmdOk.Enabled = False
End Sub

'Hi hi, doan code nay hoi "Ngo"
Private Sub SetControl(chkOK As CheckBoxUni, iValue As Boolean)
If iValue = True Then
    chkOK.Value = Checked
Else
    chkOK.Value = Unchecked
End If
End Sub
Private Sub ControlSet(chkOK As CheckBoxUni, iValue As Boolean)
If chkOK.Value = Checked Then
    iValue = True
Else
    iValue = False
End If
End Sub
Private Sub LamMoi(frmOk As FrameUni)

Dim frm As Control
Dim tmp As String
Dim tmp1 As String
tmp1 = frmOk.Name
For Each i In Me.Controls
    tmp = i.Name
    If Left(tmp, 3) = "frm" Then
        i.Visible = False
        If tmp1 = tmp Then i.Visible = True
    End If
    
Next
End Sub

Private Sub LV_Click()
If LV.SelectedItem.Index = 1 Then
    LamMoi frmSta
ElseIf LV.SelectedItem.Index = 2 Then
    LamMoi frmLang
ElseIf LV.SelectedItem.Index = 3 Then
    LamMoi frmAutoScan
End If
End Sub
Private Sub LV_KeyUp(KeyCode As Integer, Shift As Integer)
If LV.SelectedItem.Index = 1 Then
    LamMoi frmSta
ElseIf LV.SelectedItem.Index = 2 Then
    LamMoi frmLang
ElseIf LV.SelectedItem.Index = 3 Then
    LamMoi frmAutoScan
End If
End Sub
Private Sub optEng_Click()
    cmdOk.Enabled = True
End Sub
Private Sub optVie_Click()
    cmdOk.Enabled = True
End Sub
Private Sub txtPath_Change()
    cmdOk.Enabled = True
End Sub
Private Sub ResetLV()
    LV.ListItems.Clear
    LV.ListItems.Add , , GetStrOther("Sta"), 1
    LV.ListItems.Add , , GetStrOther("Lan"), 2
    LV.ListItems.Add , , GetStrOther("Aut"), 3
    LV.Refresh
End Sub
