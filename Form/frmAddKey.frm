VERSION 5.00
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add key startup"
   ClientHeight    =   1890
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
   Icon            =   "frmAddKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   2640
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   3495
   End
   Begin CtrUnicodeVN.LabelUni lblType 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   120
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
      Caption         =   "Kie63u :"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin VB.ComboBox cmbKey 
      Appearance      =   0  'Flat
      Height          =   330
      ItemData        =   "frmAddKey.frx":058A
      Left            =   1080
      List            =   "frmAddKey.frx":0594
      TabIndex        =   5
      Text            =   "HKEY_LOCAL_MACHINE"
      Top             =   120
      Width           =   3615
   End
   Begin CtrUnicodeVN.ButtonUni cmdCancel 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
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
      MICON           =   "frmAddKey.frx":05BF
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
   Begin CtrUnicodeVN.ButtonUni cmdOk 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "D9o62ng y1"
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
      MICON           =   "frmAddKey.frx":05DB
      PICN            =   "frmAddKey.frx":05F7
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
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin CtrUnicodeVN.LabelUni lblName 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   600
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
      Caption         =   "Te6n kho1a :"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblPath 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
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
      Caption         =   "D9u7o72ng da64n :"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.ButtonUni cmdBro 
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Top             =   960
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
      MICON           =   "frmAddKey.frx":0B91
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
Attribute VB_Name = "frmAddKey"
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
    cd.Filter = "Protable files (*.pif;*.exe)|*.exe;*.pif|All Files (*.*)|*.*"
    Dim Path As String
    cd.ShowOpen
    If cd.Filename <> "" Then txtPath.Text = cd.Filename
End Sub

Private Sub cmdCancel_Click()
    frmSta.Show
    Unload Me
End Sub

Private Sub cmdOk_Click()

Dim GiaTri As String
If (txtName.Text <> "") And (txtPath.Text <> "") Then
    If cmbKey.Text = "HKEY_CURRENT_USER" Then
    
        GiaTri = GetString(HKEY_CURRENT_USER, Pathkey, txtName.Text)
        If GiaTri = "" Then SaveString HKEY_CURRENT_USER, Pathkey, txtName.Text, txtPath.Text: ThongBao "vnAntiVirus", GetStr("MesComAdd")
    ElseIf cmbKey.Text = "HKEY_LOCAL_MACHINE" Then
    
        GiaTri = GetString(HKEY_LOCAL_MACHINE, Pathkey, txtName.Text)
        If GiaTri = "" Then SaveString HKEY_LOCAL_MACHINE, Pathkey, txtName.Text, txtPath.Text: ThongBao "vnAntiVirus", GetStr("MesComAdd")
    End If
Else
    ThongBao "vnAntiVirus", GetStr("MesNF")
End If
End Sub

Private Sub Form_Load()
        Language Me
End Sub
