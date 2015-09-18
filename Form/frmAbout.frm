VERSION 5.00
Object = "{2DF2546F-C700-48AD-82B8-6C31E95FB639}#1.0#0"; "viettype.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tac gia"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin CtrUnicodeVN.LabelUni lblSend 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   "Pha62n me62m na2y la2 mo1n qua2 va2o nga2y sinh nha65t My Love (PN)."
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblEmail 
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "Email : dungcoivb@gmail.com"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblAuthor 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ta1c gia3 : Le6 Nguye6n Du4ng"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblName 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "vnAntivirus v0.5 (beta)"
      SetKieugoTV     =   1
      Appearance      =   0
      ForeColor       =   -2147483635
   End
   Begin CtrUnicodeVN.LabelUni lblOS 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
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
      Caption         =   "vnAntivirus la2 mo65t pha62n me62m ma4 nguo62n mo73"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.ButtonUni cmdBack 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
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
      MICON           =   "frmAbout.frx":058A
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
   Begin CtrUnicodeVN.LabelUni lblVer 
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2880
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
      Caption         =   "Ca65p nha65t 15/8/2007"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblSup 
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
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
      Caption         =   "Ho64 tro75 tru75c tuye61n : ....."
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin CtrUnicodeVN.LabelUni lblSam 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "Ba5n co1 the63 gu73i ma64u virus ta5i : http://www.vietvirus.info"
      SetKieugoTV     =   1
      Appearance      =   0
   End
   Begin VB.Image ima 
      Height          =   720
      Left            =   120
      Picture         =   "frmAbout.frx":05A6
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmAbout"
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
Private Sub Form_Load()
    Language Me
End Sub
Private Sub lblSam_Click()
    Shell "EXPLORER.EXE " & "http://www.vietvirus.info"
End Sub
