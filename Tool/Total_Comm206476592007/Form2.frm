VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   525
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton dcButton2 
      Height          =   315
      Left            =   2760
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton dcButton1 
      Height          =   315
      Left            =   2040
      Picture         =   "Form2.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Text            =   "*.*"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filter:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   450
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dcButton1_Click()

If Text1.Text <> Form1.StatusBar1.Panels(2).Text Then
Form1.StatusBar1.Panels(2).Text = Text1.Text
Form1.StatusBar1.Tag = "Ok"
Unload Me
Else
Call dcButton2_Click
End If
End Sub

Private Sub dcButton2_Click()
Form1.StatusBar1.Tag = "Cancel"
Unload Me
End Sub

Private Sub Form_Load()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dcButton1_Click
End Sub
