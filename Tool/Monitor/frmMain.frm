VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4290
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEND 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   120
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code theo doi su thay doi cua thu muc tu PSC
Option Explicit

'Get the directory chages using ReadDirectoryChangesW
'Author: Pierre AOUN
'email: pierre_aoun@hotmail.com
Dim ThreadHandle   As Long
Dim Fin As Boolean
Private Sub cmdStart_Click()
Dim Dummy As Long
Dim Changes As String
Dim WaitNum As Long
  WSubFolder = True
  WatchStart = True
'Get Folder Handle
  If Right(PathDec, 1) <> "\" Then PathDec = PathDec + "\"
  DirHndl = GetDirHndl(PathDec)
  If (DirHndl = 0) Or (DirHndl = -1) Then MsgBox "Cannot create handle": Exit Sub
  'cmdStart.Enabled = False
  'cmdStop.Enabled = True
  'Create thread to Watch changes
Do
    ThreadHandle = CreateThread(ByVal 0&, ByVal 0&, AddressOf StartWatch, DirHndl, 0, Dummy)
    Do
    WaitNum = WaitForSingleObject(ThreadHandle, 50)
    DoEvents
    Loop Until (WaitNum = 0) Or (WatchStart = False)
    Changes = ""
    If WaitNum = 0 Then Changes = GetChanges
    
    If Changes <> "" Then
        Dim strTMP As String
        Dim strTMP1 As String
        lbl.Caption = Changes
        strTMP = Right(lbl.Caption, 3)
        If (strTMP = "exe") Or (strTMP = "pif") Then
            strTMP = CheckProcess(Changes)
                
            If strTMP <> "" Then
                strTMP1 = Split(strTMP, "|")(1)
                strTMP = Split(strTMP, "|")(0)
            End If
            If strTMP = "" Then
                ThongBao "Add file", Changes, Changes, 0
            ElseIf Drive_Type(Left(strTMP, 3)) = 2 Then
                SuspendResumeProcess Val(strTMP1), True
                ThongBao "Add file", Changes & "|" & strTMP & "|" & "Process run on USB", Changes, 2, Val(strTMP1)
            Else
                'SuspendResumeProcess Val(Split(CheckProcess(Changes), "|", , vbBinaryCompare)(1)), True
                ThongBao "Add file", Changes & "|" & strTMP, Changes, 1
            End If
        End If
        
    End If
Loop Until Not WatchStart
 'Terminate the Thread & Clear Handle
If DirHndl <> 0 Then ClearHndl DirHndl
If ThreadHandle <> 0 Then Call TerminateThread(ThreadHandle, ByVal 0&): ThreadHandle = 0
End Sub
'Private Sub cmdStop_Click()
'WatchStart = False
'cmdStop.Enabled = False
'cmdStart.Enabled = True
'End Sub
Private Sub Form_Load()
Me.Hide
App.TaskVisible = False
PathDec = "C:\" ' Command
cmdStart_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = -1
Fin = True
'cmdStop_Click
tmrEND.Enabled = True
End Sub
Private Sub tmrEND_Timer()
    End
End Sub
