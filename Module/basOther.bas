Attribute VB_Name = "basOther"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Public tb As Boolean

Public ichkUSB As Boolean
Public ichkScanI As Boolean
Public ichkAutoIT As Boolean
Public ichkSam As Boolean
Public ichkDec As Boolean
Public PathDec As String

Public ioptVie As Boolean
Public ichkShow As Boolean
Public ichkSystemTray As Boolean

Public LoadMon As Boolean

Public PathWScan As String
Public SeeSta As Boolean

Public Sub GetProcess(lstView As ListView, imaList As ImageList, picTmp As PictureBox)
'On Error Resume Next
    picTmp.Cls
    picTmp.BackColor = vbWhite
    lstView.BackColor = vbWhite
    lstView.ListItems.Clear
    lstView.SmallIcons = Nothing
    imaList.ListImages.Clear
    frmMnu.lstPro.Clear
'---------Liet ke process-------
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
  Dim exename As String
  Dim ID As Long
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
    Dim lsv As ListItem
   While theloop <> 0

      ID = proc.th32ProcessID
      theloop = ProcessNext(snap, proc)
      picTmp.Cls
      If ProcessPathByPID(proc.th32ProcessID) <> "SYSTEM" Then
      'MsgBox ProcessPathByPID(proc.th32ProcessID)

                  Set lsv = lstView.ListItems.Add()
                  lsv.Text = proc.szExeFile
                  lsv.SubItems(1) = ProcessPathByPID(proc.th32ProcessID)
                  lsv.SubItems(2) = proc.th32ProcessID
        End If
   Wend
   CloseHandle snap
       EnumWindows AddressOf EnumWindowsProc, ByVal 0&

    'Dim lsv As ListItem
    For Each lsv In lstView.ListItems
        picTmp.Cls
        GetIcon lsv.SubItems(1), picTmp
        imaList.ListImages.Add lsv.Index, , picTmp.Image
    Next
    
With lstView
  .SmallIcons = imaList
  For Each lsv In .ListItems
    lsv.SmallIcon = lsv.Index
  Next
End With

End Sub
Public Sub GetIcons(lstView As ListView, imaList As ImageList, picTmp As PictureBox)

        Dim lsv As ListItem
For Each lsv In lstView.ListItems
        picTmp.Cls
        GetIcon lsv.SubItems(1), picTmp
        imaList.ListImages.Add lsv.Index, , picTmp.Image
Next
    
With lstView
  .SmallIcons = imaList
  For Each lsv In .ListItems
    lsv.SmallIcon = lsv.Index
  Next
End With
End Sub
Public Sub ThietLap(lstView As ListView, imaList As ImageList, picTmp As PictureBox)
'On Local Error Resume Next
    picTmp.Cls
    picTmp.BackColor = vbWhite
    lstView.BackColor = vbWhite
    lstView.ListItems.Clear
    lstView.SmallIcons = Nothing
    imaList.ListImages.Clear
End Sub

