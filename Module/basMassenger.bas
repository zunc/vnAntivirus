Attribute VB_Name = "basMessage"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Public dlv As Integer
Public Sub ThongBao(strCaption As String, NoiDung As String, Optional strPath As String)
    Dim frmTmp As New frmMas
    With frmTmp
        .lblCaption.Caption = strCaption
        .lblMes.Caption = NoiDung
        .lblPath.Caption = strPath
        .Show
    End With
End Sub
Public Sub Detect(Object As String, strDetect As String, Path As String, Optional strVirDat As String)
        Dim lsv As ListItem
        Dim i As Integer
        Dim tt As Boolean
        Dim tmp As Long
        tmp = 0
        'tmp = CheckProcess(Path)
        tt = False
    If tb = False Then
        With frmDetect
                .Show
                ThietLap .LV, .Ima, .Pic
                dlv = 1
                Set lsv = .LV.ListItems.Add()
                lsv.Text = Object
                lsv.SubItems(1) = strDetect
                lsv.SubItems(2) = Path
                If tmp = 0 Then
                    lsv.SubItems(3) = GetStr("No")
                Else
                    lsv.SubItems(3) = tmp
                End If
                lsv.Checked = True
                .lstDat.AddItem strVirDat
        End With
    ElseIf tb = True Then
            With frmDetect
            'For i = 1 To .LV.ListItems.Count
            '    If .LV.ListItems(i).SubItems(2) = Path Then tt = True
            'Next
                If tt = False Then
                '.Show
                Set lsv = .LV.ListItems.Add()
                    lsv.Text = Object
                    lsv.SubItems(1) = strDetect
                    lsv.SubItems(2) = Path
                If tmp = 0 Then
                    lsv.SubItems(3) = GetStr("No")
                Else
                    lsv.SubItems(3) = tmp
                End If
                    lsv.Checked = True
                    .lstDat.AddItem strVirDat
                    dlv = dlv + 1
                End If
        End With
    End If
    'GetIconsDe frmDetect.LV, frmDetect.Ima, frmDetect.Pic
End Sub
Public Sub GetIconsDe(lstView As ListView, imaList As ImageList, picTmp As PictureBox)
    On Error Resume Next
        Dim lsv As ListItem
    For Each lsv In lstView.ListItems
            picTmp.Cls
            GetIcon lsv.SubItems(2), picTmp
            If lsv.Index = dlv Then imaList.ListImages.Add lsv.Index, , picTmp.Image
    Next
        
    With lstView
      .SmallIcons = imaList
      For Each lsv In .ListItems
        lsv.SmallIcon = lsv.Index
      Next
    End With
End Sub
