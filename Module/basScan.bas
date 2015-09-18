Attribute VB_Name = "basScan"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public m_CRC As clsCRC
Public ph As Long
'start CRC engine
Public Sub GetData(FileDat As String, lstData As ListBox, lstNameV As ListBox, XoaSach As Boolean)
If XoaSach = True Then
    lstData.Clear
    lstNameV.Clear
End If
    Open FileDat For Input As #1
        Do While Not EOF(1)
            Line Input #1, InputData
            lstData.AddItem Split(InputData, "|", , vbBinaryCompare)(0)
            lstNameV.AddItem Split(InputData, "|", , vbBinaryCompare)(1)
        Loop
    Close #1
End Sub
Public Function ScanFile(FilePath As String, ScanIcon As Boolean, ScanString As Boolean, ScanVir As Boolean, lstIma As ImageList, Picture1 As PictureBox, Picture2 As PictureBox) As String
'Thu tuc nay chi duoc su dung de quet process va cac file startup
'neu su dung thu tuc nay de quet cac thanh phan khac se cham hon binh thuong
'tai phien ban sap toi, Dung Coi se su dung thu tuc trong gan nhu tat ca cac thu tuc quet virus (Luc do se tem mot so phan tuy chon)

    On Error Resume Next
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    
'If FileExists(FilePath) = True Then
'Neu su dung dong lenh tren, thi khi quet USB se xuat hien loi khong the thoat dia USB nay khoi PC

Dim i As Integer
'Tien hanh kiem tra file theo 3 thong so (Icon,CRC,String)
    If ScanCRCMain(Hex$(m_CRC.CalculateFile(FilePath)), FilePath) = True Then GoTo KetThuc

    If ScanIcon = True Then
        Dim strTMP As String
        If UCase(Right(FilePath, 3)) = "EXE" Then
            strTMP = SoSanhImage(FilePath, Picture1, Picture2, lstIma)
            'Kiem tra file thong qua Icon (Chi kiem tra file exe)
            If strTMP <> "0" Then Detect GetStr("DecVirus"), "Virus : " & strTMP, FilePath: GoTo KetThuc
        End If
    End If

    If (ScanString = True) Or (ScanVir = True) Then
        Dim BoDem As String
        Open FilePath For Binary As #1
            BoDem = Space(LOF(1))
            Get #1, , BoDem
            Close #1
    Else
        GoTo KetThuc
    End If
    If ScanString = True Then
        'Kiem tra file thong qua cac chuoi String
        For i = 0 To frmMnu.lstStr.ListCount - 1
            If InStr(1, BoDem, frmMnu.lstStr.List(i), vbBinaryCompare) <> 0 Then
                Detect GetStr("DecFile"), frmMnu.lstSDec.List(i), FilePath
                GoTo KetThuc
            End If
        Next
    End If

    If ScanVir = True Then
        'Kiem tra file co phai la virus hay khong
        'quy trinh kiem tra giong kiem tra worm qua chuoi String, tuy nhien cong viec lai khac nhau
        'nen chung ta ne tach ra thanh 2 phan rieng biet
        For i = 0 To frmMnu.lstSVir.ListCount - 1
            If InStr(1, BoDem, frmMnu.lstSVir.List(i), vbBinaryCompare) <> 0 Then
                Detect GetStr("DecVir"), frmMnu.lstVirNa.List(i), FilePath, frmMnu.lstVirDat.List(i)
                GoTo KetThuc
            End If
        Next
    End If
        BoDem = ""
'Else

'End If
KetThuc:
End Function
Public Function ScanCRCMain(strCode As String, FilePath As String) As Boolean
'On Error Resume Next
    ScanCRCMain = False
If strCode = "0" Then GoTo KetThuc
    Dim InputData As String
    Open PathApp & "\Dat\Sign\" & Left(strCode, 2) & ".vnd" For Input As #1
        Do While Not EOF(1)
            Line Input #1, InputData
            If strCode = Split(InputData, "|", , vbBinaryCompare)(0) Then
                Detect GetStr("DecVirus"), "Virus: " & Split(InputData, "|", , vbBinaryCompare)(1), FilePath
                ph = ph + 1
                frmScan.lblPH.Caption = ph & " file"
                ScanCRCMain = True
                GoTo KetThuc
            End If
        Loop
            Close #1
KetThuc:
    Close #1

End Function
