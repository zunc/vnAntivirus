Attribute VB_Name = "basCleanVirus"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

'Module nay phu trach viec lam sach file nhiem virus

Public Sub CleanVirus(Filename As String, strData As String)
Dim ht As String

ht = Split(strData, "|", , vbBinaryCompare)(0)
Select Case ht
    Case "A"
    'Virus chen file du lieu vao cuoi file
        Dim StrFileGoc As String
        Dim StrFileStr As String
        Dim strKieu As String
        
        StrFileGoc = Split(strData, "|", , vbBinaryCompare)(1)
        StrFileStr = Split(strData, "|", , vbBinaryCompare)(2)
        strKieu = Split(strData, "|", , vbBinaryCompare)(3)
        
        Dim BoDem As String
        Open Filename For Binary As #1
            BoDem = Space(LOF(1))
            Get #1, , BoDem
        Close #1
            Dim vt As Double
            vt = InStr(1, BoDem, StrFileGoc, vbBinaryCompare)
            If StrFileStr = "*" Then
            'Neu loai virus nay chi pha huy 1 loai file duy nhat
                XoaFile Filename
                BoDem = Right(BoDem, Len(BoDem) - vt + 1)
                Open Left(Filename, Len(Filename) - 4) & "." & strKieu For Binary Access Write As #1
                    Put #1, , BoDem
                Close #1
                ThongBao "Clean file", "D9a die65t virus"
            Else
            'Neu loai virus nay pha huy nhieu loai file khac nhau
            'Luc nay phai dua vao cau truc tung loai file de xu ly
            'cach thuc nay thuc ra ko phai la toi uu, cach xu ly thang trong chuoi Byte cua virus de truy tim ra dinh dang file
            'tuy nhien do trinh do "Ngu si" nen Dung Coi se ap dung cach nay
                BoDem = Right(BoDem, Len(BoDem) - vt + 1)
                Dim i As Byte
                i = 2
                Dim XuLy As Boolean
                XuLy = False
                Do While Split(strData, "|", , vbBinaryCompare)(i) <> ""
                'Chu y : Doan code nay co van de
                    If InStr(1, BoDem, Split(strData, "|", , vbTextCompare)(i)) <> 0 Then
                    'Tien hanh "Boc tach" file goc ra khoi virus
                        XuLy = True
                        XoaFile Filename
                        Open Left(Filename, Len(Filename) - 4) & "." & Split(strData, "|", , vbBinaryCompare)(i + 1) For Binary Access Write As #1
                            Put #1, , BoDem
                        Close #1
                        ThongBao "Clean file", "D9a die65t virus"
                        BoDem = ""
                        Exit Sub
                    End If
                    i = i + 2
                Loop
                If XuLy = False Then ThongBao "vnAntiVirus", "Chu7a the63 la2m sa5ch file"
            End If
    Case "G"
    'Virus de du lieu nam giua file
    'Do hien nay Dung Coi khong co mau cua loai virus nay ne tam thoi chua nghien cuu
    
End Select

End Sub
Public Function GetFileName(FilePath As String) As String
Dim vt As Integer
    vt = InStr(1, StrReverse(FilePath), "\", vbBinaryCompare)
    GetFileName = Right(FilePath, vt - 1)
End Function

