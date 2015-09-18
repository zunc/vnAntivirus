Attribute VB_Name = "basImage"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

'Code this module  from PSC
Public Sub PaP(p1 As PictureBox, p2 As PictureBox, W As Long, H As Long, C As Long, DDD As Byte)
    Dim X, Y, a, AP, DD1, DD2
    D = 10  ' for decimal digit in percentile part. d= 1 or 10 or 100
    For X = 0 To W - 15 Step C  ' StepbyStep for PixelComparable of both picture
        For Y = 0 To H - 15 Step C
            If p2.Point(X, Y) = p1.Point(X, Y) Then AP = AP + 1 ' if picture1 pointcolor= picture2 pointcolor |>counter=++1
            a = a + 1  ' a=programme counter
        Next
    Next
    DDD = (AP * 100) \ a 'percentile
    'DD2 = Right$((AP * (100 * D)) \ a, Len(D) - 1) 'decimal part
    'DDD = DD1 & "." & DD2 & "%" ' wrought percent
End Sub
Public Function SoSanhImage(PathFile As String, Picture1 As PictureBox, Picture2 As PictureBox, imaList As ImageList) As String
    SoSanhImage = "0"
    
    Picture1.Cls
    GetIcon PathFile, Picture1
    
    Dim i As Byte
    Dim re As Byte
    For i = 1 To imaList.ListImages.Count
        Picture2.Cls
        Picture2.Picture = imaList.ListImages(i).Picture
        PaP Picture1, Picture2, Picture1.Width, Picture1.Height, 15, re
        If re = 100 Then SoSanhImage = imaList.ListImages(i).Key: Exit For
    Next

End Function
