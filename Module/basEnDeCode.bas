Attribute VB_Name = "basEnDeCode"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Public Function Encode(Data As String, Optional Depth As Integer) As String
Dim TempChar As String
Dim TempAsc As Long
Dim NewData As String
Dim vChar As Long
For vChar = 1 To Len(Data)
    TempChar = Mid$(Data, vChar, 1)
        TempAsc = Asc(TempChar)
        If Depth = 0 Then Depth = 40
        If Depth > 254 Then Depth = 254

        TempAsc = TempAsc + Depth
        If TempAsc > 255 Then TempAsc = TempAsc - 255
        TempChar = Chr(TempAsc)
        NewData = NewData & TempChar
Next vChar
Encode = NewData

End Function
Public Function Decode(Data As String, Optional Depth As Integer) As String
Dim TempChar As String
Dim TempAsc As Long
Dim NewData As String
Dim vChar As Long

For vChar = 1 To Len(Data)
    TempChar = Mid$(Data, vChar, 1)
        TempAsc = Asc(TempChar)
        If Depth = 0 Then Depth = 40
        If Depth > 254 Then Depth = 254
    TempAsc = TempAsc - Depth
        If TempAsc < 0 Then TempAsc = TempAsc + 255
        TempChar = Chr(TempAsc)
        NewData = NewData & TempChar
Next vChar
Decode = NewData

End Function
