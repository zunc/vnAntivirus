Attribute VB_Name = "basFileINI"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function ReadINI(Filename As String, Section As String, KeyName As String) As String
    Dim Ret As String, NC As Long
    Ret = String(255, 0)
    NC = GetPrivateProfileString(Section, KeyName, "", Ret, 255, Filename)
    
    If NC <> 0 Then Ret = Left$(Ret, NC)
    'Dim chuoi As String
    'chuoi = Ret
    'chuoi = StrReverse(chuoi)
    'Dim i, j, t
    'i = InStr(1, chuoi, ".", vbTextCompare)
    'j = InStr(1, Left(chuoi, i), " ", vbTextCompare)
    'chuoi = StrReverse(Right(chuoi, Len(chuoi) - j))
    'If FileExists(chuoi) = False Then chuoi = Left(FileName, Len(FileName) - InStr(1, StrReverse(FileName), "\", vbBinaryCompare)) & "\" & chuoi
    ReadINI = Ret
End Function
Public Sub WriteINI(Filename As String, Section As String, Key As String, newValue As String)
    WritePrivateProfileString Section, Key, newValue, Filename
End Sub
Public Function GetOpt()

        ichkUSB = QuyDoi(ReadINI(PathApp & "\Data.ini", "Option", "chkUSB"))
        ichkScanI = QuyDoi(ReadINI(PathApp & "\Data.ini", "Option", "chkScanI"))
        ichkAutoIT = QuyDoi(ReadINI(PathApp & "\Data.ini", "Option", "chkAutoIT"))
        ichkSam = QuyDoi(ReadINI(PathApp & "\Data.ini", "Option", "chkSam"))
        ichkDec = QuyDoi(ReadINI(PathApp & "\Data.ini", "Option", "chkDec"))
        ioptVie = QuyDoi(ReadINI(PathApp & "\Data.ini", "Option", "optVie"))
        ichkShow = QuyDoi(ReadINI(PathApp & "\Data.ini", "Option", "chkShow"))
        ichkSystemTray = QuyDoi(ReadINI(PathApp & "\Data.ini", "Option", "chkSystemTray"))
        PathDec = ReadINI(PathApp & "\Data.ini", "Option", "PathDec")
End Function
Public Function QuyDoi(So As String) As Boolean
    If So = "0" Then
        QuyDoi = False
    Else
        QuyDoi = True
    End If
End Function
Public Function QuyDoi1(GiaTri As Boolean) As String
    If GiaTri = False Then
        QuyDoi1 = "0"
    Else
        QuyDoi1 = "1"
    End If
End Function
