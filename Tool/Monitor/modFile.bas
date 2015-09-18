Attribute VB_Name = "modFile"
Public Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

'---Tim dung luong------
Const GENERIC_READ = &H80000000
Const FILE_SHARE_READ = &H1
Const OPEN_EXISTING = 3
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'-----------------------------------------
Public Function DungLuong(DuongDan As String) As Long
Dim hFile As Long, nSize As Currency
    hFile = CreateFile(DuongDan, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    GetFileSizeEx hFile, nSize
    CloseHandle hFile
DungLuong = nSize * 10000
End Function
Public Function SoSanh(File1 As String, File2 As String) As Boolean
'If TonTai(File1) = True Then
Do While TonTai(File1) = False
'Hi hi, cau thoi gian cho vui thoi nhe
Loop
    Open File1 For Binary As #1
            Dim BoDem As String
            BoDem = Space(LOF(1))
            Get #1, , BoDem
        Close #1
    If TonTai(File2) = True Then
        Open File2 For Binary As #2
                Dim BoDem1 As String
                BoDem1 = Space(LOF(2))
                Get #2, , BoDem1
            Close #2
    End If
        If BoDem1 = BoDem Then
            SoSanh = True
        Else
            SoSanh = False
        End If
        BoDem = ""
        BoDem1 = ""
'End If
End Function
Public Function Drive_Type(DriveLetter As Variant) As Long
    Dim strDL As String
    strDL = Left$(DriveLetter, 1) + ":\"
    Drive_Type = GetDriveType(strDL)
End Function
Public Function TonTai(sFilename As String) As Boolean

    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFilename, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FileExists = False
        Else
            FileExists = True
        End If
    End If
End Function
