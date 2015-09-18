Attribute VB_Name = "basFile"
Option Explicit

'##############################################################################################
'Purpose: Used for File System operations
'Author:  Richard Mewett ©2004

'Credits:
'The GetFolder() code was sourced from VB.NET (Brad Martinez & Randy Birch)
'##############################################################################################

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH = 260

Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Public Type SHQUERYRBINFO
    cbSize As Long
    i64Size As ULARGE_INTEGER
    i64NumItems As ULARGE_INTEGER
End Type

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Type SYSTEMTIME
    wYear             As Integer
    wMonth            As Integer
    wDayOfWeek        As Integer
    wDay              As Integer
    wHour             As Integer
    wMinute           As Integer
    wSecond           As Integer
    wMilliseconds     As Integer
End Type

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

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Public Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Get icon


Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hdcDest&, ByVal X&, ByVal Y&, ByVal flags&) As Long

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private SIconInfo As SHFILEINFO

'---Tim dung luong------
Const GENERIC_READ = &H80000000
Const FILE_SHARE_READ = &H1
Const OPEN_EXISTING = 3
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'-----------------------------------------

'Dimensionalize SIconInfo as SHFILEINFO type structure
Public Sub GetIcon(icPath$, pDisp As PictureBox)
pDisp.Cls
Dim hImgSmall&: hImgSmall = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'call SHGetFileInfo to return a handle to the icon associated with the specified file
 ImageList_Draw hImgSmall, SIconInfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
 'Draw the icon to the specified picturebox control
End Sub
Public Sub GetLargeIcon(icPath$, pDisp As PictureBox)
Dim hImgLrg&: hImgLrg = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
'call SHGetFileInfo to return a handle to the icon associated with the specified file
 ImageList_Draw hImgLrg, SIconInfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
 'Draw the icon to the specified picturebox control
End Sub
Public Sub ShowProperties(sFilename As String, hwndOwner As Long)
    '##############################################################################################
    'Displays the Properties of the specified file
    '##############################################################################################
    
    Dim SEI As SHELLEXECUTEINFO
    
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = hwndOwner
        .lpVerb = "properties"
        .lpFile = sFilename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    
    Call ShellExecuteEx(SEI)
End Sub
Public Function GetSpecialfolder(CSIDL As Long) As String
    '##############################################################################################
    'Returns the Path to a "Special" Folder (i.e. Internet History)
    '##############################################################################################
    
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function
Public Function FileExists(sFilename As String) As Boolean
    '##############################################################################################
    'Returns True if the specified file exists
    'Ham nay neu su dung de kiem tra file tren USB se lap tuc gay ra loi
    '##############################################################################################
    
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
Public Function FolderExists(sFolder As String) As Boolean
    '##############################################################################################
    'Returns True if the specified folder exists
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function
Public Function GetFolder(hWnd As Long, Optional sPrompt As String, Optional sStartFolder As String) As String
    '##############################################################################################
    'Displays a Folder Browser to select a Folder
    '##############################################################################################
    
    Dim BI As BROWSEINFO
    Dim pidl As Long
    Dim sFolder As String
    Dim pos As Integer
    
    'Fill the BROWSEINFO structure with the needed data
    With BI
        'hwnd of the window that receives messages from the call. Can be your application or the handle from GetDesktopWindow().
        .hOwner = hWnd
        
        'Pointer to the item identifier list specifying the location of the "root" folder to browse from.
        'If NULL, the desktop folder is used.
        .pidlRoot = 0&
    
        'message to be displayed in the Browse dialog
        If Len(sPrompt) = 0 Then
            .lpszTitle = "Select the folder:"
        Else
            .lpszTitle = sPrompt
        End If
    
        'the type of folder to return. - the constants perform differently for non networked pc's
        .ulFlags = BIF_RETURNONLYFSDIRS
        
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)
        .lParam = SHSimpleIDListFromPath(StrConv(sStartFolder, vbUnicode))
    End With
        
    'show the browse for folders dialog
     pidl = SHBrowseForFolder(BI)
    
    'the dialog has closed, so parse & display the user's returned folder selection contained in pidl
    sFolder = Space$(MAX_PATH)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal sFolder) Then
        pos = InStr(sFolder, Chr$(0))
        GetFolder = Left$(sFolder, pos - 1)
    End If
    
    Call CoTaskMemFree(pidl)
End Function
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    '############################################################################
    'Purpose: Required by GetGolder() Function
    '############################################################################
 
    Select Case uMsg
    Case BFFM_INITIALIZED
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, 0&, ByVal lpData)
                    
    Case Else
    
    End Select
End Function
Private Function FARPROC(pfn As Long) As Long
    '############################################################################
    'Purpose: Required by GetGolder() Function
    
    'A dummy procedure that receives and returns
    'the value of the AddressOf operator.
    
    'This workaround is needed as you can't
    'assign AddressOf directly to a member of a
    'user-defined type, but you can assign it
    'to another long and use that instead!
    '############################################################################
 
    FARPROC = pfn
End Function
Public Function CheckThuMuc(DuongDan As String)
If Right(DuongDan, 1) = "\" Then DuongDan = Left(DuongDan, Len(DuongDan) - 1)
Dim TenThuMuc As String
If Len(DuongDan) > 2 Then TenThuMuc = Right(DuongDan, InStr(1, StrReverse(DuongDan), "\", vbTextCompare) - 1)

If FileExists(DuongDan & "\autorun.inf") = True Then
    Detect "Autorun.inf", GetStr("MesDectectAutorun") & XuLyChuoi(ReadINI(DuongDan & "\autorun.inf", "autorun", "open"), DuongDan), DuongDan & "\autorun.inf"
    If FileExists(XuLyChuoi(ReadINI(DuongDan & "\autorun.inf", "autorun", "open"), DuongDan)) = True Then Detect "File autorun", GetStr("MesAutoBe"), XuLyChuoi(ReadINI(DuongDan & "\autorun.inf", "autorun", "open"), DuongDan)

ElseIf FileExists(DuongDan & "\" & TenThuMuc & ".exe") Then
    'Detect TenThuMuc & ".exe", "File exe like name folder", FileExists(DuongDan & "\" & TenThuMuc & ".exe")
End If

End Function
Public Function XuLyChuoi(Chuoi As String, Disk As String) As String
'Thao tac xu ly cac chuoi gia tri tu regedit de dua ra ket qua chuan
    If InStr(1, Chuoi, "%systemroot%", vbTextCompare) <> 0 Then
        Chuoi = Replace(Chuoi, "%systemroot%", WindowsDir, 1, , vbTextCompare)
    Else
        Chuoi = Chuoi
    End If
    
    Chuoi = StrReverse(Chuoi)
    Dim i, j, t
    i = InStr(1, Chuoi, ".", vbTextCompare)
    j = InStr(1, Left(Chuoi, i), " ", vbTextCompare)
    Chuoi = StrReverse(Right(Chuoi, Len(Chuoi) - j))
    
    'Chuoi = StrReverse(Chuoi)
    Dim strTMP As String
    strTMP = Right(Chuoi, InStr(1, StrReverse(Chuoi), "\", vbTextCompare))
    t = InStr(1, strTMP, " ", vbTextCompare)
    
    'MsgBox Len(strTmp) - t, , t
    If t <> 0 Then Chuoi = Left(Chuoi, Len(Chuoi) - (Len(strTMP) - t + 1))
    
    If InStr(1, Chuoi, Chr(34), vbTextCompare) <> 0 Then
        Dim Buf As String
        Dim Path As String
        Dim vt1 As Byte
        Dim vt2 As Byte
        
            vt1 = InStr(1, Chuoi, Chr(34), vbTextCompare)
            Buf = Right(Chuoi, Len(Chuoi) - vt1)
            vt2 = InStr(1, Buf, Chr(34), vbTextCompare)
            Chuoi = Left(Buf, vt2 - 1)
            
    End If
    
    If FileExists(Disk & "\" & Chuoi) = True Then Chuoi = Disk & "\" & Chuoi
    If FileExists(Disk & Chuoi) = True Then Chuoi = Disk & Chuoi
    
    XuLyChuoi = Chuoi
End Function
Public Function DungLuong(DuongDan As String) As Long
Dim hFile As Long, nSize As Currency
    hFile = CreateFile(DuongDan, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    GetFileSizeEx hFile, nSize
    CloseHandle hFile
DungLuong = nSize * 10000
End Function
Public Function SoSanh(File1 As String, File2 As String) As Boolean
    Open File1 For Binary As #1
            Dim BoDem As String
            BoDem = Space(LOF(1))
            Get #1, , BoDem
        Close #1
    Open File2 For Binary As #2
            Dim BoDem1 As String
            BoDem1 = Space(LOF(2))
            Get #2, , BoDem1
        Close #2
        If BoDem1 = BoDem Then
            SoSanh = True
        Else
            SoSanh = False
        End If
        BoDem = ""
        BoDem1 = ""
End Function
Public Sub XoaFile(Filename As String)
    DeleteFile Filename
End Sub
Public Sub AddToFile(strValue As String, FilePath As String)
    Open FilePath For Append As #1
        Print #1, strValue
    Close #1
    DoEvents
End Sub
