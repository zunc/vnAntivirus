Attribute VB_Name = "Icons"
Option Explicit
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
'the first api is used to locate the icon for each file
'the second draws the image into a picturebox

Public Const LARGE_ICON As Integer = 32
Public Const SMALL_ICON As Integer = 16
Public Const MAX_PATH = 260
Public Const ILD_TRANSPARENT = &H1                      '  Display transparent
Public Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Public Const SHGFI_EXETYPE = &H2000                     '  return exe type
Public Const SHGFI_LARGEICON = &H0                      '  get large icon
Public Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Public Const SHGFI_SMALLICON = &H1                      '  get small icon
Public Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Public Const SHGFI_TYPENAME = &H400                     '  get type name
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long                      '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80          '  out: type name
End Type

Public SHInfo As SHFILEINFO
'shinfo keeps the info from each file...
