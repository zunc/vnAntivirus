Attribute VB_Name = "basBrowse"
'Module from http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=65879&lngWId=1
Option Explicit

Public Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BROWSEINFO) As Long
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

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

Public Const BIF_NEWDIALOGSTYLE As Long = &H40
Public Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Public Const BIF_BROWSEFORPRINTER As Long = &H2000
Public Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Public Const BIF_BROWSEINCLUDEURLS As Long = &H80
Public Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Public Const BIF_EDITBOX As Long = &H10
Public Const BIF_RETURNFSANCESTORS As Long = &H8
Public Const BIF_RETURNONLYFSDIRS As Long = &H1
Public Const BIF_SHAREABLE As Long = &H8000
Public Const BIF_STATUSTEXT As Long = &H4
Public Const BIF_USENEWUI As Long = &H40
Public Const BIF_VALIDATE As Long = &H20
Public Const MAX_PATH As Long = 260

Public Const WM_USER = &H400
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
Public Const BFFM_INITIALIZED As Long = 1
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function

Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    If uMsg = BFFM_INITIALIZED Then SendMessage hWnd, BFFM_SETSELECTIONA, True, ByVal lpData
End Function

Public Function BrowseForFolder(ByVal Hndl As Long, sSelPath As String _
, ByVal Message As String, Optional IncludeFiles As Boolean, Optional UseNewLooks As Boolean _
, Optional DontGoBelowDomain As Boolean, Optional DontUseRootDirectory As Boolean _
, Optional DisplayEditBox As Boolean) As String

    Dim BIF As BROWSEINFO
    Dim pidl As Long
    Dim lpSelPath As Long
    Dim sPath As String * MAX_PATH
  
    With BIF
        .hOwner = Hndl
        .pidlRoot = 0
        .lpszTitle = Message
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        lpSelPath = LocalAlloc(LPTR, Len(sSelPath))
        CopyMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath)
        .lParam = lpSelPath
        .ulFlags = BIF_RETURNONLYFSDIRS _
        + IIf(IncludeFiles = True, BIF_BROWSEINCLUDEFILES, 0&) _
        + IIf(DontUseRootDirectory = True, BIF_RETURNFSANCESTORS, 0&) _
        + IIf(UseNewLooks = True, BIF_USENEWUI, 0&) _
        + IIf(DontGoBelowDomain = True, BIF_DONTGOBELOWDOMAIN, 0&) _
        + IIf(DisplayEditBox = True, BIF_EDITBOX, 0&)
        'Note: If you use BIF_USENEWUI or / And
        'BIF_EDITBOX, the Ok button will be always enabled.
     End With
    
    pidl = SHBrowseForFolder(BIF)
   
    If pidl Then
        If SHGetPathFromIDList(pidl, sPath) Then
           BrowseForFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
        CoTaskMemFree pidl
    End If
    LocalFree lpSelPath
End Function





