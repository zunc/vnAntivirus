Attribute VB_Name = "HiddenFolders"
Private Const MAX_PATH = 260
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const LB_GETCOUNT = &H18B
Private Const LB_INSERTSTRING = &H181
Private Const LB_ERR = (-1)

'private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
'private Const FILE_ATTRIBUTE_ARCHIVE = &H20
'private Const FILE_ATTRIBUTE_NORMAL = &H80
'private Const FILE_ATTRIBUTE_TEMPORARY = &H100
'private Const FILE_ATTRIBUTE_COMPRESSED = &H800

Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long


'* bShowSystem - shows directories with both the hidden and system attributes set
'* non-hidden system directories are already shown and are not affected by this
'* you probably shouldn't show hidden system folders (such as the
'  recycle bin and it's raw files)
Public Sub ShowHiddenDirectories(DirCtrl As DirListBox, Optional bShowSystem As Boolean)
    Dim res As Long
    Dim sF As String, sDirPath
    Dim FData As WIN32_FIND_DATA
    Dim fHand As Long, i As Long
    Dim level As Long
    Dim StillOK As Long
    Const HIDDEN_DIRECTORY = FILE_ATTRIBUTE_DIRECTORY Or FILE_ATTRIBUTE_HIDDEN
'     Const LBS_SORT = &H2&

    sDirPath = DirCtrl.Path
    'append trailing slash
    If Right$(sDirPath, 1) <> "\" Then sDirPath = sDirPath & "\"
    
    'get dir path level (i.e. c:\windows\system = 3)
    'LB_GETCOUNT counts all items in dirbox while
    'DirCtrl.ListCount gives only subdirectories
        '    VB method
        '    i = -1
        '    Do While Len(DirCtrl.List(i))
        '        i = i - 1
        '    Loop
        '    level = Abs(i) - 1
    'api
    res = SendMessage(DirCtrl.hwnd, LB_GETCOUNT, 0, 0)
    If res = LB_ERR Then Exit Sub
    level = res - DirCtrl.ListCount
    
    'Find hidden directories
    fHand = FindFirstFile(sDirPath & "*", FData)
    StillOK = fHand
    Do While StillOK > 0
        'check if file is a folder and hidden
        If (FData.dwFileAttributes And HIDDEN_DIRECTORY) >= HIDDEN_DIRECTORY Then
           'continue if we don't care if folder has system attribute
           'or if the folder doesn't have system attribute
           If bShowSystem Or ((FData.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM) = 0) Then
            sF = CutRightAt(FData.cFileName)
            If sF <> "." And sF <> ".." Then
                'add the hidden folder to the dirbox
                
                'it is ordered automatically but incorrectly with LB_ADDSTRING
                'i.e. a hidden folder called A1 would be placed before C:\
'                res = SendMessageString(DirCtrl.hwnd, LB_ADDSTRING, 0, sF)
                
                'move backwards through dirbox and insert alphabetically
                'you could of course use a binary search but for any
                'Win9X system I've seen this would be overkill
                i = DirCtrl.ListCount
                Do
                    If i > 0 Then
                        'compare against part of DirBox's item that is foldername w/o path
                        res = StrComp(sF, Right(DirCtrl.List(i - 1), Len(DirCtrl.List(i - 1)) - Len(sDirPath)), vbTextCompare)
                        If res >= 0 Then 'found insertion point, now add folder
                            'don't add folder if the 2 strings are the same
                            'could occur if you invoked this sub a 2nd time w/o cd-ing
                            If res Then res = SendMessageString(DirCtrl.hwnd, LB_INSERTSTRING, i + level, sF)
                            Exit Do
                        'else keep looking for insertion point
                        End If
                    Else 'folder is first alphabetically
                        If i = 0 Then res = SendMessageString(DirCtrl.hwnd, LB_INSERTSTRING, i + level, sF)
                    End If
                    i = i - 1
                Loop While i >= 0
            End If 'not . or ..
           End If 'system
        End If 'hidden

        StillOK = FindNextFile(fHand, FData)
    Loop

    fHand = FindClose(fHand)
End Sub

'   typical TrimNull function with option to trim at any character
Private Function CutRightAt(NormString As String, Optional ascii As Long = 0) As String
    Dim i As Long
    i = InStr(1, NormString, Chr(ascii), vbBinaryCompare)
    If i Then
        CutRightAt = Left(NormString, i - 1)
    Else
        CutRightAt = NormString
    End If
End Function

