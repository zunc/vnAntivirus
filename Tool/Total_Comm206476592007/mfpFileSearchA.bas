Attribute VB_Name = "mfpFileSearchA"
Option Explicit

Public Type mctFileInfoType
    FilePath As String
    FileName As String
    UNC As String
    Extension As String
    SIZE As Currency
    ReadOnly As Boolean
    Hidden As Boolean
        Archive As Boolean
    System As Boolean

    CreationDate As String
End Type

Public Type mctFileSearchResults
    FileCount As Long
    FileSize As Currency
    Files() As mctFileInfoType
End Type

Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
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
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Function GetFileSize_(ByVal iFileSizeHigh As Long, ByVal iFileSizeLow As Long) As Currency

    Dim curFileSizeHigh As Currency
    Dim curFileSizeLow As Currency
    Dim curFileSize As Currency

    curFileSizeHigh = CCur(iFileSizeHigh)
    curFileSizeLow = CCur(iFileSizeLow)

    curFileSize = curFileSizeLow

    If curFileSizeLow < 0 Then
        curFileSize = curFileSize + 4294967296@
    End If

    If curFileSizeHigh > 0 Then
        curFileSize = curFileSize + (curFileSizeHigh * 4294967296@)
    End If

    GetFileSize_ = curFileSize

End Function
Public Sub FileSearchA(ByVal sPath As String, ByVal sFileMask As String, ByRef taFiles As mctFileSearchResults, _
                       Optional ByVal bRecursive As Boolean = False, Optional ByVal iRecursionLevel As Long = -1)
On Error GoTo Hell

Dim sFilename As String
Dim sFolder As String
Dim iFolderCount As Long
Dim aFolders() As String
Dim aFileMask() As String
Dim iSearchHandle As Long
Dim WFD As WIN32_FIND_DATA
Dim bContinue As Long: bContinue = True
Dim Ret As Long, X As Long
Dim tSystemTime As SYSTEMTIME

    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    ' Search for subdirectories first and save'em for later
    ' --------------------------
    If bRecursive Then
        iSearchHandle = FindFirstFile(sPath & "*.", WFD)
    
        If iSearchHandle <> INVALID_HANDLE_VALUE Then
            Do While bContinue
                
                If (InStr(WFD.cFileName, Chr(0)) > 0) Then WFD.cFileName = Left(WFD.cFileName, InStr(WFD.cFileName, Chr(0)) - 1)
                sFolder = Trim$(WFD.cFileName)
                
                If (sFolder <> ".") And (sFolder <> "..") Then ' Ignore the current and encompassing directories
                    If WFD.dwFileAttributes And vbDirectory Then
                      
                
                        iFolderCount = iFolderCount + 1
                        ReDim Preserve aFolders(iFolderCount)
                        aFolders(iFolderCount) = sFolder
                   End If
                    End If
                
                
                bContinue = FindNextFile(iSearchHandle, WFD) 'Get next subdirectory.
            
            Loop
            bContinue = FindClose(iSearchHandle)
        End If
    End If
    ' --------------------------
    
    bContinue = True
    
    ' Walk through this directory and sum file sizes.
    ' --------------------------
    
    ' FindFirstFile takes one type at a time, so we'll loop the search for as many extensions as specified
    aFileMask = Split(sFileMask, ";")
    For X = 0 To UBound(aFileMask)
        
        ' Make sure it's all formatted
      '  If Left$(aFileMask(X), 1) = "." Then
      '      aFileMask(X) = "*" & aFileMask(X)
       ' ElseIf Left$(aFileMask(X), 2) <> "*." Then
       '     aFileMask(X) = "*." & aFileMask(X)
       ' End If
        
        iSearchHandle = FindFirstFile(sPath & aFileMask(X), WFD)
    
        If iSearchHandle <> INVALID_HANDLE_VALUE Then
            Do While bContinue
                
                If (InStr(WFD.cFileName, Chr(0)) > 0) Then WFD.cFileName = Left(WFD.cFileName, InStr(WFD.cFileName, Chr(0)) - 1)
                sFilename = Trim$(WFD.cFileName)
                
                ' It's a file, right?
                If (sFilename <> ".") And (sFilename <> "..") And (Not (WFD.dwFileAttributes And vbDirectory) = vbDirectory) Then
                    With taFiles
                        .FileSize = .FileSize + GetFileSize_(WFD.nFileSizeHigh, WFD.nFileSizeLow)
                        .FileCount = .FileCount + 1
                        ReDim Preserve .Files(.FileCount)
                        
                        
                        
                        With .Files(.FileCount)
                            If InStr(sFilename, ".") > 0 Then
                            .Extension = Mid$(sFilename, InStrRev(sFilename, ".") + 1)
                                                       .FileName = sFilename
 
                            Else
                            .Extension = ""
                                                        .FileName = sFilename & "."

                            End If
                           
                            .FilePath = sPath
                           
                            .ReadOnly = (WFD.dwFileAttributes And vbReadOnly) = vbReadOnly
                            .Hidden = (WFD.dwFileAttributes And vbHidden) = vbHidden
                            .Archive = (WFD.dwFileAttributes And vbArchive) = vbArchive
                            .System = (WFD.dwFileAttributes And vbSystem) = vbSystem
                            
                            .SIZE = GetFileSize_(WFD.nFileSizeHigh, WFD.nFileSizeLow)
                            .UNC = sPath & sFilename
                            If FileTimeToSystemTime(WFD.ftCreationTime, tSystemTime) Then .CreationDate = tSystemTime.wDay & "/" & tSystemTime.wMonth & "/" & tSystemTime.wYear & " " & IIf(tSystemTime.wHour > 12, tSystemTime.wHour - 12 & ":" & tSystemTime.wMinute & "PM", tSystemTime.wHour & ":" & tSystemTime.wMinute & " AM")
                        End With
                    End With
                    
                    
                End If
                bContinue = FindNextFile(iSearchHandle, WFD) ' Get next file
            Loop
            bContinue = FindClose(iSearchHandle)
        End If
    Next
    ' --------------------------
    
    ' If there are sub-directories,
    If iFolderCount > 0 Then
        ' And if we care,
        If bRecursive Then
            If iRecursionLevel <> 0 Then ' Recursively walk into them...
                iRecursionLevel = iRecursionLevel - 1
                For X = 1 To iFolderCount
                
                    FileSearchA sPath & aFolders(X) & "\", sFileMask, taFiles, bRecursive, iRecursionLevel
                Next X
            End If
        End If
    End If
    
' --------------------------------------------------------------------------
Exit Sub
Hell:
    Debug.Print Err.Description: Stop: Resume
End Sub


