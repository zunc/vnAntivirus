Attribute VB_Name = "modWatch"
Option Explicit
'Get the directory chages using ReadDirectoryChangesW
Global WatchStart As Boolean
Global DirHndl As Long
Private Const FILE_NOTIF_GLOB = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                FILE_NOTIFY_CHANGE_FILE_NAME Or _
                                FILE_NOTIFY_CHANGE_DIR_NAME Or _
                                FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                FILE_NOTIFY_CHANGE_LAST_WRITE
    Private nBufLen As Long
    Private nReadLen As Long
    Private sAction As String
    Private fiBuffer As FILE_NOTIFY_INFORMATION
    Private cBuffer() As Byte
    Private cBuff2() As Byte
    Private lpBuf As Long
Public PathDec As String
    
Public Function GetDirHndl(ByVal PathDir As String) As Long
 On Error Resume Next
 Dim hDir As Long
 If Right(PathDir, 1) <> "\" Then PathDir = PathDir + "\"
 hDir = CreateFile(PathDir, FILE_LIST_DIRECTORY, FILE_SHARE_READ + FILE_SHARE_WRITE + FILE_SHARE_DELETE, _
                   ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OVERLAPPED, ByVal 0&)
 GetDirHndl = hDir
End Function

Public Sub StartWatch()
 If (DirHndl = 0) Or (DirHndl = -1) Then Exit Sub
    nBufLen = 1024
    ReDim cBuffer(0 To nBufLen)
    Call ReadDirectoryChangesW(DirHndl, cBuffer(0), nBufLen, WSubFolder, _
                        FILE_NOTIF_GLOB, nReadLen, 0, 0)
End Sub
Public Function GetChanges() As String
   On Error Resume Next
   Dim fName As String
   MoveMemory fiBuffer.NextEntryOffset, cBuffer(0), 4
   MoveMemory fiBuffer.Action, cBuffer(4), 4
   MoveMemory fiBuffer.FileNameLength, cBuffer(8), 4
   ReDim cBuff2(0 To fiBuffer.FileNameLength)
   MoveMemory cBuff2(0), cBuffer(12), fiBuffer.FileNameLength
   fiBuffer.FileName = cBuff2
   
   If fiBuffer.Action = FILE_ACTION_ADDED Then GetChanges = PathDec & fiBuffer.FileName
   'Select Case fiBuffer.Action
   '         Case FILE_ACTION_ADDED
   '             sAction = "Added file"
   '         Case FILE_ACTION_REMOVED
   '             sAction = "Removed file"
   '         Case FILE_ACTION_MODIFIED
   '             sAction = "Modified file"
   '         Case FILE_ACTION_RENAMED_OLD_NAME
   '             sAction = "Renamed from"
   '         Case FILE_ACTION_RENAMED_NEW_NAME
   '             sAction = "Renamed to"
   '         Case Else
   '             sAction = "Unknown"
   'End Select
   'fName = sAction + "-" + pathdec + fiBuffer.FileName
   'If sAction <> "Unknown" Then GetChanges = fName
End Function
      
Public Sub ClearHndl(Handle As Long)
 CloseHandle Handle
 Handle = 0
End Sub





