Attribute VB_Name = "basDetectFile"
'Module from PSC

'Module detect changes file in a folder
'Module nhan dang nhung thay doi file trong thu muc
Option Explicit

Public FolderPath As String
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Type FILE_NOTIFY_INFORMATION
   NextEntryOffset As Long
   Action As Long
   FileNameLength As Long
   FileName As String
End Type
Public WSubFolder  As Boolean
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_LIST_DIRECTORY = &H1
Public Const FILE_SHARE_READ = &H1&
Public Const FILE_SHARE_DELETE = &H4&
Public Const OPEN_EXISTING = 3
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1&
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10&
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Public Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2

Public Const FILE_ACTION_ADDED = &H1&
Public Const FILE_ACTION_REMOVED = &H2&
Public Const FILE_ACTION_MODIFIED = &H3&
Public Const FILE_ACTION_RENAMED_OLD_NAME = &H4&
Public Const FILE_ACTION_RENAMED_NEW_NAME = &H5&


Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpcSource As Any, ByVal dwLength As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadDirectoryChangesW Lib "kernel32" (ByVal hDirectory As Long, lpBuffer As Any, ByVal nBufferLength As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, lpBytesReturned As Long, ByVal PassZero As Long, ByVal PassZero As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal PassZero As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal PassZero As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

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
   Select Case fiBuffer.Action
            Case FILE_ACTION_ADDED
                sAction = "Added file"
            Case FILE_ACTION_REMOVED
                sAction = "Removed file"
            Case FILE_ACTION_MODIFIED
                sAction = "Modified file"
            Case FILE_ACTION_RENAMED_OLD_NAME
                sAction = "Renamed from"
            Case FILE_ACTION_RENAMED_NEW_NAME
                sAction = "Renamed to"
            Case Else
                sAction = "Unknown"
   End Select
   fName = sAction + "|" & FolderPath + fiBuffer.FileName
   
   If sAction <> "Unknown" Then GetChanges = fName
End Function
      
Public Sub ClearHndl(Handle As Long)
 CloseHandle Handle
 Handle = 0
End Sub
