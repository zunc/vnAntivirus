Attribute VB_Name = "modFunction"
Option Explicit
'Public FolderPath As String
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
'Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal PassZero As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal PassZero As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hthread As Long, ByVal dwExitCode As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

