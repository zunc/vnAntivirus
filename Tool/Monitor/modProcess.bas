Attribute VB_Name = "modProcess"
'To suspend or resume thread
Public Declare Function SuspendThread Lib "kernel32" (ByVal hthread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hthread As Long) As Long
'To open a thread
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwProcessId As Long) As Long
Public Declare Function Thread32First Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
'To open a thread
Public Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
'THREAD_SUSPEND_RESUME can be used instead of THREAD_ALL_ACCESS for suspending and resuming a thread
Public Const THREAD_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF
'Current running process PID or ID
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

' This function is not from ToolHelp but you need it to destroy a snapshot
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hsnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hsnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Const MAX_PATH As Integer = 260

'define PROCESSENTRY32 structure

Public Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * MAX_PATH
End Type
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_VM_READ = &H10
Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type
Function ProcessPathByPID(pid As Long) As String
'Return path to the executable from PID
'http://support.microsoft.com/default.aspx?scid=kb;en-us;187913
Dim cbNeeded As Long
Dim Modules(1 To 200) As Long
Dim Ret As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long

hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
    Or PROCESS_VM_READ, 0, pid)
            
If hProcess <> 0 Then
                
    Ret = EnumProcessModules(hProcess, Modules(1), _
        200, cbNeeded)
                
    If Ret <> 0 Then
        ModuleName = Space(MAX_PATH)
        nSize = 500
        Ret = GetModuleFileNameExA(hProcess, _
            Modules(1), ModuleName, nSize)
        ProcessPathByPID = Left(ModuleName, Ret)
    End If
End If
          
Ret = CloseHandle(hProcess)

If ProcessPathByPID = "" Then
    ProcessPathByPID = "SYSTEM"
End If

End Function
Public Function SuspendResumeProcess(ByVal procid As Long, ByVal suspendresume As Boolean) As Boolean
Dim hsnapshot As Long
Dim htthread As Long
Dim pthread As Boolean
Dim pt As THREADENTRY32

SuspendResumeProcess = False

hsnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPthread, 0)

pt.dwSize = Len(pt)

pthread = Thread32First(hsnapshot, pt)

While pthread
    If pt.th32OwnerProcessID = procid Then
        htthread = OpenThread(THREAD_ALL_ACCESS, 0, pt.th32ThreadID)
        If htthread <> 0 Then
            If suspendresume Then SuspendThread (htthread) Else ResumeThread (htthread)
            CloseHandle htthread
            SuspendResumeProcess = True
        End If
    End If
    pthread = Thread32Next(hsnapshot, pt)
Wend
CloseHandle hsnapshot
End Function
Public Sub KillProcessById(p_lngProcessId As Long)
  Dim lnghProcess As Long
  Dim lngReturn As Long
    
    lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
    lngReturn = TerminateProcess(lnghProcess, 0&)
    
End Sub
Public Function CheckProcess(PathFile As String) As String
'On Error Resume Next
'---------Liet ke process-------
  CheckProcess = ""
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
  Dim exename As String
  Dim ID As Long
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
   While theloop <> 0

      ID = proc.th32ProcessID
      theloop = ProcessNext(snap, proc)
      If ProcessPathByPID(proc.th32ProcessID) <> "SYSTEM" Then
      'MsgBox ProcessPathByPID(proc.th32ProcessID)
        If DungLuong(PathFile) = DungLuong(ProcessPathByPID(proc.th32ProcessID)) Then
            If SoSanh(PathFile, ProcessPathByPID(proc.th32ProcessID)) = True Then CheckProcess = ProcessPathByPID(proc.th32ProcessID) & "|" & proc.th32ProcessID
        End If
        End If
   Wend
   CloseHandle snap
End Function
