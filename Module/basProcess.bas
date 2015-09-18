Attribute VB_Name = "basProcess"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

'Module nay toi chan thanh cam on : PhamTienSinh (PhamTrungHai) va PhuongThanh37
'Code by PhuongThanh37

Option Explicit

Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const ANYSIZE_ARRAY = 1
Public Const SE_DEBUG_NAME = "SeDebugPrivilege"

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

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

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Sub GetDebugPriv()
Dim hToken As Long
Dim sedebugnameValue As LUID
Dim tkp As TOKEN_PRIVILEGES, mNewPriv As TOKEN_PRIVILEGES

If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
    If LookupPrivilegeValue(vbNullString, SE_DEBUG_NAME, sedebugnameValue) Then
        tkp.PrivilegeCount = 1
        tkp.Privileges(0).pLuid = sedebugnameValue
        tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        AdjustTokenPrivileges hToken, False, tkp, Len(tkp), mNewPriv, Len(mNewPriv)
        CloseHandle (hToken)
    Else
        CloseHandle (hToken)
        Exit Sub
    End If
End If
End Sub
Function KillProcessById(lPID As Long)

Dim hnd As Long, t1 As Long
    hnd = OpenProcess(&H1&, 0, lPID)
    Call GetDebugPriv
    Call GetExitCodeProcess(hnd, t1)
    Call TerminateProcess(hnd, t1)
    Call CloseHandle(hnd)
End Function
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
Public Function CheckProcess(FilePath As String) As Long
    CheckProcess = 0
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
  Dim strTMP As String
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
   While theloop <> 0

      theloop = ProcessNext(snap, proc)
      strTMP = ProcessPathByPID(proc.th32ProcessID)
      If strTMP <> "SYSTEM" Then
            If strTMP = FilePath Then CheckProcess = proc.th32ProcessID: GoTo KetThuc
      End If
   Wend
   CloseHandle snap
KetThuc:
End Function
Public Function CheckID(ID As Long) As Long
    CheckID = 0
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
   While theloop <> 0
      theloop = ProcessNext(snap, proc)
        If proc.th32ProcessID = ID Then CheckID = ID: GoTo KetThuc
   Wend
   CloseHandle snap
KetThuc:
End Function
