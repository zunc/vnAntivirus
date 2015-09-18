Attribute VB_Name = "basSysTray"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

'Code this module  from PSC

Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Const GWL_USERDATA = (-21&)
Public Const GWL_WNDPROC = (-4&)
Public Const WM_USER = &H400&
Public Const TRAY_CALLBACK = (WM_USER + 101&)
Public Const NIM_ADD = &H0&
Public Const NIM_MODIFY = &H1&
Public Const NIM_DELETE = &H2&
Public Const NIF_MESSAGE = &H1&
Public Const NIF_ICON = &H2&
Public Const NIF_TIP = &H4&
Public Const WM_MOUSEMOVE = &H200&
Public Const WM_LBUTTONDOWN = &H201&
Public Const WM_LBUTTONUP = &H202&
Public Const WM_LBUTTONDBLCLK = &H203&
Public Const WM_RBUTTONDOWN = &H204&
Public Const WM_RBUTTONUP = &H205&
Public Const WM_RBUTTONDBLCLK = &H206&
Public Const BDR_RAISEDOUTER = &H1&
Public Const BDR_RAISEDINNER = &H4&
Public Const BF_LEFT = &H1&
Public Const BF_TOP = &H2&
Public Const BF_RIGHT = &H4&
Public Const BF_BOTTOM = &H8&
Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Public Const BF_SOFT = &H1000&
Public Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Public PrevWndProc As Long
'------------------------------------------------------------
Public Function SubWndProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'------------------------------------------------------------
Dim SysTray As cSysTray
Dim ClassAddr As Long
'------------------------------------------------------------
Select Case MSG
Case TRAY_CALLBACK
ClassAddr = GetWindowLong(hWnd, GWL_USERDATA)
CopyMemory SysTray, ClassAddr, 4

SysTray.SendEvent lParam, wParam

CopyMemory SysTray, 0&, 4
End Select

SubWndProc = CallWindowProc(PrevWndProc, hWnd, MSG, wParam, lParam)
'------------------------------------------------------------
End Function
'------------

