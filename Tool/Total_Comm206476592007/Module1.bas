Attribute VB_Name = "Module1"
Option Explicit








Public Type INITCOMMONCONTROLSEX_TYPE
    dwSize As Long
    dwICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As _
    INITCOMMONCONTROLSEX_TYPE) As Long
Public Const ICC_INTERNET_CLASSES = &H800


' GetWindowsLong Constants
Private Const GWL_WNDPROC = (-4)
' Windows Message Constants
Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2
' Column Header Notification Meassage Constants
Private Const HDN_FIRST = -300&
Private Const HDN_BEGINTRACK = (HDN_FIRST - 6)
' Column Header Item Info Message Constants
Private Const HDI_WIDTH = &H1
' Notify Message Header Type
Private Type NMHDR
hWndFrom As Long
idFrom As Long
code As Long
End Type
' Notify Message Header for Listview
Private Type NMHEADER
hdr As NMHDR
iItem As Long
iButton As Long
lPtrHDItem As Long ' HDITEM FAR* pItem
End Type
' Header Item Type
Private Type HDITEM
mask As Long
cxy As Long
pszText As Long
hbm As Long
cchTextMax As Long
fmt As Long
lParam As Long
iImage As Long
iOrder As Long
End Type
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private mlPrevWndProc As Long



















Private Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tNMH As NMHDR
Dim tNMHEADER As NMHEADER
Dim tITEM As HDITEM
Select Case Msg
Case WM_NOTIFY
' Copy the Notify Message Header to a Header Structure
CopyMemory tNMH, ByVal lParam, Len(tNMH)
Select Case tNMH.code
Case HDN_BEGINTRACK
' If the user is trying to Size a Column Header...
' Extract Information about the Header being Sized
CopyMemory tNMHEADER, ByVal lParam, Len(tNMHEADER)
' Get Item Info. about the header (i.e. Width)
CopyMemory tITEM, ByVal tNMHEADER.lPtrHDItem, Len(tITEM)
' Don't allow Zero Width Columns to be Sized.
If (tITEM.mask And HDI_WIDTH) = HDI_WIDTH And tITEM.cxy = 0 Then
WindowProc = 1
Exit Function
End If
End Select
Case WM_DESTROY
' Remove Subclassing when Listview is Destroyed (Form unloaded.)
WindowProc = CallWindowProc(mlPrevWndProc, hwnd, Msg, wParam, lParam)
Call SetWindowLong(hwnd, GWL_WNDPROC, mlPrevWndProc)
Exit Function
End Select
' Call Default Window Handler
WindowProc = CallWindowProc(mlPrevWndProc, hwnd, Msg, wParam, lParam)
End Function
Public Sub SubClassHwnd(ByVal hwnd As Long)
mlPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub








