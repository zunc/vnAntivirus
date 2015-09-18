VERSION 5.00
Begin VB.UserControl cSysTray 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "cSysTray.ctx":0000
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      Picture         =   "cSysTray.ctx":0312
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "cSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private gInTray As Boolean
Private gTrayId As Long
Private gTrayTip As String
Private gTrayHwnd As Long
Private gTrayIcon As StdPicture
Private gAddedToTray As Boolean
Const MAX_SIZE = 510
Private Const defInTray = False
Private Const defTrayTip = "SendLan v1.0" & vbNullChar
Private Const sInTray = "InTray"
Private Const sTrayIcon = "TrayIcon"
Private Const sTrayTip = "TrayTip"
Public Event MouseMove(Id As Long)
Public Event MouseDown(Button As Integer, Id As Long)
Public Event MouseUp(Button As Integer, Id As Long)
Public Event MouseDblClick(Button As Integer, Id As Long)
'-------------------------------------------------------
Private Sub UserControl_Initialize()
'-------------------------------------------------------
gInTray = defInTray
gAddedToTray = False
gTrayId = 0
gTrayHwnd = hWnd
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub UserControl_InitProperties()
'-------------------------------------------------------
InTray = defInTray
TrayTip = defTrayTip
Set TrayIcon = Picture
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub UserControl_Paint()
'-------------------------------------------------------
Dim edge As RECT
'-------------------------------------------------------
edge.Left = 0
edge.Top = 0
edge.Bottom = ScaleHeight
edge.Right = ScaleWidth
DrawEdge hDC, edge, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'-------------------------------------------------------
With PropBag
InTray = .ReadProperty(sInTray, defInTray)
Set TrayIcon = .ReadProperty(sTrayIcon, Picture)
TrayTip = .ReadProperty(sTrayTip, defTrayTip)
End With
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'-------------------------------------------------------
With PropBag
.WriteProperty sInTray, gInTray
.WriteProperty sTrayIcon, gTrayIcon
.WriteProperty sTrayTip, gTrayTip
End With
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub UserControl_Resize()
'-------------------------------------------------------
Height = MAX_SIZE
Width = MAX_SIZE
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub UserControl_Terminate()
'-------------------------------------------------------
If InTray Then
InTray = False
End If
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Public Property Set TrayIcon(Icon As StdPicture)
'-------------------------------------------------------
Dim Tray As NOTIFYICONDATA
Dim rc As Long
'-------------------------------------------------------
If Not (Icon Is Nothing) Then
If (Icon.Type = vbPicTypeIcon) Then
If gAddedToTray Then
Tray.uID = gTrayId
Tray.hWnd = gTrayHwnd
Tray.hIcon = Icon.Handle
Tray.uFlags = NIF_ICON
Tray.cbSize = Len(Tray)

rc = Shell_NotifyIcon(NIM_MODIFY, Tray)
End If

Set gTrayIcon = Icon
Set Picture = Icon
PropertyChanged sTrayIcon
End If
End If
'-------------------------------------------------------
End Property
'-------------------------------------------------------
'-------------------------------------------------------
Public Property Get TrayIcon() As StdPicture
'-------------------------------------------------------
Set TrayIcon = gTrayIcon
'-------------------------------------------------------
End Property
'-------------------------------------------------------
'-------------------------------------------------------
Public Property Let TrayTip(Tip As String)
'-------------------------------------------------------
Dim Tray As NOTIFYICONDATA
Dim rc As Long
'-------------------------------------------------------
If gAddedToTray Then
Tray.uID = gTrayId
Tray.hWnd = gTrayHwnd
Tray.szTip = Tip & vbNullChar
Tray.uFlags = NIF_TIP
Tray.cbSize = Len(Tray)

rc = Shell_NotifyIcon(NIM_MODIFY, Tray)
End If

gTrayTip = Tip
PropertyChanged sTrayTip
'-------------------------------------------------------
End Property
'-------------------------------------------------------
'-------------------------------------------------------
Public Property Get TrayTip() As String
'-------------------------------------------------------
TrayTip = gTrayTip
'-------------------------------------------------------
End Property
'-------------------------------------------------------
'-------------------------------------------------------
Public Property Let InTray(Show As Boolean)
'-------------------------------------------------------
Dim ClassAddr As Long
'-------------------------------------------------------
If (Show <> gInTray) Then
If Show Then
If Ambient.UserMode Then
PrevWndProc = SetWindowLong(gTrayHwnd, GWL_WNDPROC, AddressOf SubWndProc)


SetWindowLong gTrayHwnd, GWL_USERDATA, ObjPtr(Me)

AddIcon gTrayHwnd, gTrayId, TrayTip, TrayIcon
gAddedToTray = True
End If
Else
If gAddedToTray Then
DeleteIcon gTrayHwnd, gTrayId

SetWindowLong gTrayHwnd, GWL_WNDPROC, PrevWndProc
gAddedToTray = False
End If
End If

gInTray = Show
PropertyChanged sInTray
End If
'-------------------------------------------------------
End Property
'-------------------------------------------------------
'-------------------------------------------------------
Public Property Get InTray() As Boolean
'-------------------------------------------------------
InTray = gInTray
'-------------------------------------------------------
End Property
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub AddIcon(hWnd As Long, Id As Long, Tip As String, Icon As StdPicture)
'-------------------------------------------------------
Dim Tray As NOTIFYICONDATA
Dim tFlags As Long
Dim rc As Long
'-------------------------------------------------------
Tray.uID = Id
Tray.hWnd = hWnd

If Not (Icon Is Nothing) Then
Tray.hIcon = Icon.Handle
Tray.uFlags = Tray.uFlags Or NIF_ICON
Set gTrayIcon = Icon
End If

If (Tip <> "") Then
Tray.szTip = Tip & vbNullChar
Tray.uFlags = Tray.uFlags Or NIF_TIP
gTrayTip = Tip
End If

Tray.uCallbackMessage = TRAY_CALLBACK
Tray.uFlags = Tray.uFlags Or NIF_MESSAGE
Tray.cbSize = Len(Tray)

rc = Shell_NotifyIcon(NIM_ADD, Tray)
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub DeleteIcon(hWnd As Long, Id As Long)
'-------------------------------------------------------
Dim Tray As NOTIFYICONDATA
Dim rc As Long
'-------------------------------------------------------
Tray.uID = Id
Tray.hWnd = hWnd
Tray.uFlags = 0&
Tray.cbSize = Len(Tray)

rc = Shell_NotifyIcon(NIM_DELETE, Tray)
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-------------------------------------------------------
Friend Sub SendEvent(MouseEvent As Long, Id As Long)
'-------------------------------------------------------
Select Case MouseEvent
Case WM_MOUSEMOVE
RaiseEvent MouseMove(Id)
Case WM_LBUTTONDOWN
RaiseEvent MouseDown(vbLeftButton, Id)
Case WM_LBUTTONUP
RaiseEvent MouseUp(vbLeftButton, Id)
Case WM_LBUTTONDBLCLK
RaiseEvent MouseDblClick(vbLeftButton, Id)
Case WM_RBUTTONDOWN
RaiseEvent MouseDown(vbRightButton, Id)
Case WM_RBUTTONUP
RaiseEvent MouseUp(vbRightButton, Id)
Case WM_RBUTTONDBLCLK
RaiseEvent MouseDblClick(vbRightButton, Id)
End Select
'-------------------------------------------------------
End Sub

