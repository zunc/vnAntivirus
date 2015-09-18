VERSION 5.00
Begin VB.UserControl ucFirefoxWait 
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   ToolboxBitmap   =   "Firefox Wait.ctx":0000
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2640
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   840
      Width           =   240
   End
   Begin VB.PictureBox picOriginal 
      Height          =   300
      Left            =   240
      Picture         =   "Firefox Wait.ctx":0312
      ScaleHeight     =   240
      ScaleWidth      =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   2220
   End
   Begin VB.Timer tmrPlay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   360
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   240
      Picture         =   "Firefox Wait.ctx":1E54
      ScaleHeight     =   240
      ScaleWidth      =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   2220
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "ucFirefoxWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  ________________________________                          _______
' / ucFirefoxWait                  \________________________/ v1.01 |
' |                                                                 |
' |     Control Desc.:  A control that mimics the wait shown by     |
' |                     the Firefox web browser when navigating to  |
' |                     a web object.                               |
' |                                                                 |
' |   Original Author:  CubeSolver                                  |
' |      Date Created:  November 12, 2004                           |
' |      OS Tested On:  Windows NT 4 SP 6a, Windows XP              |
' |           Credits:  Code used is noted where applicable.        |
' |                     Mozilla for creating the Firefox web        |
' |                     browser. Get more info about their browser  |
' |                     at http://www.mozilla.org/products/firefox/ |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.01     CubeSolver  Nov 17, 2004  Added IsPlaying and        |
'  |                                    BackColor properties.      |
'  | 1.00     CubeSolver  Nov 12, 2004  Original release.          |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private lFrameNum As Long                   ' Pointer to the frame for showing
Private l_def_BackColor As Long

Private m_BackColor As Long                 ' Background color of the control
Private m_IsPlaying As Boolean              ' Tell whether the animation is playing or not
Private m_Speed As Long                     ' The timer interval

Private Const m_def_IsPlaying As Boolean = False
Private Const m_def_Speed As Long = 100

Private Type RECT
  lLeft As Long
  lTop As Long
  lRight As Long
  lBottom As Long
End Type

Private Const SRCAND As Long = &H8800C6
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCINVERT As Long = &H660046
Private Const SRCPAINT As Long = &HEE0086

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Sub ColorBackground()
  picSource.Cls
  picSource.Picture = picOriginal.Picture

  Call ReplaceColor(picSource, RGB(255, 0, 255), m_BackColor)     ' Make the background color the same as the form's (RGB(255, 0, 255) is our mask color)
End Sub
Private Function CreateDC(ByRef picThis As VB.PictureBox, ByVal lW As Long, ByVal lH As Long, ByRef lhDC As Long, ByRef lhBmp As Long, _
       ByRef lhBmpOld As Long, Optional ByVal bMono As Boolean = False) As Boolean
  ' Code taken from:
  ' http://www.vbaccelerator.com/home/VB/Tips/Replace_One_Colour_With_Another_in_a_Picture/article.asp
  If (bMono) Then
    lhDC = CreateCompatibleDC(0)
  Else
    lhDC = CreateCompatibleDC(picThis.hDC)
  End If

  If (lhDC <> 0) Then
    If (bMono) Then
      lhBmp = CreateCompatibleBitmap(lhDC, lW, lH)
    Else
      lhBmp = CreateCompatibleBitmap(picThis.hDC, lW, lH)
    End If
    If (lhBmp <> 0) Then
      lhBmpOld = SelectObject(lhDC, lhBmp)
      CreateDC = True
    Else
      Call DeleteObject(lhDC)
      lhDC = 0
    End If
  End If
End Function
Private Sub DisplayImage()
  Dim tRect As RECT

  ' Loop through the image strip
  If lFrameNum = 7 Then
    lFrameNum = 0
  Else
    lFrameNum = lFrameNum + 1
  End If

  Call GetClientRect(picDestination.hWnd, tRect)        ' Get our client rectangle

  picDestination.Cls                      ' Prep the picture box - required

  ' Read from the image strip at certain coordinates and paint into the destination picture box
  Call BitBlt(picDestination.hDC, 0, 0, tRect.lRight - tRect.lLeft, tRect.lBottom - tRect.lTop, picSource.hDC, lFrameNum * 16, 0, SRCCOPY)
End Sub
Private Sub ReplaceColor(ByRef picThis As VB.PictureBox, ByVal lFromColor As Long, ByVal lToColor As Long)
  ' Code taken from:
  ' http://www.vbaccelerator.com/home/VB/Tips/Replace_One_Colour_With_Another_in_a_Picture/article.asp
  Dim lW As Long, lH As Long
  Dim lMaskDC As Long, lMaskBMP As Long, lMaskBMPOLd As Long
  Dim lCopyDC As Long, lCopyBMP As Long, lCopyBMPOLd As Long
  Dim tR As RECT
  Dim hBr As Long

  ' Cache the width & height of the picture
  lW = picThis.ScaleWidth \ Screen.TwipsPerPixelX
  lH = picThis.ScaleHeight \ Screen.TwipsPerPixelY
  ' Create a Mono DC & Bitmap
  If (CreateDC(picThis, lW, lH, lMaskDC, lMaskBMP, lMaskBMPOLd, True)) Then
    ' Create a DC & Bitmap with the same color depth as the picture
    If (CreateDC(picThis, lW, lH, lCopyDC, lCopyBMP, lCopyBMPOLd)) Then
      ' Make a mask from the picture which is white in the replace color area
      Call SetBkColor(picThis.hDC, lFromColor)
      Call BitBlt(lMaskDC, 0, 0, lW, lH, picThis.hDC, 0, 0, SRCCOPY)

      ' Fill the color DC with the color we want to replace with
      tR.lRight = lW
      tR.lBottom = lH
      hBr = CreateSolidBrush(lToColor)
      Call FillRect(lCopyDC, tR, hBr)
      Call DeleteObject(hBr)
      ' Turn the color DC black except where the mask is white
      Call BitBlt(lCopyDC, 0, 0, lW, lH, lMaskDC, 0, 0, SRCAND)

      ' Create an inverted mask, so it is black where the
      ' color is to be replaced but white otherwise
      hBr = CreateSolidBrush(&HFFFFFF)
      Call FillRect(lMaskDC, tR, hBr)
      Call DeleteObject(hBr)
      Call BitBlt(lMaskDC, 0, 0, lW, lH, picThis.hDC, 0, 0, SRCINVERT)

      ' AND the inverted mask with the picture. The picture
      ' goes black where the color is to be replaced, but is
      ' unaffected otherwise
      Call SetBkColor(picThis.hDC, &HFFFFFF)
      Call BitBlt(picThis.hDC, 0, 0, lW, lH, lMaskDC, 0, 0, SRCAND)

      ' Finally, OR the colored item with the picture. Where
      ' the picture is black and the colored DC isn't, the
      ' color will be transferred
      Call BitBlt(picThis.hDC, 0, 0, lW, lH, lCopyDC, 0, 0, SRCPAINT)
      picThis.Refresh

      ' Clear up the color DC
      Call SelectObject(lCopyDC, lCopyBMPOLd)
      Call DeleteObject(lCopyBMP)
      Call DeleteObject(lCopyDC)
    End If

    ' Clear up the mask DC
    Call SelectObject(lMaskDC, lMaskBMPOLd)
    Call DeleteObject(lMaskBMP)
    Call DeleteObject(lMaskDC)
  End If
End Sub
Private Sub ShowIdle()
  Dim tRect As RECT

  Call GetClientRect(picDestination.hWnd, tRect)    ' Get the dimensions of the picture box

  picDestination.Cls                                ' Prep the picture box - required

  ' Read the idle image from the image strip and paint into the destination picture box
  Call BitBlt(picDestination.hDC, 0, 0, tRect.lRight - tRect.lLeft, tRect.lBottom - tRect.lTop, picSource.hDC, 8 * 16, 0, SRCCOPY)
End Sub
Public Sub PlayWait()
Attribute PlayWait.VB_Description = "Play the animation"
  ' Start the timer
  If Not m_IsPlaying Then                           ' Start the animation only if it's not already running
    IsPlaying = True
    tmrPlay.Interval = m_Speed
    tmrPlay.Enabled = IsPlaying
  End If
End Sub
Public Sub StopWait()
Attribute StopWait.VB_Description = "Return to an idle state"
  ' Stop the timer
  If m_IsPlaying Then                               ' Stop the animation only if it's currently running
    IsPlaying = False
    tmrPlay.Enabled = IsPlaying
    lFrameNum = 0                                   ' So the animation will always start in the same position

    Call ShowIdle
  End If
End Sub
Private Sub tmrPlay_Timer()
  Call DisplayImage
End Sub
Private Sub UserControl_Initialize()
  ' Grab the color of the form using picColor as a default color
  l_def_BackColor = GetPixel(picColor.hDC, 17 \ Screen.TwipsPerPixelX, 17 \ Screen.TwipsPerPixelY)

  Call ShowIdle
End Sub
Private Sub UserControl_InitProperties()
  m_Speed = m_def_Speed
  m_IsPlaying = m_def_IsPlaying
  m_BackColor = l_def_BackColor
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error Resume Next

  m_Speed = PropBag.ReadProperty("AnimationSpeed", m_def_Speed)
  m_BackColor = PropBag.ReadProperty("BackColor", l_def_BackColor)
  m_IsPlaying = PropBag.ReadProperty("IsPlaying", m_def_IsPlaying)
End Sub
Private Sub UserControl_Resize()
  ' No use allowing any size changes
  UserControl.Width = 240
  UserControl.Height = 240
End Sub
Private Sub UserControl_Show()
  Call ColorBackground
  Call ShowIdle
End Sub
Private Sub UserControl_Terminate()
  tmrPlay.Enabled = False
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("AnimationSpeed", m_Speed, m_def_Speed)
  Call PropBag.WriteProperty("BackColor", m_BackColor, l_def_BackColor)
  Call PropBag.WriteProperty("IsPlaying", m_IsPlaying, m_def_IsPlaying)
End Sub
Public Property Get AnimationSpeed() As Long
Attribute AnimationSpeed.VB_Description = "Time in milliseconds before each frame in the animation displays. Use larger numbers for slower speed."
  AnimationSpeed = m_Speed
End Property
Public Property Let AnimationSpeed(ByVal l_Milliseconds As Long)
  ' Property for setting the speed of the animation. Larger numbers equal slower speeds
  m_Speed = l_Milliseconds
  Call PropertyChanged("AnimationSpeed")
  tmrPlay.Interval = m_Speed
End Property
Public Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal l_BackColor As OLE_COLOR)
  ' Property for changing the background color of the control
  If l_BackColor < 0 Then l_BackColor = l_def_BackColor
  m_BackColor = l_BackColor
  Call PropertyChanged("BackColor")
  Call ColorBackground
  Call ShowIdle
End Property
Public Property Get IsPlaying() As Boolean
Attribute IsPlaying.VB_Description = "Returns/sets a value that determines whether the animation is running or not."
  IsPlaying = m_IsPlaying
End Property
Public Property Let IsPlaying(ByVal bPlaying As Boolean)
  ' Property for determining whether the animation is currently playing or not
  m_IsPlaying = bPlaying
  PropertyChanged "IsPlaying"
End Property
