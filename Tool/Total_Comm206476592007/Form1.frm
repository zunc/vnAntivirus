VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer see"
   ClientHeight    =   7395
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicFiles16 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8160
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txtPath 
         Height          =   270
         Left            =   0
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   6840
         Width           =   7935
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   360
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   210
         Left            =   1320
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   345
         LargeChange     =   500
         Left            =   7440
         Max             =   1000
         Min             =   1
         SmallChange     =   500
         TabIndex        =   9
         Top             =   9830
         Value           =   1
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   7815
         Begin VB.CommandButton dcButton4 
            Caption         =   "..."
            Height          =   495
            Left            =   1080
            TabIndex        =   16
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton dcButton3 
            Caption         =   "\"
            Height          =   495
            Left            =   600
            TabIndex        =   15
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton dcButton2 
            Height          =   495
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   495
         End
         Begin MSComctlLib.StatusBar StatusBar2 
            Height          =   300
            Left            =   3840
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   529
            ShowTips        =   0   'False
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   10583
                  MinWidth        =   10583
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   480
         TabIndex        =   8
         Top             =   9000
         Width           =   7800
         Begin VB.CommandButton dcButton1 
            Height          =   975
            Index           =   0
            Left            =   240
            Picture         =   "Form1.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   735
         End
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   7080
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   10583
               MinWidth        =   10583
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   3528
               MinWidth        =   3528
               Picture         =   "Form1.frx":0C54
               Text            =   "*.*"
               TextSave        =   "*.*"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5895
         Left            =   0
         TabIndex        =   6
         Top             =   960
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   10398
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":11EE
         NumItems        =   0
      End
      Begin VB.Label barra 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Tag             =   "0"
         Top             =   720
         Width           =   7920
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ListView basu 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Tag             =   "0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   661
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long

Dim R As Long
Dim Item As ListItem
Dim imgObj As ListImage
Dim hSIcon As Long
Dim oldSortkey As Integer
Dim oldSortorder As Integer
Dim Button As Button

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim FrameChildren As New Collection
Dim LastScrollValue As Single

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long






Private Sub barra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim pt As POINTAPI
' Dim dl As Long
' Dim lCurX As Long
' Dim lCurY As Long
' Dim sFilePath As String
' On Error Resume Next
    
'Select Case Button
'Case 2
'dl = GetCursorPos(pt)
    
'    sFilePath = barra.Caption & Chr$(0)
    
'    dl = DoExplorerMenu((Me.Hwnd), sFilePath, pt.X, pt.Y)
    
'End Select

End Sub








Private Sub Command1_Click()
End Sub

Private Sub dcButton1_Click(Index As Integer)


On Error GoTo pepe
     
     status2 (Mid(dcButton1(Index).Tag, 1, 1))
'     Call GetDiskFreeSpaceEx(Mid(dcButton1(Index).Tag, 1, 1) & ":\", BytesFreeToCalller, TotalBytes, TotalFreeBytes)
' StatusBar2.Panels(1).Text = Format$(((TotalFreeBytes * 10000) / 1024), "###,###,###,##0") & " Kb" & "  of  " & Format$(((TotalBytes * 10000) / 1024), "###,###,###,##0") & " Kb free.  (" & Format(100 - ((TotalFreeBytes * 100) / TotalBytes), "##") & "% used.)"

'dcButton1(barra.Tag).PictureOpacity = 30
'dcButton1(barra.Tag).Value = False
'dcButton1(Index).PictureOpacity = 100
'dcButton1(Index).Value = True
barra.Tag = Index
'dcButton1(barra.Tag).PictureOpacity = 30
'dcButton1(barra.Tag).Value = False

'dcButton1(Index).PictureOpacity = 100
'dcButton1(Index).Value = True

barra.Tag = Index

Dir1.Path = dcButton1(Index).Tag

Exit Sub

pepe:

Select Case Err.Number

Case 68
MsgBox "Device is not ready", vbCritical, "Error"
Case Else
End Select







End Sub




Private Sub dcButton2_Click(Index As Integer)
On Error GoTo pepe
  '   Call GetDiskFreeSpaceEx(dcButton2(Index).Caption & ":\", BytesFreeToCalller, TotalBytes, TotalFreeBytes)
 status2 (dcButton2(Index).Caption)
Call boton(dcButton2(Index).Caption & ":", dcButton2(Index).Caption & ":\")

Dir1.Path = dcButton2(Index).Caption & ":\"
Exit Sub

pepe:

Select Case Err.Number

Case 68
MsgBox "Device is not ready", vbCritical, "Error"
Case Else
End Select
End Sub







Private Sub dcButton2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Shift = 4 And KeyCode = vbKeyF1 Then
Dim A As Integer
For A = 1 To List1.ListCount
' If Len(List1.List(a)) > Len(Label2.Caption) Then Label2.Caption = List1.List(a)
If LCase(Mid(List1.List(A), 3, 1)) = LCase(Mid(Dir1.Path, 1, 1)) Then
List1.Selected(A) = True
List1.Visible = True
List1.SetFocus

End If

Next A


Else
List1.Visible = False
End If

End Sub

Private Sub dcButton3_Click()
On Error Resume Next
Dir1.Path = Mid(Dir1.Path, 1, 1) & ":\"

Call boton(Mid(Dir1.Path, 1, 1) & ":\", Mid(Dir1.Path, 1, 1) & ":\")

End Sub

Private Sub dcButton4_Click()
On Error Resume Next
Dir1.Path = Mid(Dir1.Path, 1, InStrRev(Dir1.Path, "\"))

Call boton(Mid(Dir1.Path, 1, InStrRev(Dir1.Path, "\")), Mid(Dir1.Path, 1, InStrRev(Dir1.Path, "\")))

End Sub


Private Sub Dir1_Change()
On Error Resume Next
ShowHiddenDirectories Dir1, True
Dim n As Currency
Screen.MousePointer = vbHourglass
Call LockWindowUpdate(ListView1.Hwnd)

Err.Clear
barra.Caption = Dir1.Path
txtPath.Text = Dir1.Path
'Label3.AutoSize = True
'Text1.Left = Label3.Left + Label3.Width + 100

'TabStrip1.Tabs.Add , Dir1.Path, Mid(Dir1.Path, InStrRev(Dir1.Path, "\") + 1, Len(Dir1.Path)), ImageList2.ListImages.Count
'If TabStrip1.Tabs(TabStrip1.Tabs.Count).caption = "" Then TabStrip1.Tabs(TabStrip1.Tabs.Count).caption = Dir1.Path
'If Err.Number = 35602 Then
'Dim jota As Byte
'For jota = 0 To TabStrip1.Tabs.Count

'If TabStrip1.Tabs(jota).Key = Dir1.Path Then
'TabStrip1.Tabs(jota).Selected = True



'End If

'Next


'Else
'TabStrip1.Tabs(TabStrip1.Tabs.Count).Selected = True
'ImageCombo1.ComboItems(ImageCombo1.ComboItems.Count).Selected = True


'End If


ListView1.ListItems.Clear
ListView1.SmallIcons = Nothing
ListView1.Icons = Nothing
basu.ListItems.Clear
'basu.SmallIcons = Nothing
'basu.Icons = Nothing



Me.ImageList1.ListImages.Clear
Dim sSearchPath As String, sExtensionList As String
Dim taFiles As mctFileSearchResults
Dim X As Long


If Len(Dir1.Path) <> 3 Then
basu.ListItems.Add , , "[...]"
basu.ListItems(basu.ListItems.Count).Tag = Mid(Dir1.Path, 1, InStrRev(Dir1.Path, "\"))
basu.ListItems(basu.ListItems.Count).SubItems(2) = Mid(Dir1.Path, 1, InStrRev(Dir1.Path, "\"))

End If




For X = 0 To Dir1.ListCount - 1

Set Item = basu.ListItems.Add(, , Mid(Dir1.List(X), InStrRev(Dir1.List(X), "\") + 1, Len(Dir1.List(X))))
Item.Tag = Dir1.List(X)
Item.SubItems(2) = Dir1.List(X)
'item.Bold = True
Item.SubItems(1) = ""
Item.SubItems(4) = Format(FileDateTime(Dir1.List(X)), "DD/MM/YYYY HH:MM") ' listview1.ListItems(listview1.ListItems.Count).Text
Item.SubItems(5) = "----"

Next X


' --------------------------------------------------------------------------



    sSearchPath = Dir1.Path
    sExtensionList = StatusBar1.Panels(2).Text ' "*.*" '"*.txt;*.exe"

    FileSearchA sSearchPath, sExtensionList, taFiles, False

    If taFiles.FileCount > 0 Then
        
For X = 1 To UBound(taFiles.Files)
Set Item = basu.ListItems.Add(, , Mid(taFiles.Files(X).FileName, 1, InStrRev(taFiles.Files(X).FileName, ".") - 1))
Item.SubItems(1) = taFiles.Files(X).Extension
Item.SubItems(2) = taFiles.Files(X).UNC
Item.SubItems(3) = FormatNumber(taFiles.Files(X).SIZE, 0) 'Format$(taFiles.Files(x).Size, "###,###,###,##0")
Item.SubItems(4) = Format(taFiles.Files(X).CreationDate, "DD/MM/YYYY HH:MM")
Item.SubItems(5) = IIf(taFiles.Files(X).Archive, "a", "-") & IIf(taFiles.Files(X).ReadOnly, "r", "-") & IIf(taFiles.Files(X).Hidden, "h", "-") & IIf(taFiles.Files(X).System, "s", "-")
Item.SubItems(6) = Format(Item.SubItems(3), "00000000000")
Item.SubItems(7) = Mid(Item.SubItems(4), 7, 4) & Mid(Item.SubItems(4), 4, 2) & Mid(Item.SubItems(4), 1, 2) & Mid(Item.SubItems(4), 11, 2) & Mid(Item.SubItems(4), 13, 2)

'If LCase(taFiles.Files(X).Extension) = "mp3" Then ListView2.ListItems.Add , taFiles.Files(X).UNC, taFiles.Files(X).FileName
Next

    End If
 







Call loadicons


For n = 1 To Dir1.ListCount + 1


If basu.ListItems(n).SubItems(3) = "" Then

Set Item = ListView1.ListItems.Add(, , basu.ListItems(n).Text, , basu.ListItems(n).SmallIcon)
Item.SubItems(1) = basu.ListItems(n).SubItems(1)
Item.SubItems(2) = basu.ListItems(n).SubItems(2)
Item.SubItems(3) = basu.ListItems(n).SubItems(3)
Item.SubItems(4) = basu.ListItems(n).SubItems(4)
Item.SubItems(5) = basu.ListItems(n).SubItems(5)
Item.SubItems(6) = basu.ListItems(n).SubItems(6)
Item.SubItems(7) = basu.ListItems(n).SubItems(7)
Item.Tag = basu.ListItems(n).Tag

 

End If
Next n

basu.SortOrder = ListView1.SortOrder
basu.SortKey = ListView1.SortKey
basu.Sorted = True


For n = 1 To basu.ListItems.Count


If basu.ListItems(n).SubItems(3) <> "" Then
Set Item = ListView1.ListItems.Add(, , basu.ListItems(n).Text, , basu.ListItems(n).SmallIcon)
Item.SubItems(1) = basu.ListItems(n).SubItems(1)
Item.SubItems(2) = basu.ListItems(n).SubItems(2)
Item.SubItems(3) = basu.ListItems(n).SubItems(3)
Item.SubItems(4) = basu.ListItems(n).SubItems(4)
Item.SubItems(5) = basu.ListItems(n).SubItems(5)
Item.SubItems(6) = basu.ListItems(n).SubItems(6)
Item.SubItems(7) = basu.ListItems(n).SubItems(7)

End If
Next n


basu.ListItems.Clear
basu.SmallIcons = Nothing
If Label1.Tag = "0" Then
LVDeselectAll ListView1
ListView1.SetFocus
ListView1.ListItems(CInt(Label1.Caption)).Selected = True
Else
LVDeselectAll ListView1
ListView1.SetFocus

ListView1.ListItems(1).Selected = True

End If

StatusBar1.Panels(1).Text = Format$((taFiles.FileSize / 1024), "###,###,###,##0") & " Kb in " & taFiles.FileCount & " file(s)"
 

If ListView1.ListItems(1).Text = "[...]" Then
ImageList1.ListImages.Add , , LoadPicture(App.Path & "\arriba.ico")
ListView1.ListItems(1).SmallIcon = ImageList1.ListImages.Count
End If
Call LockWindowUpdate(0)

Screen.MousePointer = vbDefault


Exit Sub
Hell:

End Sub


Private Sub Form_Initialize()
    Dim comctls As INITCOMMONCONTROLSEX_TYPE  ' identifies the control to register
    Dim retval As Long                        ' generic return value
    With comctls
        .dwSize = Len(comctls)
        .dwICC = ICC_INTERNET_CLASSES
    End With
    retval = InitCommonControlsEx(comctls)
End Sub



Private Sub Form_Load()
            FrameChildren.Add Me.dcButton1(0)





With ListView1
            With .ColumnHeaders
                .Add , , "Filename", 3500
                .Add , , "Ext", 600
                .Add , , "Path", 0
                .Add , , "Size (Kb)", 1040
                .Add , , "Date", 1700
                .Add , , "Atrib", 585
                .Add , , "_Size", 0
                .Add , , "_Date", 0
            
            .Item(3).Position = 1
            End With

End With
With basu
            With .ColumnHeaders
                .Add , , "Filename", 0
                .Add , , "Ext", 0
                .Add , , "Path", 0
                .Add , , "Size", 0
                .Add , , "Date", 0
                .Add , , "Atrib", 0
                .Add , , "_Size", 0
                .Add , , "_Date", 0

          '  .item(3).Position = 1
            End With

End With


Load_tbDrives
'Dir1.Path = Drive1.Drive
Call Dir1_Change










dcButton1(0).Caption = Mid(Dir1.Path, InStrRev(Dir1.Path, "\") + 1, Len(Dir1.Path))
dcButton1(0).Width = TextWidth(dcButton1(0).Caption) + 450

dcButton1(0).Tag = Dir1.Path & "\"
dcButton1(0).ToolTipText = dcButton1(0).Tag
  SubClassHwnd ListView1.Hwnd

End Sub



Private Sub Form_Resize()
'Frame1.Left = Form1.Width - Frame1.Width - 150
End Sub

Private Sub HScroll1_Change()


If dcButton1(dcButton1.Count - 1).Left + dcButton1(dcButton1.Count - 1).Width < HScroll1.Left And HScroll1.Value > LastScrollValue Then HScroll1.Value = HScroll1.Value - 500: Exit Sub
Dim Ctrl As Control
For Each Ctrl In FrameChildren
    Ctrl.Left = Ctrl.Left + (LastScrollValue - HScroll1.Value)    '* Screen.TwipsPerPixelY
Next
LastScrollValue = HScroll1.Value

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

Case 13
On Error GoTo pepe
 '    Call GetDiskFreeSpaceEx(Mid(List1.List(List1.ListIndex), 3, 1) & ":\", BytesFreeToCalller, TotalBytes, TotalFreeBytes)
 'StatusBar2.Panels(1).Text = Format$(((TotalFreeBytes * 10000) / 1024), "###,###,###,##0") & " Kb" & "  of  " & Format$(((TotalBytes * 10000) / 1024), "###,###,###,##0") & " Kb free.  (" & Format(100 - ((TotalFreeBytes * 100) / TotalBytes), "##") & "% used.)"
status2 (Mid(List1.List(List1.ListIndex), 3, 1))
Call boton(Mid(List1.List(List1.ListIndex), 3, 1) & ":", Mid(List1.List(List1.ListIndex), 3, 1) & ":\")

Dir1.Path = Mid(List1.List(List1.ListIndex), 3, 1) & ":\"
List1.Visible = False
ListView1.SetFocus
Case 27

List1.Visible = False
ListView1.SetFocus

End Select

Exit Sub

pepe:

Select Case Err.Number

Case 68
MsgBox "Device is not ready", vbCritical, "Error"
Case Else
End Select
List1.Visible = False
ListView1.SetFocus

End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo pepe
     
'     Call GetDiskFreeSpaceEx(Mid(List1.List(List1.ListIndex), 3, 1) & ":\", BytesFreeToCalller, TotalBytes, TotalFreeBytes)
' StatusBar2.Panels(1).Text = Format$(((TotalFreeBytes * 10000) / 1024), "###,###,###,##0") & " Kb" & "  of  " & Format$(((TotalBytes * 10000) / 1024), "###,###,###,##0") & " Kb free.  (" & Format(100 - ((TotalFreeBytes * 100) / TotalBytes), "##") & "% used.)"
status2 (Mid(List1.List(List1.ListIndex), 3, 1))
Call boton(Mid(List1.List(List1.ListIndex), 3, 1) & ":", Mid(List1.List(List1.ListIndex), 3, 1) & ":\")

Dir1.Path = Mid(List1.List(List1.ListIndex), 3, 1) & ":\"
List1.Visible = False
ListView1.SetFocus
Exit Sub

pepe:

Select Case Err.Number

Case 68
MsgBox "Device is not ready", vbCritical, "Error"
Case Else
End Select
List1.Visible = False
ListView1.SetFocus

End Sub


Private Sub ListView1_Click()
If Len(Dir1.Path) > 3 Then
    If ListView1.SelectedItem.SubItems(1) = "" Then
        If ListView1.SelectedItem.Text = "[...]" Then
            txtPath.Text = Dir1.Path
        Else
            txtPath.Text = Dir1.Path & "\" & ListView1.SelectedItem.Text
        End If
    Else
        txtPath.Text = Dir1.Path & "\" & ListView1.SelectedItem.Text & "." & ListView1.SelectedItem.SubItems(1)
    End If
Else
    If ListView1.SelectedItem.SubItems(1) = "" Then
        If ListView1.SelectedItem.Text = "[...]" Then
            txtPath.Text = Dir1.Path
        Else
            txtPath.Text = Dir1.Path & ListView1.SelectedItem.Text
        End If
    Else
        txtPath.Text = Dir1.Path & ListView1.SelectedItem.Text & "." & ListView1.SelectedItem.SubItems(1)
    End If
End If
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
Screen.MousePointer = vbHourglass
Call SortListView(ListView1, ColumnHeader)
Call Dir1_Change
Screen.MousePointer = vbDefault

End Sub

Private Sub ListView1_DblClick()
On Error Resume Next


If ListView1.SelectedItem.SubItems(3) = "" Then









If ListView1.SelectedItem.Text <> "[...]" Then
Call boton(Mid(ListView1.SelectedItem.Tag, InStrRev(ListView1.SelectedItem.Tag, "\") + 1, Len(ListView1.SelectedItem.Tag)), ListView1.SelectedItem.Tag & "\")

Label1.Caption = ListView1.SelectedItem.Index
Label1.Tag = "1"

Else
Call boton(Mid(Mid(ListView1.SelectedItem.Tag, 1, Len(ListView1.SelectedItem.Tag) - 1), InStrRev(Mid(ListView1.SelectedItem.Tag, 1, Len(ListView1.SelectedItem.Tag) - 1), "\") + 1, Len(ListView1.SelectedItem.Tag)), ListView1.SelectedItem.Tag)

Label1.Tag = "0"
End If



'Picture1.Cls
'Picture1.Picture = ImageList1.ListImages(ListView1.SelectedItem.SmallIcon).Picture
'ImageList2.ListImages.Add , , Picture1.Image
Dir1.Path = ListView1.SelectedItem.Tag

Else

  Set Item = ListView1.SelectedItem 'sets the item as the item selected in the selected list
    ShellExecute Me.Hwnd, "open", Item.SubItems(2), "", "", 3

End If
End Sub







Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
Call ListView1_DblClick

Case 116

Call Dir1_Change
'ListView1.ListItems(1).Selected = True


End Select

If Shift = 4 And KeyCode = vbKeyF1 Then
Dim A As Integer
For A = 1 To List1.ListCount
' If Len(List1.List(a)) > Len(Label2.Caption) Then Label2.Caption = List1.List(a)
If LCase(Mid(List1.List(A), 3, 1)) = LCase(Mid(Dir1.Path, 1, 1)) Then
List1.Selected(A) = True
List1.Visible = True
List1.SetFocus

End If

Next A


Else
List1.Visible = False
End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If Len(Dir1.Path) > 3 Then
    If ListView1.SelectedItem.SubItems(1) = "" Then
        If ListView1.SelectedItem.Text = "[...]" Then
            txtPath.Text = Dir1.Path
        Else
            txtPath.Text = Dir1.Path & "\" & ListView1.SelectedItem.Text
        End If
    Else
        txtPath.Text = Dir1.Path & "\" & ListView1.SelectedItem.Text & "." & ListView1.SelectedItem.SubItems(1)
    End If
Else
    If ListView1.SelectedItem.SubItems(1) = "" Then
        If ListView1.SelectedItem.Text = "[...]" Then
            txtPath.Text = Dir1.Path
        Else
            txtPath.Text = Dir1.Path & ListView1.SelectedItem.Text
        End If
    Else
        txtPath.Text = Dir1.Path & ListView1.SelectedItem.Text & "." & ListView1.SelectedItem.SubItems(1)
    End If
End If
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pt As POINTAPI
    Dim dl As Long
    Dim lCurX As Long
    Dim lCurY As Long
    Dim sFilePath As String
    On Error Resume Next
      Text1.Visible = False
  Text1.Tag = ""

Select Case Button
Case 2
'ListView1.SelectedItem.Selected = False
' ListView1.HitTest(X, Y).Selected = True
'dl = GetCursorPos(pt)
    
 '   sFilePath = ListView1.SelectedItem.SubItems(2) & Chr$(0)
    
 '   dl = DoExplorerMenu((Me.Hwnd), sFilePath, pt.X, pt.Y)
    
  '  Call Dir1_Change

Case 4
Set ListView1.MouseIcon = LoadPicture(App.Path & "\Clipboard01.ico") 'ILDrives.ListImages(5).Picture
If ListView1.MousePointer = ccCustom Then ListView1.MousePointer = ccDefault Else ListView1.MousePointer = ccCustom

Case 1


If ListView1.HitTest(X, Y).Selected = True And ListView1.SelectedItem.Text <> "[...]" Then
'Text1.Text = Mid(ListView1.SelectedItem.SubItems(2), InStrRev(ListView1.SelectedItem.SubItems(2), "\") + 1, Len(ListView1.SelectedItem.SubItems(2))) 'ListView1.SelectedItem.Text & "." & ListView1.SelectedItem.SubItems(1)
'Text1.Top = ListView1.SelectedItem.Top + ListView1.Top + 20 '+ Frame1.Top
'Text1.Left = ListView1.SelectedItem.Left + ListView1.Left + 300
'Text1.Width = (ListView1.ColumnHeaders(2).Left + ListView1.Left + ListView1.ColumnHeaders(2).Width) - Text1.Left   '+ 300
'Text1.SelStart = 0
'Text1.SelLength = Len(Text1.Text)
'Text1.Visible = True
'Text1.Tag = Text1.Text
End If
End Select
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo popo



If ListView1.MousePointer = ccCustom Then



If Y >= ListView1.Height - 1000 Then
ListView1.ListItems(ListView1.HitTest(X, Y).Index + 2).EnsureVisible
ElseIf Y <= 1000 Then
ListView1.ListItems(ListView1.HitTest(X, Y).Index - 2).EnsureVisible
End If
End If
'GoTo pepe
Exit Sub
popo:
End Sub






Private Sub loadicons()
On Error Resume Next
basu.SmallIcons = Nothing
'basu.Icons = Nothing
ImageList1.ListImages.Clear



With basu
  oldSortkey = .SortKey
  oldSortorder = .SortOrder
  .SortOrder = lvwAscending
  .SortKey = 1
  .Sorted = True
End With




Dim strExt As String
Dim strLastExt As String


  For Each Item In basu.ListItems
  
  
If Item.SubItems(1) <> "" Then
strExt = LCase(Item.SubItems(1))
      
If strExt <> strLastExt Then
hSIcon = SHGetFileInfo(Item.SubItems(2), 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
With PicFiles16
Set .Picture = LoadPicture("")
.AutoRedraw = True
R = ImageList_Draw(hSIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
.Refresh
End With

        
If strExt <> "exe" And strExt <> "ico" And strExt <> "lnk" Then
Set imgObj = ImageList1.ListImages.Add(, , PicFiles16.Image)
strLastExt = strExt
Else
Set imgObj = ImageList1.ListImages.Add(, "exe" & Item.Index, PicFiles16.Image)
strLastExt = "anything_but_exe_or_ico"
End If
      
        
        
End If

ElseIf Item.SubItems(1) = "" Then
hSIcon = SHGetFileInfo(Item.SubItems(2), 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
With PicFiles16
Set .Picture = LoadPicture("")
.AutoRedraw = True
R = ImageList_Draw(hSIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
.Refresh
End With
Set imgObj = ImageList1.ListImages.Add(, Item.SubItems(2), PicFiles16.Image)
End If
Next
  
  
    Dim strBadIconPath As String
  If Right(App.Path, 1) <> "\" Then
    strBadIconPath = App.Path & "\noicon.ico"
  Else
    strBadIconPath = App.Path & "noicon.ico"
  End If



      hSIcon = SHGetFileInfo(strBadIconPath, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
      With PicFiles16
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        R = ImageList_Draw(hSIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      
    '  basu.Icons = ImageList1
    
    
    
    
      basu.SmallIcons = ImageList1
      ListView1.SmallIcons = ImageList1
      'la.InitImlSmall
      
      
      
        strLastExt = ""

      
  Dim Image As ListImage
  For Each Item In basu.ListItems
          If Item.SubItems(1) <> "" Then

      strExt = LCase(Item.SubItems(1))
      If strExt <> strLastExt Then
      
      
        If Item.Index = 1 Then
         ' item.Icon = 1
          Item.SmallIcon = 1
          strLastExt = strExt
        Else
          Item.SmallIcon = basu.ListItems(Item.Index - 1).SmallIcon + 1
         ' item.SmallIcon = item.Icon
          strLastExt = strExt
        End If
        
        
        
      Else
      
      
      
        If strExt <> "exe" And strExt <> "ico" And strExt <> "lnk" Then
          'item.Icon = basu.ListItems(item.Index - 1).Icon
          Item.SmallIcon = basu.ListItems(Item.Index - 1).SmallIcon
        Else
          For Each Image In ImageList1.ListImages
            If Image.Key = "exe" & Item.Index Then
            '  item.Icon = Image.Index
              Item.SmallIcon = Image.Index
            End If
          Next
        End If
      End If
    
    ElseIf Item.SubItems(1) = "" Then
    
          For Each Image In ImageList1.ListImages
        If Image.Key = Item.SubItems(2) Then
          'item.Icon = Image.Index
          Item.SmallIcon = Image.Index
        End If
      Next

    End If
    
    
  Next
      
      
      
      

      

      
      
      
      
      


End Sub


Public Sub SortListView(ByRef oListView As MSComctlLib.ListView, _
                        ByRef oColumnHeader As MSComctlLib.ColumnHeader)
    
    With ListView1
    
        
       

        If oColumnHeader.Index = 4 Then
        .SortKey = 6
        ElseIf oColumnHeader.Index = 5 Then
        .SortKey = 7
        Else
        .SortKey = oColumnHeader.Index - 1

        End If
        
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        
        
       

    End With
    
End Sub

Private Sub Load_tbDrives()
    Static i As Integer
    Dim sDrive As String, strSave As String
Dim Ret As String
Dim keer As Integer




    strSave = String(255, Chr$(0))
    Ret = GetLogicalDriveStrings(255, strSave)
    For keer = 1 To 100
        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
            sDrive = Left$(LCase(strSave), InStr(1, strSave, Chr$(0)) - 3)
                
                strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
                
        If Not i = 0 Then
    Load dcButton2(i)
    dcButton2(i).Caption = sDrive
    dcButton2(i).Left = dcButton2(i - 1).Left + dcButton2(i - 1).Width + 70
    dcButton2(i).Visible = True
    
    End If
    
     dcButton2(i).Caption = sDrive
     
      
     
  Select Case GetDriveType(sDrive & ":\")
  Case 2
  Set dcButton2(i).Picture = LoadPicture(App.Path & "\icons\2.ico")
  dcButton2(i).ToolTipText = "Removable"
  List1.AddItem "[-" & sDrive & "-]" & vbTab & "3½"""
  Case 3
  Set dcButton2(i).Picture = LoadPicture(App.Path & "\icons\3.ico")
  dcButton2(i).ToolTipText = "Fixed"
    List1.AddItem "[-" & sDrive & "-]" & vbTab & VolumeName(sDrive)

  Case 4
  Set dcButton2(i).Picture = LoadPicture(App.Path & "\icons\4.ico")
  dcButton2(i).ToolTipText = "Remote"
      List1.AddItem "[-" & sDrive & "-]" & vbTab & "Remote"

  Case 5
  Set dcButton2(i).Picture = LoadPicture(App.Path & "\icons\5.ico")
  dcButton2(i).ToolTipText = "CD-ROM"
    List1.AddItem "[-" & sDrive & "-]" & vbTab & "CD-ROM"
  
  Case 6
  Set dcButton2(i).Picture = LoadPicture(App.Path & "\icons\6.ico")
  dcButton2(i).ToolTipText = "RAM Disk"
      List1.AddItem "[-" & sDrive & "-]" & vbTab & "RAM Disk"

  Case Else
  Set dcButton2(i).Picture = LoadPicture(App.Path & "\icons\7.ico")
  dcButton2(i).ToolTipText = "Unknown"
    List1.AddItem "[-" & sDrive & "-]" & vbTab & "Unknown"
  End Select
    
      

    
 If LCase(Mid(sDrive, 1, 1)) = LCase(Mid(Dir1.Path, 1, 1)) Then

     'Call GetDiskFreeSpaceEx(Mid(sDrive, 1, 1) & ":\", BytesFreeToCalller, TotalBytes, TotalFreeBytes)

 'StatusBar2.Panels(1).Text = Format$(((TotalFreeBytes * 10000) / 1024), "###,###,###,##0") & " Kb" & "  of  " & Format$(((TotalBytes * 10000) / 1024), "###,###,###,##0") & " Kb free.  (" & Format(100 - ((TotalFreeBytes * 100) / TotalBytes), "##") & "% used.)"

status2 (Mid(sDrive, 1, 1))
End If
    
    
    i = i + 1
 
    
    
          
    Next keer

    List1.Height = List1.Height * List1.ListCount

    dcButton3.Left = dcButton2(i - 1).Left + dcButton2(i - 1).Width + 70
    dcButton4.Left = dcButton3.Left + dcButton3.Width + 70
StatusBar2.Left = dcButton4.Left + dcButton4.Width + 70
StatusBar2.Width = ListView1.Width - StatusBar2.Left






















End Sub

Private Sub boton(Caption As String, pt As String, Optional cual As String)
    
    On Error Resume Next
    Dim C As Integer
    Dim suma As Long
    Static i As Integer
    
    

For C = dcButton1.LBound To dcButton1.UBound
If LCase(dcButton1(C).Tag) = LCase(pt) Then
'dcButton1(barra.Tag).Picture = 30
'dcButton1(barra.Tag).Value = False

'dcButton1(C).PictureOpacity = 100
'dcButton1(C).Value = True
barra.Tag = C
Exit Sub
End If
Next C

'dcButton1(barra.Tag).PictureOpacity = 30
'dcButton1(barra.Tag).Value = False
i = i + 1
Load dcButton1(i)
    
    
    dcButton1(i).Left = dcButton1(i - 1).Left + dcButton1(i - 1).Width + 60
    dcButton1(i).Top = dcButton1(i - 1).Top
    dcButton1(i).Caption = Caption
    dcButton1(i).Tag = pt
  '      dcButton1(i).Width = PixelsToTwips_width(GetSystemFontTextWidth(dcButton1(i).Caption)) + 450
 ' Me.Caption = TextWidth(dcButton1(i).Caption)

dcButton1(i).Width = TextWidth(dcButton1(i).Caption) + 450
    dcButton1(i).Visible = True

    If dcButton1(i).Left + dcButton1(i).Width >= Frame2.Left + Frame2.Width Then
    Me.HScroll1.Visible = True
    Else
    Me.HScroll1.Visible = False
    
End If
    
 ' Set dcbutton1(i).PictureNormal = ImageList1.ListImages(ListView1.SelectedItem.Index).Picture
    
'dcButton1(i).PictureOpacity = 100
'dcButton1(i).Value = True
    barra.Tag = i
    
    
'    dcButton1(i).ButtonShape = ebsCutSides
    Dim Ctrl As Control

            FrameChildren.Add Me.dcButton1(i)
dcButton1(i).ToolTipText = dcButton1(i).Tag
    Me.HScroll1.Max = Me.HScroll1.Max + dcButton1(i).Left

Handler1:



End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Visible = True Then Text1.SetFocus: Text1.SelLength = 0: Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
Select Case Panel.Index

Case 2
StatusBar1.Tag = ""
Form2.Text1.Text = StatusBar1.Panels(2).Text
Form2.Top = StatusBar1.Top  ' - Form2.Height
Form2.Left = Frame1.Left + ListView1.Left + ListView1.Width - Form2.Width
Form2.Text1.SelStart = InStrRev(Form2.Text1.Text, ".")
Form2.Text1.SelLength = Len(Form2.Text1.Text)
Form2.Show vbModal, Me
If StatusBar1.Tag = "Ok" Then Call Dir1_Change
End Select
End Sub

Public Function VolumeName(Optional Drive As String)
Dim sBuffer As String
sBuffer = Space$(255) 'fix bad parameter values
If Len(Drive) = 1 Then Drive = Drive & ":\"
If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
If GetVolumeInformation(Drive, sBuffer, Len(sBuffer), 0, 0, 0, Space$(255), 255) = 0 Then
Else
VolumeName = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
End If
End Function

Public Function status2(uni As String)
Dim R As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
Call GetDiskFreeSpaceEx(uni & ":\", BytesFreeToCalller, TotalBytes, TotalFreeBytes)
Call LockWindowUpdate(StatusBar2.Hwnd)

StatusBar2.Panels(1).Text = Format$(((TotalFreeBytes * 10000) / 1024), "###,###,###,##0") & " Kb" & "  of  " & Format$(((TotalBytes * 10000) / 1024), "###,###,###,##0") & " Kb free.  (" & Format(100 - ((TotalFreeBytes * 100) / TotalBytes), "##") & "% used.)"
Call LockWindowUpdate(0)

End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo pepe


If Text1.Text = "" Then GoTo pepe

Select Case KeyCode
Case 13
If InStr(Text1.Text, ".") <> 0 Then

If ListView1.SelectedItem.SubItems(3) = "" Then
ListView1.SelectedItem.Text = Text1.Text
'ListView1.SelectedItem.SubItems(1) = Mid(Text1.Text, InStrRev(Text1.Text, ".") + 1, Len(Text1.Text))
Else
ListView1.SelectedItem.Text = Mid(Text1.Text, 1, InStrRev(Text1.Text, ".") - 1)
ListView1.SelectedItem.SubItems(1) = Mid(Text1.Text, InStrRev(Text1.Text, ".") + 1, Len(Text1.Text))


End If


Else
ListView1.SelectedItem.Text = Text1.Text
ListView1.SelectedItem.SubItems(1) = ""

End If

If ListView1.SelectedItem.SubItems(3) <> "" Then
FileRename ListView1.SelectedItem.SubItems(2), Replace(ListView1.SelectedItem.SubItems(2), Text1.Tag, Text1.Text), True
ListView1.SelectedItem.SubItems(2) = Replace(ListView1.SelectedItem.SubItems(2), Text1.Tag, Text1.Text)
Text1.Visible = False
'Call Dir1_Change
loadparticular ListView1.SelectedItem.SubItems(2)

ListView1.SelectedItem.SmallIcon = Me.ImageList1.ListImages.Count

ListView1.SelectedItem.Selected = False
Else
FileRename ListView1.SelectedItem.Tag & "\", Replace(ListView1.SelectedItem.SubItems(2) & "\", Text1.Tag, Text1.Text), True
'botonrename Text1.Text, ListView1.SelectedItem.SubItems(2), Replace(ListView1.SelectedItem.SubItems(2), Text1.Tag, Text1.Text)
ListView1.SelectedItem.SubItems(2) = Replace(ListView1.SelectedItem.SubItems(2), Text1.Tag, Text1.Text)
ListView1.SelectedItem.Tag = ListView1.SelectedItem.SubItems(2)
loadparticular ListView1.SelectedItem.Tag
ListView1.SelectedItem.SmallIcon = Me.ImageList1.ListImages.Count
Text1.Visible = False
'Call Dir1_Change
ListView1.SelectedItem.Selected = False


End If

Case 27
Text1.Visible = False
End Select
Exit Sub
pepe:
Text1.Visible = False
Text1.Tag = ""

End Sub






Private Sub loadparticular(arch As String)
On Error Resume Next
PicFiles16.Cls
hSIcon = SHGetFileInfo(arch, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
With PicFiles16
Set .Picture = LoadPicture("")
.AutoRedraw = True
R = ImageList_Draw(hSIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
.Refresh
End With
     
Set imgObj = ImageList1.ListImages.Add(, , PicFiles16.Image)
      
      
      
      

      

      
      
      
      
      


End Sub
Private Sub botonrename(original As String, pt As String, pt2 As String)
    
    On Error Resume Next
    Dim C As Integer
    
   

For C = dcButton1.LBound To dcButton1.UBound
Debug.Print dcButton1(C).Tag & vbTab & pt
If LCase(dcButton1(C).Tag) = LCase(pt) Then
dcButton1(C).Caption = original
dcButton1(C).Tag = pt2
Exit Sub
End If
Next C




End Sub

Private Sub txtPath_Click()
    With txtPath
    .SelStart = 0
    .SelLength = Len(txtPath)
    End With
End Sub
