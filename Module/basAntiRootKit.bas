Attribute VB_Name = "basAntiRootKit"
'Module nay, toi chan thanh cam on PhamTienSinh (Pham Trung Hai)

'Ok, mot phuong phap nhan dang RootKit mot cach that don gian va tuyet voi
'Cam on PTS rat nhieu voi ky thuat nay

'Chu y : Phuong phap nay chi nhan ra nhung loai RootKit van con de day vet la cac thong so Handl, neu ko co, chuong trinh se ko the nhan ra
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    Dim ID As Long
    Dim CoChua As Boolean
    Dim i As Integer
    CoChua = False
    GetWindowThreadProcessId hWnd, ID
    
    With frmMnu.lstPro
        For i = 0 To .ListCount - 1
            If ID = Val(.List(i)) Then CoChua = True
        Next
        If CoChua = False Then .AddItem ID
    End With
    
If CoChua = False Then
    If CheckID(ID) <> ID Then
        Dim tmp As String
        tmp = ProcessPathByPID(ID)
    
        CoChua = False
        With frmPro

            Set lsv = .LV.ListItems.Add()
            lsv.Text = GetFileName(tmp)
            lsv.SubItems(1) = tmp
            lsv.SubItems(2) = ID
            lsv.ForeColor = vbRed
            
        End With
        
    End If
End If

    EnumWindowsProc = True
End Function
