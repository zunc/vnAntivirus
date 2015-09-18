Attribute VB_Name = "Module2"


Option Explicit
'-----------------------------ListView API----------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'-----------------------------ListView messages-----------------
Private Const LVM_FIRST = &H1000
Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Private Const LVNI_SELECTED = &H2
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)





Function LVDeselectAll(oListView As ListView) As Boolean
    Dim sThisItem As Long, lLvHwnd As Long, lSelectedItems As Long, lItemIndex As Long
    
    On Error GoTo ErrFailed
    
    With oListView
        lLvHwnd = .Hwnd
        .Visible = False             'For speed. Need to remove the line in VBA
        lSelectedItems = SendMessage(lLvHwnd, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
        lItemIndex = -1
        For sThisItem = 1 To lSelectedItems
            lItemIndex = SendMessage(lLvHwnd, LVM_GETNEXTITEM, lItemIndex, ByVal LVNI_SELECTED)
            .ListItems(lItemIndex + 1).Selected = False
        Next
        .Visible = True              'For speed. Need to remove the line in VBA
    End With
    Exit Function
    
ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
    LVDeselectAll = True
End Function



Function LVSortColumns(LVSort As ListView, LVColumnHeader As ColumnHeader) As Long
    
    On Error GoTo ErrFailed
    With LVSort
        'HACK: Protects against an occassional 'division by zero' general protection fault when sorting an empty listview
        If .ListItems.Count > 0 Then
            .Visible = False        'For speed. Need to remove the line in VBA
            .SortKey = LVColumnHeader.Index - 1
            .SortOrder = 1 - LVSort.SortOrder
            .Sorted = True
            .Visible = True         'For speed. Need to remove the line in VBA
        End If
    End With
    
    Exit Function
    
ErrFailed:
    Debug.Assert False
    LVSortColumns = Err.Number
    On Error Resume Next
End Function



Function LvItemExists(oLv As ListView, sKeyName As String) As Boolean
    Dim bTest As Boolean
    On Error GoTo ErrFailed
    bTest = oLv.ListItems(sKeyName).Bold
    LvItemExists = True
    Exit Function

ErrFailed:
    LvItemExists = False
    On Error GoTo 0
End Function

