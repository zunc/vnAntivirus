Attribute VB_Name = "winapi"
Public Declare Function DoExplorerMenu Lib "cfexpmnu.dll" (ByVal Hwnd As Long, ByVal sFilePath As String, ByVal X As Long, ByVal Y As Long) As Boolean
Public Type POINTAPI
        X As Long
        Y As Long
End Type
        
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long



Function FileRename(sOriginalName As String, sNewName As String, Optional bOverWrite As Boolean) As Boolean

    On Error GoTo ErrFailed
    If Len(Dir$(sOriginalName)) > 0 And Len(sOriginalName) > 0 Then     'Check File Exitsts
        If bOverWrite = True And Len(Dir(sNewName)) > 0 Then
            'Delete file with same name as new file
            VBA.Kill sNewName
        End If
        Name sOriginalName As sNewName
        FileRename = True
    End If

    Exit Function
    
ErrFailed:
    'Failed to Rename File
    FileRename = False
    On Error GoTo 0
End Function

