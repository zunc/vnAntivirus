Attribute VB_Name = "basMessage"
Public dlv As Integer
Public ChiSo As Byte
Public Sub ThongBao(strCaption As String, NoiDung As String, strPath As String, So As Byte, Optional ID As Long)
    Dim frmTmp As New frmMas
    Dim i
    With frmTmp
    ChiSo = So
    If So = 0 Then
        .pic(0).Visible = True
        .txtMes.Text = NoiDung
        .cmdResume.Visible = False
    Else
        .txtMes.Visible = False
        For i = 0 To So
            .pic(i).Visible = True
            .lblMas(i).Caption = Split(NoiDung, "|")(i)
        Next
        .cmdResume.Visible = False
        If So = 2 Then .lblMas(3).Caption = "Warning : It can is worm. Blocked this process": .cmdResume.Visible = True
    End If
        .lblCaption.Caption = strCaption
        .txtPath.Text = strPath
        .lblID.Caption = ID
        .Show
    End With
End Sub
