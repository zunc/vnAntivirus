Attribute VB_Name = "basLanguage"
'vnAntiVirus 0.5

'Author : Dung Nguyen Le
'Email : dungcoivb@gmail.com
'My forum : www.vietvirus.info
'This is a software open source

Public Lang As String
Public PathApp As String
Public bLang As Boolean
Public Sub SetLang(FileLang As String, FormName As Form)
'Hi hi, dau co le la mot trong nhung nhung thanh phan hay nhat cua cai Project nay
    Dim cc As Control
    Dim na As String
    Dim naf As String
    
    For Each cc In FormName.Controls
        na = cc.Name
        naf = FormName.Name
            If (Left(na, 3) = "cmd") Or ((Left(na, 3) = "opt") Or ((Left(na, 3) = "chk")) Or ((Left(na, 3) = "frm")) Or ((Left(na, 3) = "mnu")) Or ((Left(na, 3) = "lbl"))) Then
                'MsgBox ReadINI(FileLang, FormName.Name, na)
                cc.Caption = ReadINI(FileLang, naf, na)
            ElseIf Left(na, 2) = "LV" Then
                Dim i As Byte
                For i = 1 To cc.ColumnHeaders.Count
                    cc.ColumnHeaders.Item(i).Text = ReadINI(FileLang, naf, na & "(" & i & ")")
                Next
            End If
    Next
End Sub
Public Function GetStr(MesKey As String) As String
    If ioptVie = True Then
        GetStr = ReadINI(PathApp & "\Language\VietNam.lng", "Message", MesKey)
    Else
        GetStr = ReadINI(PathApp & "\Language\EngLish.lng", "Message", MesKey)
    End If
End Function
Public Function GetStrOther(MesKey As String) As String
    If ioptVie = True Then
        GetStrOther = ReadINI(PathApp & "\Language\VietNam.lng", "Other", MesKey)
    Else
        GetStrOther = ReadINI(PathApp & "\Language\EngLish.lng", "Other", MesKey)
    End If
End Function
Public Sub Language(FormName As Form)
If bLang = True Then
    If ioptVie = True Then
        SetLang PathApp & "\Language\VietNam.lng", FormName
    Else
        SetLang PathApp & "\Language\EngLish.lng", FormName
    End If
End If
End Sub
