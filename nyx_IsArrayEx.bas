Attribute VB_Name = "nyx_IsArrayEx"
'***********************************************************
' 機能   : 引数が配列か判定し、配列の場合は空かどうかも判定する
' 引数   : varArray  配列
' 戻り値 : 判定結果（1:配列/0:空の配列/-1:配列じゃない）
'***********************************************************
'http://www.openreference.org/articles/view/583

Public Function IsArrayEx(varArray As Variant) As Long
On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_:
    If err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function

Public Sub sample()
    Dim strArray() As String
    
    Select Case IsArrayEx(strArray)
        Case 1
            Debug.Print "strArrayは配列です。"
        Case 0
            Debug.Print "strArrayは空の配列です。"
        Case -1
            Debug.Print "strArrayは配列ではありません。"
    End Select
End Sub
