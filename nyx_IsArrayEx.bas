Attribute VB_Name = "nyx_IsArrayEx"
'***********************************************************
' �@�\   : �������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' ����   : varArray  �z��
' �߂�l : ���茋�ʁi1:�z��/0:��̔z��/-1:�z�񂶂�Ȃ��j
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
            Debug.Print "strArray�͔z��ł��B"
        Case 0
            Debug.Print "strArray�͋�̔z��ł��B"
        Case -1
            Debug.Print "strArray�͔z��ł͂���܂���B"
    End Select
End Sub
