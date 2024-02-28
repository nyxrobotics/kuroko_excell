Attribute VB_Name = "nyx_export"
'--------------------------------------------------------------
'ファイル名から拡張子を除く名前を取り出す関数
Function GetFNameFromFStr(sFileName As String) As String
    Dim sFileStr As String
    Dim lFindPoint As Long
    Dim lStrLen As Long
    
    '文字列の右端から"."を検索し、左端からの位置を取得する
    lFindPoint = InStrRev(sFileName, ".")
    
    '拡張子を除いたファイル名の取得
    sFileStr = Left(sFileName, lFindPoint - 1)

    GetFNameFromFStr = sFileStr
End Function
'--------------------------------------------------------------------

Sub MotionExport_Move()
    '----------------------------------
    'motion_exportフォルダがなければ作成する
    
    Dim MotionExportDirectry As String
    MotionExportDirectry = ActiveWorkbook.Path & "\motion_export"
    If Dir(MotionExportDirectry, vbDirectory) = "" Then
        MkDir MotionExportDirectry
    End If
    
    '----------------------------------
    '出力ファイル名を決定
    Dim Filename As String
    Filename = GetFNameFromFStr(ActiveWorkbook.Name) & "_" & ActiveSheet.Name
    Dim OutputFile As String
    OutputFile = ActiveWorkbook.Path & "\motion_export\" & Filename & ".c"
    
    '-----------------------------------
    
    Dim i As Long, LngLoop As Long
    Dim IntFlNo As Integer
    
    'Worksheets("Sheet1").Activate
    LngLoop = Range("a65536").End(xlUp).Row
    'LngLoop = Range("a100").End(xlUp).Row
    
    IntFlNo = FreeFile
    Open OutputFile For Output As #IntFlNo
    
    Dim Shtname As String
    Shtname = ActiveSheet.Name
    Dim StartFrame As Integer, EndFrame As Integer, LoopStart As Integer, LoopEnd As Integer
    StartFrame = 0
    EndFrame = 0
    LoopStart = 0
    LoopEnd = 0
    '-------------
    '開始位置を調べる
    For k = 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 1 Then
            StartFrame = k - 8
            EndFrame = StartFrame
            k = 68
        End If
    Next
   
    '-------------
    '終了位置を調べる
    For k = StartFrame + 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 2 Then
            EndFrame = k - 8
            k = 68
        End If
    Next
    
    '-------------
    'ループ位置を調べる
    For k = (StartFrame + 8) To (EndFrame + 8)
        If Not Sheets(Shtname).Cells(1, k) = "" Then
            LoopStart = Sheets(Shtname).Cells(1, k).Value
            LoopEnd = k - 8
            k = EndFrame + 8
        End If
    Next
    
    '-------------
    
    
    
    Dim TotalFrame As Integer
    '----------------------------------
    'スターティングモーション
    If LoopStart <> 0 Then
    TotalFrame = LoopStart - StartFrame
    Else
        If EndFrame > StartFrame Then
        TotalFrame = EndFrame - StartFrame + 1
        Else
        TotalFrame = 0
        End If
    End If
    
    
    Dim A(1000) As String
    
    A(1) = String(4 - LenB(StrConv(TotalFrame, vbFromUnicode)), " ") & TotalFrame
    'Format(TotalFrame, "0000")
    A(1) = A(1) & ", " & String(4 - LenB(StrConv(servosend, vbFromUnicode)), " ") & servosend
    'a(1) & ", " & Format(servosend, "0000")
        For i = 8 To (servosend + 7)
        A(1) = A(1) & ", " & String(4 - LenB(StrConv(Cells(i, 2).Value, vbFromUnicode)), " ") & Cells(i, 2).Value
        'Format(Cells(i, 2).Value, "0000")
        Next i
        
    A(2) = String(4 - LenB(StrConv(Cells(5, 8).Value, vbFromUnicode)), " ") & Cells(5, 8).Value
    'Format(Cells(5, 8).Value, "0000")
    A(2) = A(2) & ", " & String(4 - LenB(StrConv(Cells(6, 8).Value, vbFromUnicode)), " ") & Cells(6, 8).Value
    '& Format(Cells(6, 8).Value, "0000")
        For i = 1 To servosend
        If (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value) < 0 Then
        A(2) = A(2) & "," & String(4 - LenB(StrConv(-10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ") & 10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        'Format(10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), "0000")
        Else
        A(2) = A(2) & ", " & String(4 - LenB(StrConv(10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ") & 10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        '& Format(10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), "0000")
        End If
        Next i
        
        
    If TotalFrame > 0 Then
        For i = 3 To TotalFrame + 2
            A(i) = String(4 - LenB(StrConv(Cells(5, 7 + StartFrame + i - 2).Value, vbFromUnicode)), " ") & Cells(5, 7 + StartFrame + i - 2).Value
            'Format(Cells(5, 7 + StartFrame + i - 2).Value, "0000")
            A(i) = A(i) & ", " & String(4 - LenB(StrConv(Cells(6, 7 + StartFrame + i - 2).Value, vbFromUnicode)), " ") & Cells(6, 7 + StartFrame + i - 2).Value
            ' Format(Cells(6, 7 + StartFrame + i - 2).Value, "0000")
            For j = 1 To servosend
              If Cells(7 + j, i + StartFrame + 5).Value < 0 Then
                A(i) = A(i) & "," & String(4 - LenB(StrConv(-10 * CInt(Cells(7 + j, i + StartFrame + 5).Value), vbFromUnicode)), " ") & 10 * CInt(Cells(7 + j, i + StartFrame + 5).Value)
                '& Format(10 * Cells(7 + j, i + StartFrame + 5).Value, "0000")
              Else
                A(i) = A(i) & ", " & String(4 - LenB(StrConv(10 * CInt(Cells(7 + j, i + StartFrame + 5).Value), vbFromUnicode)), " ") & 10 * CInt(Cells(7 + j, i + StartFrame + 5).Value)
                ' Format(10 * Cells(7 + j, i + StartFrame + 5).Value, "0000")
              End If
            Next j
        Next i
    End If
    
     'Print #IntFlNo, StartFrame & EndFrame & LoopStart & LoopEnd & ""
    
    If TotalFrame > 0 Then 'こっちにすると何もないと出力しなくなる
    Print #IntFlNo, "int " & ActiveSheet.Name & "_Motion_Start[" & TotalFrame + 2 & "][" & servosend + 2 & "]={"
    Print #IntFlNo, "{" & A(1) & "},//(一個目モーションの総フレーム数、二個目サーボの総数、三個目以降が指定したID)"
    Print #IntFlNo, "{" & A(2) & "},//初期姿勢(一個目移動時間、二個目待機時間、三個目以降角度)"
        If TotalFrame > 1 Then
        Print #IntFlNo, "{" & A(3) & "},//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
            If TotalFrame > 2 Then
            For i = 4 To (TotalFrame + 1)
            Print #IntFlNo, "{" & A(i) & "},"
            Next i
            End If
        Print #IntFlNo, "{" & A(TotalFrame + 2) & "}"
        Else
        Print #IntFlNo, "{" & A(3) & "}//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
        End If
    Print #IntFlNo, "};"
    Print #IntFlNo, ""
    End If
    

    '----------------------------------------
    'ループモーション
    If LoopStart = 0 Then
    TotalFrame = 0
    Else
        If LoopEnd = 0 Then
        TotalFrame = 0
        Else
            If LoopEnd - LoopStart < 0 Then
            TotalFrame = 0
            Else
            TotalFrame = LoopEnd - LoopStart + 1
            End If
        End If
    End If
    

    
    A(1) = String(4 - LenB(StrConv(TotalFrame, vbFromUnicode)), " ") & TotalFrame
    A(1) = A(1) & ", " & String(4 - LenB(StrConv(servosend, vbFromUnicode)), " ") & servosend
        For i = 8 To (servosend + 7)
        A(1) = A(1) & ", " & String(4 - LenB(StrConv(Cells(i, 2).Value, vbFromUnicode)), " ") & Cells(i, 2).Value
        Next i
        
    A(2) = String(4 - LenB(StrConv(Cells(5, 8).Value, vbFromUnicode)), " ") & Cells(5, 8).Value
    A(2) = A(2) & ", " & String(4 - LenB(StrConv(Cells(6, 8).Value, vbFromUnicode)), " ") & Cells(6, 8).Value
        For i = 1 To servosend
        If (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value) < 0 Then
        A(2) = A(2) & "," & String(4 - LenB(StrConv(-10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ") & 10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        Else
        A(2) = A(2) & ", " & String(4 - LenB(StrConv(10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ") & 10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        End If
        Next i
        
        
    If TotalFrame > 0 Then
        For i = 3 To TotalFrame + 2
            A(i) = String(4 - LenB(StrConv(Cells(5, 7 + LoopStart + i - 2).Value, vbFromUnicode)), " ") & Cells(5, 7 + LoopStart + i - 2).Value
            'Format(Cells(5, 7 + LoopStart + i - 2).Value, "0000")
            A(i) = A(i) & ", " & String(4 - LenB(StrConv(Cells(6, 7 + LoopStart + i - 2).Value, vbFromUnicode)), " ") & Cells(6, 7 + LoopStart + i - 2).Value
            'Format(Cells(6, 7 + LoopStart + i - 2).Value, "0000")
            For j = 1 To servosend
              If Cells(7 + j, i + LoopStart + 5).Value < 0 Then
                A(i) = A(i) & "," & String(4 - LenB(StrConv(-10 * CInt(Cells(7 + j, i + LoopStart + 5).Value), vbFromUnicode)), " ") & 10 * CInt(Cells(7 + j, i + LoopStart + 5).Value)
                'Format(10 * Cells(7 + j, i + LoopStart + 5).Value, "0000")
              Else
                A(i) = A(i) & ", " & String(4 - LenB(StrConv(10 * CInt(Cells(7 + j, i + LoopStart + 5).Value), vbFromUnicode)), " ") & 10 * CInt(Cells(7 + j, i + LoopStart + 5).Value)
                'Format(10 * Cells(7 + j, i + LoopStart + 5).Value, "0000")
              End If
            Next j
        Next i
    End If
    
    
    If TotalFrame > 0 Then 'こっちにすると何もないと出力しなくなる
    Print #IntFlNo, "int " & ActiveSheet.Name & "_Motion_Loop[" & TotalFrame + 2 & "][" & servosend + 2 & "]={"
    Print #IntFlNo, "{" & A(1) & "},//(一個目モーションの総フレーム数、二個目サーボの総数、三個目以降が指定したID)"
    Print #IntFlNo, "{" & A(2) & "},//初期姿勢(一個目移動時間、二個目待機時間、三個目以降角度)"
        If TotalFrame > 1 Then
        Print #IntFlNo, "{" & A(3) & "},//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
            If TotalFrame > 2 Then
            For i = 4 To (TotalFrame + 1)
            Print #IntFlNo, "{" & A(i) & "},"
            Next i
            End If
        Print #IntFlNo, "{" & A(TotalFrame + 2) & "}"
        Else
        Print #IntFlNo, "{" & A(3) & "}//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
        End If
    Print #IntFlNo, "};"
    Print #IntFlNo, ""
    End If
    '----------------------------------------
    'エンドモーション
    If LoopEnd = 0 Then
    TotalFrame = 0
    Else
        If LoopEnd = 0 Then
        TotalFrame = 0
        Else
            If LoopEnd - LoopStart < 0 Then
            TotalFrame = 0
            Else
            TotalFrame = EndFrame - LoopEnd
            End If
        End If
    End If

    A(1) = String(4 - LenB(StrConv(TotalFrame, vbFromUnicode)), " ") & TotalFrame
    A(1) = A(1) & ", " & String(4 - LenB(StrConv(servosend, vbFromUnicode)), " ") & servosend
        For i = 8 To (servosend + 7)
        A(1) = A(1) & ", " & String(4 - LenB(StrConv(Cells(i, 2).Value, vbFromUnicode)), " ") & Cells(i, 2).Value
        Next i
        
    A(2) = String(4 - LenB(StrConv(Cells(5, 8).Value, vbFromUnicode)), " ") & Cells(5, 8).Value
    A(2) = A(2) & ", " & String(4 - LenB(StrConv(Cells(6, 8).Value, vbFromUnicode)), " ") & Cells(6, 8).Value
        For i = 1 To servosend
        If (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value) < 0 Then
        A(2) = A(2) & "," & String(4 - LenB(StrConv(-10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ") & 10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        Else
        A(2) = A(2) & ", " & String(4 - LenB(StrConv(10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ") & 10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        End If
        Next i
        
        
    If TotalFrame > 0 Then
        For i = 3 To TotalFrame + 2
            A(i) = String(4 - LenB(StrConv(Cells(5, 7 + LoopEnd + i - 1).Value, vbFromUnicode)), " ") & Cells(5, 7 + LoopEnd + i - 1).Value
            'Format(Cells(5, 7 + LoopEnd + i - 1).Value, "0000")
            A(i) = A(i) & ", " & String(4 - LenB(StrConv(Cells(6, 7 + LoopEnd + i - 1).Value, vbFromUnicode)), " ") & Cells(6, 7 + LoopEnd + i - 1).Value
            'Format(Cells(6, 7 + LoopEnd + i - 1).Value, "0000")
            For j = 1 To servosend
              If Cells(7 + j, i + LoopEnd + 6).Value < 0 Then
                A(i) = A(i) & "," & String(4 - LenB(StrConv(-10 * CInt(Cells(7 + j, i + LoopEnd + 6).Value), vbFromUnicode)), " ") & 10 * CInt(Cells(7 + j, i + LoopEnd + 6).Value)
                'Format(10 * Cells(7 + j, i + LoopEnd + 6).Value, "0000")
              Else
                A(i) = A(i) & ", " & String(4 - LenB(StrConv(10 * CInt(Cells(7 + j, i + LoopEnd + 6).Value), vbFromUnicode)), " ") & 10 * CInt(Cells(7 + j, i + LoopEnd + 6).Value)
                'Format(10 * Cells(7 + j, i + LoopEnd + 6).Value, "0000")
              End If
            Next j
        Next i
    End If
    
    
    If TotalFrame > 0 Then 'こっちにすると何もないと出力しなくなる
    Print #IntFlNo, "int " & ActiveSheet.Name & "_Motion_End[" & TotalFrame + 2 & "][" & servosend + 2 & "]={"
    Print #IntFlNo, "{" & A(1) & "},//(一個目モーションの総フレーム数、二個目サーボの総数、三個目以降が指定したID)"
    Print #IntFlNo, "{" & A(2) & "},//初期姿勢(一個目移動時間、二個目待機時間、三個目以降角度)"
        If TotalFrame > 1 Then
        Print #IntFlNo, "{" & A(3) & "},//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
            If TotalFrame > 2 Then
            For i = 4 To (TotalFrame + 1)
            Print #IntFlNo, "{" & A(i) & "},"
            Next i
            End If
        Print #IntFlNo, "{" & A(TotalFrame + 2) & "}"
        Else
        Print #IntFlNo, "{" & A(3) & "}//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
        End If
    Print #IntFlNo, "};"
    Print #IntFlNo, ""
    Else
    End If
    

    '----------------------------------------
    
      
    Close #IntFlNo
End Sub



Sub MotionExport_Atk()
    '----------------------------------
    'motion_exportフォルダがなければ作成する
    Dim MotionExportDirectry As String
    MotionExportDirectry = ActiveWorkbook.Path & "\motion_export"
    If Dir(MotionExportDirectry, vbDirectory) = "" Then
        MkDir MotionExportDirectry
    End If
    '----------------------------------
    '出力ファイル名を決定
    Dim Filename As String
    Filename = GetFNameFromFStr(ActiveWorkbook.Name) & "_" & ActiveSheet.Name
    Dim OutputFile As String
    OutputFile = ActiveWorkbook.Path & "\motion_export\" & Filename & ".c"
    '-----------------------------------
    Dim i As Long, LngLoop As Long
    Dim IntFlNo As Integer
    LngLoop = Range("a65536").End(xlUp).Row
    IntFlNo = FreeFile
    Open OutputFile For Output As #IntFlNo
    Dim Shtname As String
    Shtname = ActiveSheet.Name
    Dim StartFrame As Integer, EndFrame As Integer, LoopStart As Integer, LoopEnd As Integer
    StartFrame = 0
    EndFrame = 0
    LoopStart = 0
    LoopEnd = 0
    '-------------
    '開始位置を調べる
    For k = 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 1 Then
            StartFrame = k - 8
            EndFrame = StartFrame
            k = 68
        End If
    Next
    '-------------
    '終了位置を調べる
    For k = StartFrame + 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 2 Then
            EndFrame = k - 8
            k = 68
        End If
    Next
    '-------------
    '分割位置を調べる
    Dim DotPoint(100) As Integer
    Dim Dots As Integer
    Dim l As Integer, m As Integer
    m = (StartFrame + 8)
    For l = 1 To 100
        DotPoint(l) = 0
    Next
    For l = 1 To 100
        For k = m To (EndFrame + 8)
            If Sheets(Shtname).Cells(1, k) = "d" Then
            m = k + 1
            DotPoint(l) = k - 8
            k = EndFrame + 8
            End If
        Next
    Next
    For l = 1 To 100
        If DotPoint(l) = 0 Then
        'Dots = l - 1
        Dots = l
        DotPoint(l) = EndFrame
        l = 100
        Else
        'Print #IntFlNo, DotPoint(l)
        End If
    Next
    Dim A(1000) As String
    '-------------
    If Dots > 0 Then
        m = StartFrame - 1
        For l = 1 To Dots
                TotalFrame = DotPoint(l) - m + 1
                A(1) = String(4 - LenB(StrConv(TotalFrame - 1, vbFromUnicode)), " ") & TotalFrame - 1
                A(1) = A(1) & ", " & String(4 - LenB(StrConv(servosend, vbFromUnicode)), " ") & servosend
                    For i = 8 To (servosend + 7)
                    A(1) = A(1) & ", " & String(4 - LenB(StrConv(Cells(i, 2).Value, vbFromUnicode)), " ") & Cells(i, 2).Value
                    'Format(Cells(i, 2).Value, "0000")
                    Next i
                A(2) = String(4 - LenB(StrConv(Cells(5, 8).Value, vbFromUnicode)), " ") & Cells(5, 8).Value
                A(2) = A(2) & ", " & String(4 - LenB(StrConv(Cells(6, 8).Value, vbFromUnicode)), " ") & Cells(6, 8).Value
                    For i = 1 To servosend
                        If (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value) < 0 Then
                        A(2) = A(2) & "," & String(4 - LenB(StrConv(-10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ") & 10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
                        Else
                        A(2) = A(2) & ", " & String(4 - LenB(StrConv(10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ") & 10 * (Cells(i + 7, 5).Value + Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
                        End If
                    Next i
                If DotPoint(l) > m Then
                    For i = 3 To (DotPoint(l) - m) + 2
                    A(i) = String(4 - LenB(StrConv(Cells(5, m + i + 6).Value, vbFromUnicode)), " ") & Cells(5, m + i + 6).Value
                    'Format(Cells(5, m + i + 6).Value, "0000")
                    A(i) = A(i) & ", " & String(4 - LenB(StrConv(Cells(6, m + i + 6).Value, vbFromUnicode)), " ") & Cells(6, m + i + 6).Value
                    'Format(Cells(6, m + i + 6).Value, "0000")
                        For j = 1 To servosend
                            If Cells(7 + j, m + i + 6).Value < 0 Then
                            A(i) = A(i) & "," & String(4 - LenB(StrConv(-10 * Cells(7 + j, m + i + 6).Value, vbFromUnicode)), " ") & 10 * Cells(7 + j, m + i + 6).Value
                            'Format(10 * Cells(7 + j, m + i + 6).Value, "0000")
                            Else
                            A(i) = A(i) & ", " & String(4 - LenB(StrConv(10 * Cells(7 + j, m + i + 6).Value, vbFromUnicode)), " ") & 10 * Cells(7 + j, m + i + 6).Value
                            'Format(10 * Cells(7 + j, m + i + 6).Value, "0000")
                            End If
                        Next j
                    Next i
                'End If
                'If TotalFrame > 0 Then 'こっちにすると何もないと出力しなくなる
                    Print #IntFlNo, "int " & ActiveSheet.Name & "_Motion_" & l & "[" & TotalFrame + 1 & "][" & servosend + 2 & "]={"
                    Print #IntFlNo, "{" & A(1) & "},//(一個目モーションの総フレーム数、二個目サーボの総数、三個目以降が指定したID)"
                    Print #IntFlNo, "{" & A(2) & "},//初期姿勢(一個目移動時間、二個目待機時間、三個目以降角度)"
                        If TotalFrame > 2 Then
                        Print #IntFlNo, "{" & A(3) & "},//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
                            If TotalFrame > 2 Then
                            For i = 4 To (DotPoint(l) - m) + 1
                            Print #IntFlNo, "{" & A(i) & "},"
                            Next i
                            End If
                        Print #IntFlNo, "{" & A((DotPoint(l) - m) + 2) & "}"
                        Else
                        Print #IntFlNo, "{" & A(3) & "}//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
                        End If
                    Print #IntFlNo, "};"
                    Print #IntFlNo, ""
                End If
                
                m = DotPoint(l)
        Next l
    End If
    Close #IntFlNo
End Sub







Sub MotionExport_Move_2()

    '----------------------------------
    'motion_exportフォルダがなければ作成する
    
    Dim MotionExportDirectry As String
    MotionExportDirectry = ActiveWorkbook.Path & "\motion_export"
    If Dir(MotionExportDirectry, vbDirectory) = "" Then
        MkDir MotionExportDirectry
    End If
    
    '----------------------------------
    '出力ファイル名を決定
    Dim Filename As String
    Filename = GetFNameFromFStr(ActiveWorkbook.Name) & "_" & ActiveSheet.Name
    Dim OutputFile As String
    OutputFile = ActiveWorkbook.Path & "\motion_export\" & Filename & ".c"
    
    '-----------------------------------
    
    Dim i As Long, LngLoop As Long
    Dim IntFlNo As Integer
    LngLoop = Range("a65536").End(xlUp).Row
    IntFlNo = FreeFile
    Open OutputFile For Output As #IntFlNo
    
    Dim Shtname As String
    Shtname = ActiveSheet.Name
    Dim StartFrame As Integer, EndFrame As Integer, LoopStart As Integer, LoopEnd As Integer
    StartFrame = 0
    EndFrame = 0
    LoopStart = 0
    LoopEnd = 0
    '-------------
    '開始位置を調べる
    For k = 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 1 Then
            StartFrame = k - 8
            EndFrame = StartFrame
            k = 68
        End If
    Next
   
    '-------------
    '終了位置を調べる
    For k = StartFrame + 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 2 Then
            EndFrame = k - 8
            k = 68
        End If
    Next
    
    '-------------
    'ループ位置を調べる
    For k = (StartFrame + 8) To (EndFrame + 8)
        If Not Sheets(Shtname).Cells(1, k) = "" Then
            LoopStart = Sheets(Shtname).Cells(1, k).Value
            LoopEnd = k - 8
            k = EndFrame + 8
        End If
    Next
    
    '-------------
    
    
    Dim TotalFrame As Integer
    '----------------------------------
    'スターティングモーション
    If LoopStart <> 0 Then
    TotalFrame = LoopStart - StartFrame
    Else
        If EndFrame > StartFrame Then
        TotalFrame = EndFrame - StartFrame + 1
        Else
        TotalFrame = 0
        End If
    End If
    
    
    Dim A(1000) As String
    
    A(1) = String(4 - LenB(StrConv(TotalFrame, vbFromUnicode)), " ") & TotalFrame
    A(1) = A(1) & ", " & String(4 - LenB(StrConv(servosend, vbFromUnicode)), " ") & servosend
    For i = 8 To (servosend + 7)
        A(1) = A(1) & ", " & String(4 - LenB(StrConv(Cells(i, 2).Value, vbFromUnicode)), " ") & Cells(i, 2).Value
    Next i
        
    A(2) = String(4 - LenB(StrConv(Cells(5, 8).Value, vbFromUnicode)), " ") & Cells(5, 8).Value
    A(2) = A(2) & ", " & String(4 - LenB(StrConv(Cells(6, 8).Value, vbFromUnicode)), " ") & Cells(6, 8).Value
    For i = 1 To servosend
        If (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value) < 0 Then '(-1 ^ Cells(i + 7, 5).Value) *
            A(2) = A(2) & ","
            A(2) = A(2) & String(4 - LenB(StrConv(-10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ")
            A(2) = A(2) & 10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        Else
            A(2) = A(2) & ", "
            A(2) = A(2) & String(4 - LenB(StrConv(10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ")
            A(2) = A(2) & 10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        End If
    Next i
    
    If TotalFrame > 0 Then
        For i = 3 To TotalFrame + 2
            A(i) = String(4 - LenB(StrConv(Cells(5, 7 + StartFrame + i - 2).Value, vbFromUnicode)), " ") & Cells(5, 7 + StartFrame + i - 2).Value
            A(i) = A(i) & ", " & String(4 - LenB(StrConv(Cells(6, 7 + StartFrame + i - 2).Value, vbFromUnicode)), " ") & Cells(6, 7 + StartFrame + i - 2).Value
            For j = 1 To servosend
                If (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * Cells(7 + j, i + StartFrame + 5).Value < 0 Then
                    A(i) = A(i) & ","
                    A(i) = A(i) & String(4 - LenB(StrConv(-10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + StartFrame + 5).Value), vbFromUnicode)), " ")
                    A(i) = A(i) & 10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + StartFrame + 5).Value)
                Else
                    A(i) = A(i) & ", "
                    A(i) = A(i) & String(4 - LenB(StrConv(10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + StartFrame + 5).Value), vbFromUnicode)), " ")
                    A(i) = A(i) & 10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + StartFrame + 5).Value)
                End If
            Next j
        Next i
    End If
    
    If TotalFrame > 0 Then
    Print #IntFlNo, "int " & ActiveSheet.Name & "_Motion_Start[" & TotalFrame + 2 & "][" & servosend + 2 & "]={"
    Print #IntFlNo, "{" & A(1) & "},//(一個目モーションの総フレーム数、二個目サーボの総数、三個目以降が指定したID)"
    Print #IntFlNo, "{" & A(2) & "},//初期姿勢(一個目移動時間、二個目待機時間、三個目以降角度)"
        If TotalFrame > 1 Then
        Print #IntFlNo, "{" & A(3) & "},//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
            If TotalFrame > 2 Then
            For i = 4 To (TotalFrame + 1)
            Print #IntFlNo, "{" & A(i) & "},"
            Next i
            End If
        Print #IntFlNo, "{" & A(TotalFrame + 2) & "}"
        Else
        Print #IntFlNo, "{" & A(3) & "}//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
        End If
    Print #IntFlNo, "};"
    Print #IntFlNo, ""
    End If
    

    '----------------------------------------
    'ループモーション
    If LoopStart = 0 Then
    TotalFrame = 0
    Else
        If LoopEnd = 0 Then
        TotalFrame = 0
        Else
            If LoopEnd - LoopStart < 0 Then
            TotalFrame = 0
            Else
            TotalFrame = LoopEnd - LoopStart + 1
            End If
        End If
    End If
    

    
    A(1) = String(4 - LenB(StrConv(TotalFrame, vbFromUnicode)), " ") & TotalFrame
    A(1) = A(1) & ", " & String(4 - LenB(StrConv(servosend, vbFromUnicode)), " ") & servosend
    For i = 8 To (servosend + 7)
        A(1) = A(1) & ", " & String(4 - LenB(StrConv(Cells(i, 2).Value, vbFromUnicode)), " ") & Cells(i, 2).Value
    Next i
        
    A(2) = String(4 - LenB(StrConv(Cells(5, 8).Value, vbFromUnicode)), " ") & Cells(5, 8).Value
    A(2) = A(2) & ", " & String(4 - LenB(StrConv(Cells(6, 8).Value, vbFromUnicode)), " ") & Cells(6, 8).Value
    For i = 1 To servosend
        If (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value) < 0 Then '(-1 ^ Cells(i + 7, 5).Value) *
            A(2) = A(2) & ","
            A(2) = A(2) & String(4 - LenB(StrConv(-10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ")
            A(2) = A(2) & 10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        Else
            A(2) = A(2) & ", "
            A(2) = A(2) & String(4 - LenB(StrConv(10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ")
            A(2) = A(2) & 10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        End If
    Next i
        
    If TotalFrame > 0 Then
        For i = 3 To TotalFrame + 2
            A(i) = String(4 - LenB(StrConv(Cells(5, 7 + LoopStart + i - 2).Value, vbFromUnicode)), " ") & Cells(5, 7 + LoopStart + i - 2).Value
            A(i) = A(i) & ", " & String(4 - LenB(StrConv(Cells(6, 7 + LoopStart + i - 2).Value, vbFromUnicode)), " ") & Cells(6, 7 + LoopStart + i - 2).Value
            For j = 1 To servosend
                If (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * Cells(7 + j, i + LoopStart + 5).Value < 0 Then
                    A(i) = A(i) & ","
                    A(i) = A(i) & String(4 - LenB(StrConv(-10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + LoopStart + 5).Value), vbFromUnicode)), " ")
                    A(i) = A(i) & 10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + LoopStart + 5).Value)
                Else
                    A(i) = A(i) & ", "
                    A(i) = A(i) & String(4 - LenB(StrConv(10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + LoopStart + 5).Value), vbFromUnicode)), " ")
                    A(i) = A(i) & 10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + LoopStart + 5).Value)
                End If
            Next j
        Next i
    End If
    
    
    If TotalFrame > 0 Then
    Print #IntFlNo, "int " & ActiveSheet.Name & "_Motion_Loop[" & TotalFrame + 2 & "][" & servosend + 2 & "]={"
    Print #IntFlNo, "{" & A(1) & "},//(一個目モーションの総フレーム数、二個目サーボの総数、三個目以降が指定したID)"
    Print #IntFlNo, "{" & A(2) & "},//初期姿勢(一個目移動時間、二個目待機時間、三個目以降角度)"
        If TotalFrame > 1 Then
        Print #IntFlNo, "{" & A(3) & "},//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
            If TotalFrame > 2 Then
            For i = 4 To (TotalFrame + 1)
            Print #IntFlNo, "{" & A(i) & "},"
            Next i
            End If
        Print #IntFlNo, "{" & A(TotalFrame + 2) & "}"
        Else
        Print #IntFlNo, "{" & A(3) & "}//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
        End If
    Print #IntFlNo, "};"
    Print #IntFlNo, ""
    End If
    '----------------------------------------
    'エンドモーション
    If LoopEnd = 0 Then
    TotalFrame = 0
    Else
        If LoopEnd = 0 Then
        TotalFrame = 0
        Else
            If LoopEnd - LoopStart < 0 Then
                TotalFrame = 0
            Else
                TotalFrame = EndFrame - LoopEnd
            End If
        End If
    End If

    A(1) = String(4 - LenB(StrConv(TotalFrame, vbFromUnicode)), " ") & TotalFrame
    A(1) = A(1) & ", " & String(4 - LenB(StrConv(servosend, vbFromUnicode)), " ") & servosend
    For i = 8 To (servosend + 7)
        A(1) = A(1) & ", " & String(4 - LenB(StrConv(Cells(i, 2).Value, vbFromUnicode)), " ") & Cells(i, 2).Value
    Next i
        
    A(2) = String(4 - LenB(StrConv(Cells(5, 8).Value, vbFromUnicode)), " ") & Cells(5, 8).Value
    A(2) = A(2) & ", " & String(4 - LenB(StrConv(Cells(6, 8).Value, vbFromUnicode)), " ") & Cells(6, 8).Value
    For i = 1 To servosend
        If (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value) < 0 Then '(-1 ^ Cells(i + 7, 5).Value) *
            A(2) = A(2) & ","
            A(2) = A(2) & String(4 - LenB(StrConv(-10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ")
            A(2) = A(2) & 10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        Else
            A(2) = A(2) & ", "
            A(2) = A(2) & String(4 - LenB(StrConv(10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ")
            A(2) = A(2) & 10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
        End If
    Next i
    
    If TotalFrame > 0 Then
        For i = 3 To TotalFrame + 2
            A(i) = String(4 - LenB(StrConv(Cells(5, 8 + LoopEnd + i - 2).Value, vbFromUnicode)), " ") & Cells(5, 7 + LoopEnd + 1 + i - 2).Value
            A(i) = A(i) & ", " & String(4 - LenB(StrConv(Cells(6, 8 + LoopEnd + i - 2).Value, vbFromUnicode)), " ") & Cells(6, 8 + LoopEnd + i - 2).Value
            For j = 1 To servosend
                If (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * Cells(7 + j, i + LoopEnd + 6).Value < 0 Then
                    A(i) = A(i) & ","
                    A(i) = A(i) & String(4 - LenB(StrConv(-10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + LoopEnd + 6).Value), vbFromUnicode)), " ")
                    A(i) = A(i) & 10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + LoopEnd + 6).Value)
                Else
                    A(i) = A(i) & ", "
                    A(i) = A(i) & String(4 - LenB(StrConv(10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + LoopEnd + 6).Value), vbFromUnicode)), " ")
                    A(i) = A(i) & 10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + LoopEnd + 6).Value)
                End If
            Next j
        Next i
    End If
    
    
    If TotalFrame > 0 Then
    Print #IntFlNo, "int " & ActiveSheet.Name & "_Motion_End[" & TotalFrame + 2 & "][" & servosend + 2 & "]={"
    Print #IntFlNo, "{" & A(1) & "},//(一個目モーションの総フレーム数、二個目サーボの総数、三個目以降が指定したID)"
    Print #IntFlNo, "{" & A(2) & "},//初期姿勢(一個目移動時間、二個目待機時間、三個目以降角度)"
        If TotalFrame > 1 Then
        Print #IntFlNo, "{" & A(3) & "},//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
            If TotalFrame > 2 Then
            For i = 4 To (TotalFrame + 1)
            Print #IntFlNo, "{" & A(i) & "},"
            Next i
            End If
        Print #IntFlNo, "{" & A(TotalFrame + 2) & "}"
        Else
        Print #IntFlNo, "{" & A(3) & "}//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
        End If
    Print #IntFlNo, "};"
    Print #IntFlNo, ""
    Else
    End If
    

    '----------------------------------------
    
      
    Close #IntFlNo
End Sub



Sub MotionExport_Atk_2()
    '----------------------------------
    'motion_exportフォルダがなければ作成する
    Dim MotionExportDirectry As String
    MotionExportDirectry = ActiveWorkbook.Path & "\motion_export"
    If Dir(MotionExportDirectry, vbDirectory) = "" Then
        MkDir MotionExportDirectry
    End If
    '----------------------------------
    '出力ファイル名を決定
    Dim Filename As String
    Filename = GetFNameFromFStr(ActiveWorkbook.Name) & "_" & ActiveSheet.Name
    Dim OutputFile As String
    OutputFile = ActiveWorkbook.Path & "\motion_export\" & Filename & ".c"
    '-----------------------------------
    Dim i As Long, LngLoop As Long
    Dim IntFlNo As Integer
    LngLoop = Range("a65536").End(xlUp).Row
    IntFlNo = FreeFile
    Open OutputFile For Output As #IntFlNo
    Dim Shtname As String
    Shtname = ActiveSheet.Name
    Dim StartFrame As Integer, EndFrame As Integer, LoopStart As Integer, LoopEnd As Integer
    StartFrame = 0
    EndFrame = 0
    LoopStart = 0
    LoopEnd = 0
    '-------------
    '開始位置を調べる
    For k = 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 1 Then
            StartFrame = k - 8
            EndFrame = StartFrame
            k = 68
        End If
    Next
    '-------------
    '終了位置を調べる
    For k = StartFrame + 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 2 Then
            EndFrame = k - 8
            k = 68
        End If
    Next
    '-------------
    '分割位置を調べる
    Dim DotPoint(100) As Integer
    Dim Dots As Integer
    Dim l As Integer, m As Integer
    m = (StartFrame + 8)
    For l = 1 To 100
        DotPoint(l) = 0
    Next
    For l = 1 To 100
        For k = m To (EndFrame + 8)
            If Sheets(Shtname).Cells(1, k) = "d" Then
            m = k + 1
            DotPoint(l) = k - 8
            k = EndFrame + 8
            End If
        Next
    Next
    For l = 1 To 100
        If DotPoint(l) = 0 Then
        Dots = l
        DotPoint(l) = EndFrame
        l = 100
        Else
        End If
    Next
    Dim A(1000) As String
    '-------------
    If Dots > 0 Then
        m = StartFrame - 1
        For l = 1 To Dots
                TotalFrame = DotPoint(l) - m + 1
                A(1) = String(4 - LenB(StrConv(TotalFrame - 1, vbFromUnicode)), " ") & TotalFrame - 1
                A(1) = A(1) & ", " & String(4 - LenB(StrConv(servosend, vbFromUnicode)), " ") & servosend
                For i = 8 To (servosend + 7)
                    A(1) = A(1) & ", " & String(4 - LenB(StrConv(Cells(i, 2).Value, vbFromUnicode)), " ") & Cells(i, 2).Value
                Next i
                A(2) = String(4 - LenB(StrConv(Cells(5, 8).Value, vbFromUnicode)), " ") & Cells(5, 8).Value
                A(2) = A(2) & ", " & String(4 - LenB(StrConv(Cells(6, 8).Value, vbFromUnicode)), " ") & Cells(6, 8).Value
                For i = 1 To servosend
                    If (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value) < 0 Then '(-1 ^ Cells(i + 7, 5).Value) *
                        A(2) = A(2) & ","
                        A(2) = A(2) & String(4 - LenB(StrConv(-10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ")
                        A(2) = A(2) & 10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
                    Else
                        A(2) = A(2) & ", "
                        A(2) = A(2) & String(4 - LenB(StrConv(10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value), vbFromUnicode)), " ")
                        A(2) = A(2) & 10 * (1 - 2 * ((Cells(i + 7, 5).Value) Mod 2)) * (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value)
                    End If
                Next i
                If DotPoint(l) > m Then
                    For i = 3 To (DotPoint(l) - m) + 2
                        A(i) = String(4 - LenB(StrConv(Cells(5, m + i + 6).Value, vbFromUnicode)), " ") & Cells(5, m + i + 6).Value
                        A(i) = A(i) & ", " & String(4 - LenB(StrConv(Cells(6, m + i + 6).Value, vbFromUnicode)), " ") & Cells(6, m + i + 6).Value
                        For j = 1 To servosend
                            If (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * Cells(7 + j, i + m + 6).Value < 0 Then
                                A(i) = A(i) & ","
                                A(i) = A(i) & String(4 - LenB(StrConv(-10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + m + 6).Value), vbFromUnicode)), " ")
                                A(i) = A(i) & 10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + m + 6).Value)
                            Else
                                A(i) = A(i) & ", "
                                A(i) = A(i) & String(4 - LenB(StrConv(10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + m + 6).Value), vbFromUnicode)), " ")
                                A(i) = A(i) & 10 * (1 - 2 * ((Cells(j + 7, 5).Value) Mod 2)) * CInt(Cells(7 + j, i + m + 6).Value)
                            End If
                        Next j
                    Next i
                    
                    
                    
                    
                    
                    Print #IntFlNo, "int " & ActiveSheet.Name & "_Motion_" & l & "[" & TotalFrame + 1 & "][" & servosend + 2 & "]={"
                    Print #IntFlNo, "{" & A(1) & "},//(一個目モーションの総フレーム数、二個目サーボの総数、三個目以降が指定したID)"
                    Print #IntFlNo, "{" & A(2) & "},//初期姿勢(一個目移動時間、二個目待機時間、三個目以降角度)"
                        If TotalFrame > 2 Then
                        Print #IntFlNo, "{" & A(3) & "},//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
                            If TotalFrame > 2 Then
                            For i = 4 To (DotPoint(l) - m) + 1
                            Print #IntFlNo, "{" & A(i) & "},"
                            Next i
                            End If
                        Print #IntFlNo, "{" & A((DotPoint(l) - m) + 2) & "}"
                        Else
                        Print #IntFlNo, "{" & A(3) & "}//以下モーションデータ(一個目移動時間、二個目待機時間、三個目以降角度)"
                        End If
                    Print #IntFlNo, "};"
                    Print #IntFlNo, ""
                End If
                m = DotPoint(l)
        Next l
    End If
    Close #IntFlNo
End Sub



Sub Change_Right_Left()
    Application.ScreenUpdating = False
    Dim Shtname As String
    Shtname = ActiveSheet.Name
    Dim StartFrame As Integer, EndFrame As Integer, LoopStart As Integer, LoopEnd As Integer
    StartFrame = 0
    EndFrame = 0
    LoopStart = 0
    LoopEnd = 0
    '-------------
    '開始位置を調べる
    For k = 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 1 Then
            StartFrame = k - 8
            EndFrame = StartFrame
            k = 68
        End If
    Next
    '-------------
    '終了位置を調べる
    For k = StartFrame + 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 2 Then
            EndFrame = k - 8
            k = 68
        End If
    Next
    '-------------
    For i = 1 To servosend
        If (Cells(i + 7, 4).Value = "o") Then
            For k = StartFrame To EndFrame
                Cells(i + 7, k + 8).Value = -1 * Cells(i + 7, k + 8).Value
            Next k
        Else
        For j = i + 1 To servosend
                If Cells(i + 7, 4).Value = Cells(j + 7, 4).Value And Not (Cells(i + 7, 4).Value = "") And Not (Cells(j + 7, 4).Value = "") Then
                Dim temp As Integer
                For k = StartFrame To EndFrame
                    temp = Cells(i + 7, k + 8).Value
                    Cells(i + 7, k + 8).Value = Cells(j + 7, k + 8).Value
                    Cells(j + 7, k + 8).Value = temp
                Next k
            End If
        Next j
        End If
    Next i
    
    
End Sub
Sub Change_Start_End()
    Application.ScreenUpdating = False
    Dim Shtname As String
    Shtname = ActiveSheet.Name
    Dim StartFrame As Integer, EndFrame As Integer, LoopStart As Integer, LoopEnd As Integer
    StartFrame = 0
    EndFrame = 0
    LoopStart = 0
    LoopEnd = 0
    '-------------
    '開始位置を調べる
    For k = 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 1 Then
            StartFrame = k - 8
            EndFrame = StartFrame
            k = 68
        End If
    Next
    '-------------
    '終了位置を調べる
    For k = StartFrame + 8 To 68
        If Sheets(Shtname).Cells(7, k).Value = 2 Then
            EndFrame = k - 8
            k = 68
        End If
    Next
    '-------------
    'ループ位置を調べる
    For k = (StartFrame + 8) To (EndFrame + 8)
        If Not Sheets(Shtname).Cells(1, k) = "" Then
            LoopStart = Sheets(Shtname).Cells(1, k).Value
            LoopEnd = k - 8
            k = EndFrame + 8
        End If
    Next
    '-------------
    
    'LOOP
    For j = 0 To (EndFrame - StartFrame)
        If (j = (LoopEnd - LoopStart) + (EndFrame - LoopEnd)) Then
            Cells(1, j + StartFrame + 8).Value = LoopStart + ((EndFrame - LoopEnd) - (LoopStart - StartFrame))
        Else
            Cells(1, j + StartFrame + 8).Value = ""
        End If
    Next j
    
    Dim temp(1000) As String
    'Comment
    For j = 0 To (EndFrame - StartFrame)
        temp(j) = Cells(3, j + StartFrame + 8).Value
    Next j
    For j = 0 To (EndFrame - StartFrame)
        If (j < EndFrame - LoopEnd) Then
            Cells(3, j + StartFrame + 8).Value = temp((EndFrame - StartFrame) - j)
        ElseIf (j <= EndFrame - LoopStart) Then
            Cells(3, j + StartFrame + 8).Value = temp((EndFrame - StartFrame) - j)
        ElseIf (j <= EndFrame - StartFrame) Then
            Cells(3, j + StartFrame + 8).Value = temp((EndFrame - StartFrame) - j)
        End If
    Next j
    'FLAG
    For j = 0 To (EndFrame - StartFrame)
        temp(j) = Cells(7, j + StartFrame + 8).Value
    Next j
    For j = 0 To (EndFrame - StartFrame)
        Cells(7, j + StartFrame + 8).Value = temp((EndFrame - StartFrame) - j)
    Next j
    For j = 0 To (EndFrame - StartFrame)
        If (Cells(7, j + StartFrame + 8).Value = "2") Then
            Cells(7, j + StartFrame + 8).Value = "1"
            For k = j + 1 To (EndFrame - StartFrame)
                If (Cells(7, k + StartFrame + 8).Value = "1") Then
                    Cells(7, k + StartFrame + 8).Value = "2"
                    k = (EndFrame - StartFrame)
                    j = (EndFrame - StartFrame)
                End If
            Next k
        End If
    Next j
    'TIMES
    For i = 0 To 1
        For j = 0 To (EndFrame - StartFrame)
            temp(j) = Cells(5 + i, j + StartFrame + 8).Value
        Next j
        For j = 0 To (EndFrame - StartFrame)
            Cells(5 + i, j + StartFrame + 8).Value = temp((EndFrame - StartFrame) - j)
        Next j
    Next i

    'servos
    For i = 1 To servosend
        If (Cells(i + 7, 4).Value = "n") Then
        Else
            For j = 0 To (EndFrame - StartFrame)
                temp(j) = Cells(7 + i, j + StartFrame + 8).Value
            Next j
            For j = 0 To (EndFrame - StartFrame)
                Cells(7 + i, j + StartFrame + 8).Value = temp((EndFrame - StartFrame) - j)
            Next j
        End If
    Next i
End Sub


Sub MotionExport2()
    If Cells(2, 2) = "a" Then
        Call MotionExport_Atk_2
    Else
        Call MotionExport_Move_2
    End If
End Sub


