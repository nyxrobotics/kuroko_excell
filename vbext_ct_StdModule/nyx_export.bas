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



Sub MotionExport_withLoop_cpp()

    '----------------------------------
    '歩行・起き上がりモーション用
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
    LngLoop = Range("a65536").End(xlUp).row
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



Sub MotionExport_withBranch_cpp()
    '----------------------------------
    '攻撃モーション用
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
    LngLoop = Range("a65536").End(xlUp).row
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


Sub Calc_calibration_offset()
    Application.ScreenUpdating = False
    Dim Shtname As String
    Shtname = ActiveSheet.Name
    Dim StartFrame As Integer, EndFrame As Integer, LoopStart As Integer, LoopEnd As Integer
    StartFrame = 0
    EndFrame = 0
    LoopStart = 0
    LoopEnd = 0
    '-------------
    'Compare the sum of the Home and Offset angles in the left and right pairs, with the average being Home and the difference from the average being Offset.
    For i = 1 To servosend
        If (Cells(i + 7, 4).Value = "o") Then
            Cells(i + 7, 7).Value = Cells(i + 7, 7).Value
            Cells(i + 7, 8).Value = Cells(i + 7, 8).Value
        Else
        For j = i + 1 To servosend
            If Cells(i + 7, 4).Value = Cells(j + 7, 4).Value And Not (Cells(i + 7, 4).Value = "") And Not (Cells(j + 7, 4).Value = "") Then
                Dim average As Currency
                Dim diff_i As Currency
                Dim diff_j As Currency
                average = (Cells(i + 7, 6).Value + Cells(i + 7, 7).Value + Cells(j + 7, 6).Value + Cells(j + 7, 7).Value) * 0.5
                diff_i = Cells(i + 7, 6).Value + Cells(i + 7, 7).Value - average
                diff_j = Cells(j + 7, 6).Value + Cells(j + 7, 7).Value - average
                Cells(i + 7, 6).Value = average
                Cells(i + 7, 7).Value = diff_i
                Cells(j + 7, 6).Value = average
                Cells(j + 7, 7).Value = diff_j
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

Sub MotionExport_Yaml()

    ' 初期設定
    Dim MotionExportDirectry As String
    MotionExportDirectry = ActiveWorkbook.Path & "\motion_export"
    If Dir(MotionExportDirectry, vbDirectory) = "" Then
        MkDir MotionExportDirectry
    End If
    
    Dim Filename As String
    Filename = GetFNameFromFStr(ActiveWorkbook.Name) & "_" & ActiveSheet.Name
    Dim OutputFile As String
    OutputFile = ActiveWorkbook.Path & "\motion_export\" & Filename & ".yaml"
    
    ' ファイルを開く
    Dim IntFlNo As Integer
    IntFlNo = FreeFile
    Open OutputFile For Output As #IntFlNo
    
    ' Joint namesの取得
    Dim joint_names() As String
    Dim total_joints As Integer
    total_joints = 0
    Dim row As Integer
    row = 8
    
    ' 1列目が空欄になるまでのループでジョイント数を取得
    Do While Not IsEmpty(Cells(row, 1).Value)
        total_joints = total_joints + 1
        row = row + 1
    Loop
    
    ' ジョイント配列の再割り当て
    ReDim joint_names(total_joints - 1)
    row = 8
    For i = 0 To total_joints - 1
        joint_names(i) = Cells(row, 1).Value
        row = row + 1
    Next i
    
    ' セクションの出力
    Print #IntFlNo, "sections:"
    
    ' ループの開始・終了フレームをチェック
    Dim loop_start_frame As Integer
    Dim loop_end_frame As Integer
    loop_end_frame = -1
    loop_start_frame = -1
    
    ' 1行目に何かしらの数値が入っている列をループの終了フレームとして設定
    For col = 8 To 100
        If IsNumeric(Cells(1, col).Value) Then
            loop_end_frame = col - 8 ' 1行目の列がループ終了フレーム（0番フレームは8列目）
            loop_start_frame = Cells(1, col).Value ' セルの値がループ開始フレーム
            Exit For
        End If
    Next col
    
    ' セクションの状態管理
    Dim current_section As String
    current_section = "start_section"
    Dim section_written As Boolean
    section_written = False
    Dim end_flag As Boolean
    end_flag = False
    
    ' フレームごとの処理
    row = 8 ' 0番フレームは8列目
    Do While Not IsEmpty(Cells(7, row).Value) And Not end_flag
        Dim flag As Integer
        flag = Cells(7, row).Value
        Dim frame_num As Integer
        frame_num = row - 8 ' 0番フレームが8列目に対応
        
        ' セクションの判定と切り替え
        If frame_num = loop_start_frame Then
            current_section = "loop_section"
            section_written = False
        ElseIf frame_num = loop_end_frame Then
            current_section = "finish_section"
            section_written = False
        End If
        
        ' 終了フラグ (2) の場合はfinish_section
        If flag = 2 Then
            current_section = "finish_section"
            section_written = False
            end_flag = True
        End If
        
        ' 各セクションの出力
        If Not section_written Then
            Select Case current_section
                Case "start_section"
                    If frame_num < loop_start_frame Then
                        Print #IntFlNo, "  - section_name: " & current_section
                        Print #IntFlNo, "    next_sections:"
                        Print #IntFlNo, "      - loop_section"
                        section_written = True
                    End If
                Case "loop_section"
                    Print #IntFlNo, "  - section_name: " & current_section
                    Print #IntFlNo, "    next_sections:"
                    Print #IntFlNo, "      - finish_section"
                    section_written = True
                Case "finish_section"
                    If frame_num >= loop_end_frame Then
                        Print #IntFlNo, "  - section_name: " & current_section
                        section_written = True
                    End If
            End Select
            
            ' ジョイント名の出力
            Print #IntFlNo, "    joint_trajectory:"
            Print #IntFlNo, "      joint_names:"
            For i = LBound(joint_names) To UBound(joint_names)
                Print #IntFlNo, "        - " & joint_names(i)
            Next i
            Print #IntFlNo, "      points:"
        End If
        
        ' ジョイントの位置、速度、加速度、努力の出力
        Dim positions As String
        positions = ""
        For i = 0 To total_joints - 1
            If i > 0 Then
                positions = positions & ", " & Cells(i + 8, row).Value
            Else
                positions = Cells(i + 8, row).Value
            End If
        Next i
        
        ' velocitiesの計算
        Dim velocities As String
        velocities = ""
        Dim movingTime As Double
        movingTime = Cells(4, row).Value
        
        ' 初回フレームやMovingTimeが0なら速度を0に
        If row = 8 Or movingTime = 0 Then
            For i = 0 To total_joints - 1
                If i > 0 Then
                    velocities = velocities & ", 0"
                Else
                    velocities = "0"
                End If
            Next i
        Else
            ' 速度を計算
            For i = 0 To total_joints - 1
                If i > 0 Then
                    velocities = velocities & ", " & (Cells(i + 8, row + 1).Value - Cells(i + 8, row).Value) / (movingTime / 100)
                Else
                    velocities = (Cells(i + 8, row + 1).Value - Cells(i + 8, row).Value) / (movingTime / 100)
                End If
            Next i
        End If
        
        ' accelerationsとeffortを0で埋める
        Dim accelerations As String
        Dim effort As String
        accelerations = "0"
        effort = "0"
        For i = 1 To total_joints - 1
            accelerations = accelerations & ", 0"
            effort = effort & ", 0"
        Next i
        
        ' YAML形式で出力
        Print #IntFlNo, "        - positions: [" & positions & "]"
        Print #IntFlNo, "          velocities: [" & velocities & "]"
        Print #IntFlNo, "          accelerations: [" & accelerations & "]"
        Print #IntFlNo, "          effort: [" & effort & "]"
        Print #IntFlNo, "          time_from_start: " & (movingTime + Cells(5, row).Value) / 100 ' MovingTimeとWaitingTimeを合計した値
        
        row = row + 1
    Loop
    
    ' ファイルを閉じる
    Close #IntFlNo

End Sub




Sub MotionExport()
    If Cells(2, 2) = "a" Then
        Call MotionExport_withBranch_cpp
    Else
        Call MotionExport_withLoop_cpp
    End If
End Sub


