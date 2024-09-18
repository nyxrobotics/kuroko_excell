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
    Print #IntFlNo, "  - section_name: start_section"
    Print #IntFlNo, "    next_sections: []"
    Print #IntFlNo, "    joint_trajectory:"
    Print #IntFlNo, "      joint_names:"
    
    ' ジョイント名の出力
    For i = LBound(joint_names) To UBound(joint_names)
        Print #IntFlNo, "        - " & joint_names(i)
    Next i
    
    ' フレームごとの処理
    Print #IntFlNo, "      points:"
    
    Dim col As Integer ' 0番フレームは8列目に対応
    Dim prev_col As Integer ' 直前のフレーム（skipでないもの）
    prev_col = -1
    
    col = 8 ' フレームの開始位置は8列目（0番フレーム）
    Dim processing As Boolean
    processing = False ' モーション開始フレームを見つけるまでフレームを無視する

    Do While Not IsEmpty(Cells(7, col).Value)
        Dim flag As Integer
        flag = Cells(7, col).Value
        
        ' STARTフレーム (1) が見つかるまで処理をスキップ
        If flag = 1 Then
            processing = True
        End If
        
        ' STARTからENDの間、SKIPを除いて処理
        If processing And (flag = 1 Or flag = 0 Or flag = 2) Then
            ' ジョイントの位置、速度、加速度、努力の出力
            Dim positions As String
            positions = ""
            For i = 0 To total_joints - 1
                Dim position_with_home As Double
                Dim reverse_flag As Integer
                reverse_flag = Cells(i + 8, 5).Value ' 5列目が符号反転フラグ
                
                position_with_home = Cells(i + 8, col).Value + Cells(i + 8, 6).Value ' 現在の角度にHome（初期位置）を加算
                
                ' 反転フラグが1の場合、角度に-1を掛ける
                If reverse_flag = 1 Then
                    position_with_home = position_with_home * -1
                End If
                
                ' 角度をdegreeからradianに変換し、小数点以下3桁に丸める
                position_with_home = Round(position_with_home * WorksheetFunction.Pi() / 180, 3)
                
                If i > 0 Then
                    positions = positions & ", " & Format(position_with_home, "0.000")
                Else
                    positions = Format(position_with_home, "0.000")
                End If
            Next i
            
            ' velocitiesの計算
            Dim velocities As String
            velocities = ""
            Dim movingTime As Double
            movingTime = Cells(5, col).Value ' 5列目はMovingTime
            
            ' 初回フレームやMovingTimeが0なら速度を0に
            If prev_col = -1 Or movingTime = 0 Then
                For i = 0 To total_joints - 1
                    If i > 0 Then
                        velocities = velocities & ", 0.000"
                    Else
                        velocities = "0.000"
                    End If
                Next i
            Else
                ' 速度を計算（直前のskipでないフレームとの差）
                For i = 0 To total_joints - 1
                    Dim velocity As Double
                    ' 角度差をラジアンで計算し、時間に対して割る (rad/sec)
                    velocity = (Cells(i + 8, col).Value - Cells(i + 8, prev_col).Value) * WorksheetFunction.Pi() / 180 / (movingTime / 100)
                    
                    ' reverseフラグが1の場合、速度の符号を反転
                    If Cells(i + 8, 5).Value = 1 Then
                        velocity = velocity * -1
                    End If
                    
                    If i > 0 Then
                        velocities = velocities & ", " & Format(Round(velocity, 3), "0.000")
                    Else
                        velocities = Format(Round(velocity, 3), "0.000")
                    End If
                Next i
            End If
            
            ' accelerationsとeffortを0で埋める
            Dim accelerations As String
            Dim effort As String
            accelerations = "0.000"
            effort = "0.000"
            For i = 1 To total_joints - 1
                accelerations = accelerations & ", 0.000"
                effort = effort & ", 0.000"
            Next i
            
            ' MovingTime（5行目）とWaitingTime（6行目）を使ってtime_from_startを計算
            Dim time_from_start As Double
            time_from_start = (Cells(5, col).Value + Cells(6, col).Value) / 100 ' 5列目と6列目の和を100で割る

            ' YAML形式で出力
            Print #IntFlNo, "        - positions: [" & positions & "]"
            Print #IntFlNo, "          velocities: [" & velocities & "]"
            Print #IntFlNo, "          accelerations: [" & accelerations & "]"
            Print #IntFlNo, "          effort: [" & effort & "]"
            Print #IntFlNo, "          time_from_start: " & Format(time_from_start, "0.000") ' MovingTimeとWaitingTimeを合計した値
            
            ' 現在のフレームを前回のフレームとして保存
            prev_col = col
        End If
        
        ' 終了フラグ (2) ならループを終了
        If flag = 2 Then
            Exit Do
        End If
        
        col = col + 1
    Loop
    
    ' ファイルを閉じる
    Close #IntFlNo

End Sub





Sub MotionExport_InitialPose_Yaml()

    ' 初期設定
    Dim MotionExportDirectry As String
    MotionExportDirectry = ActiveWorkbook.Path & "\motion_export"
    If Dir(MotionExportDirectry, vbDirectory) = "" Then
        MkDir MotionExportDirectry
    End If
    
    Dim Filename As String
    Filename = "initial_pose.yaml"
    Dim OutputFile As String
    OutputFile = ActiveWorkbook.Path & "\motion_export\" & Filename
    
    ' ファイルを開く
    Dim IntFlNo As Integer
    IntFlNo = FreeFile
    Open OutputFile For Output As #IntFlNo
    
    ' YAML出力の初期行
    Print #IntFlNo, "initial_pose:"
    
    ' Joint namesと初期位置の取得・書き出し
    Dim row As Integer
    row = 8
    Do While Not IsEmpty(Cells(row, 1).Value)
        Dim joint_name As String
        joint_name = Cells(row, 1).Value
        Dim initial_pose As Double
        initial_pose = Cells(row, 6).Value ' 6列目が初期位置
        Dim reverse As Integer
        reverse = Cells(row, 5).Value ' 5列目がreverseフラグ
        
        ' reverseフラグが1のときは符号反転
        If reverse = 1 Then
            initial_pose = -initial_pose
        End If
        
        initial_pose = Round(initial_pose * WorksheetFunction.Pi() / 180, 3) ' 度をラジアンに変換し、小数点以下3桁に丸める
        Print #IntFlNo, "  " & joint_name & ": " & Format(initial_pose, "0.000")
        row = row + 1
    Loop
    
    ' ファイルを閉じる
    Close #IntFlNo

End Sub


Sub MotionExport_Offset_Yaml()

    ' 初期設定
    Dim MotionExportDirectry As String
    MotionExportDirectry = ActiveWorkbook.Path & "\motion_export"
    If Dir(MotionExportDirectry, vbDirectory) = "" Then
        MkDir MotionExportDirectry
    End If
    
    Dim Filename As String
    Filename = "offset.yaml"
    Dim OutputFile As String
    OutputFile = ActiveWorkbook.Path & "\motion_export\" & Filename
    
    ' ファイルを開く
    Dim IntFlNo As Integer
    IntFlNo = FreeFile
    Open OutputFile For Output As #IntFlNo
    
    ' YAML出力の初期行
    Print #IntFlNo, "offset:"
    
    ' Joint namesとオフセット値の取得・書き出し
    Dim row As Integer
    row = 8
    Do While Not IsEmpty(Cells(row, 1).Value)
        Dim joint_name As String
        joint_name = Cells(row, 1).Value
        Dim offset As Double
        offset = Cells(row, 7).Value ' 7列目がオフセット値
        Dim reverse As Integer
        reverse = Cells(row, 5).Value ' 5列目がreverseフラグ
        
        ' reverseフラグが1のときは符号反転
        If reverse = 1 Then
            offset = -offset
        End If
        
        offset = Round(offset * WorksheetFunction.Pi() / 180, 3) ' 度をラジアンに変換し、小数点以下3桁に丸める
        Print #IntFlNo, "  " & joint_name & ": " & Format(offset, "0.000")
        row = row + 1
    Loop
    
    ' ファイルを閉じる
    Close #IntFlNo

End Sub



Sub MotionExport()
    Call MotionExport_Yaml
    Call MotionExport_InitialPose_Yaml
    Call MotionExport_Offset_Yaml
End Sub


