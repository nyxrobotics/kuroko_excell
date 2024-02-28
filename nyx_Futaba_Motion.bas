Attribute VB_Name = "nyx_Futaba_Motion"
'************************************************
'
'   RUPUTEN-Editor32
'
'************************************************

Public posnum, sdeg, deg, load, temp As Integer
Public nyx(100) As Byte
Public nyx_deg(100) As Byte
Public Const servosend As Integer = 23   '実際に送信するサーボの数
Public Const sendbyte  As Integer = 6 + 5 * servosend + 1 '実際に送信するサーボの数
Public Const servonum  As Integer = 32   'データの数
Public Const getstart  As Integer = 8    'データゲットスタートセル
Public Const cellstart As Integer = 8    'データセルの開始
Public Const cellend   As Integer = 39   'データセルの終わり
Public Const comnum    As Integer = 8    'COMポートナンバー




Public return_pos(servosend) As Currency
Public return_tor(servosend) As Integer
Public return_tem(servosend) As Integer
Public return_vol(servosend) As Integer
Public return_err(servosend) As Integer
Public return_single_buf(5) As Currency 'pose,load,temp,volt,error




Sub com_init()
    Application.ScreenUpdating = False
    Dim com_port As Integer
    Dim com_baud As Long
    Dim com_parity As String
    Dim com_length As String
    Dim com_stop As Integer
    
    Dim ec_set As String
    
    Call QPC_start_counting
    baud_rate = Sheets("COM").Cells(2, 2).Value
    com_port = Sheets("COM").Cells(1, 2).Value
    ec_set = SPrintF("%d,n,8,1", baud_rate)
    
    ec.COMn = com_port 'COMを開きます
    ec.Setting = ec_set
    QPC_wait_ms (20)
End Sub

Sub com_init2()
    Application.ScreenUpdating = False
    Call QPC_start_counting
    Call com_init
    QPC_wait_ms (20)
End Sub

Sub servo_on()
    Call QPC_start_counting
    Dim sndbin(8) As Byte
'-------------------------------------
    ec.COMn = comnum                                     'COMを開きます
    ec.Setting = "115200,n,8,1"
    ec.HandShaking = ec.HANDSHAKEs.No
    ec.AsciiLineTimeOut = 2000                          'タイムアウトを2秒に設定します．
    ec.Delimiter = ec.DELIMs.CrLf                       'デリミタをCrLfに設定します．
    ec.OutBuffer = 10& * 1024&
    ec.InBuffer = 10& * 1024&



    sndbin(0) = &HFA 'Header1
    sndbin(1) = &HAF 'Header2
    sndbin(2) = &HFF 'ID:all
    sndbin(3) = &H0
    sndbin(4) = &H24
    sndbin(5) = &H1
    sndbin(6) = &H1
    sndbin(7) = &H1
    '----- checksum ------
    sndbin(8) = 0
    For i = 2 To 7
        sndbin(8) = sndbin(8) Xor sndbin(i)
    Next
    ec.Binary = sndbin()
    QPC_wait_ms (10)
    
    
    sndbin(0) = &HFA 'Header1
    sndbin(1) = &HAF 'Header2
    sndbin(2) = &HFF 'ID:all
    sndbin(3) = &H0
    sndbin(4) = &H23
    sndbin(5) = &H1
    sndbin(6) = &H1
    sndbin(7) = "&H" + Hex(100) 'maximum power(%)
    '----- checksum ------
    sndbin(8) = 0
    For i = 2 To 7
        sndbin(8) = sndbin(8) Xor sndbin(i)
    Next
    ec.Binary = sndbin()
    QPC_wait_ms (10)
     
    ec.COMn = 0
    'QPC_wait_ms (10)                                   '0.01秒待ちます．
End Sub
Sub servo_off()
    Call QPC_start_counting
    Dim sndbin(8) As Byte
'-------------------------------------
    ec.COMn = comnum                                     'COMを開きます
    ec.Setting = "115200,n,8,1"
    ec.HandShaking = ec.HANDSHAKEs.No
    ec.AsciiLineTimeOut = 2000                          'タイムアウトを2秒に設定します．
    ec.Delimiter = ec.DELIMs.CrLf                       'デリミタをCrLfに設定します．
    ec.OutBuffer = 10& * 1024&
    ec.InBuffer = 10& * 1024&
'----- checksum ------
    ChkSum = 0
      
    sndbin(0) = &HFA
    sndbin(1) = &HAF
    sndbin(2) = &HFF
    sndbin(3) = &H0
    sndbin(4) = &H24
    sndbin(5) = &H1
    sndbin(6) = &H1
    sndbin(7) = &H2
    
    For i = 2 To 7
    sndbin(8) = sndbin(8) Xor sndbin(i)
    Next
    
     ec.Binary = sndbin()
         
    ec.COMn = 0
    
    QPC_wait_ms (10)

End Sub


Sub servo_free()
    Call QPC_start_counting
    Dim sndbin(8) As Byte
'-------------------------------------
    ec.COMn = comnum                                     'COMを開きます
    ec.Setting = "115200,n,8,1"
    ec.HandShaking = ec.HANDSHAKEs.No
    ec.AsciiLineTimeOut = 2000                          'タイムアウトを2秒に設定します．
    ec.Delimiter = ec.DELIMs.CrLf                       'デリミタをCrLfに設定します．
    ec.OutBuffer = 10& * 1024&
    ec.InBuffer = 10& * 1024&
'----- checksum ------
    ChkSum = 0
      
    sndbin(0) = &HFA 'Header1
    sndbin(1) = &HAF 'Header2
    sndbin(2) = &HFF 'ID:all
    sndbin(3) = &H0
    sndbin(4) = &H24
    sndbin(5) = &H1
    sndbin(6) = &H1
    sndbin(7) = &H0
    
    For i = 2 To 7
    sndbin(8) = sndbin(8) Xor sndbin(i)
    Next
    
     ec.Binary = sndbin()
     
    ec.COMn = 0

    QPC_wait_ms (200)                                '0.01秒待ちます．
    
End Sub
     




Sub get_single_buf(SID)
    Application.ScreenUpdating = False
    'Call com_init2
    Dim sndbin_2(8) As Byte
    Dim Recbin() As Byte
    Dim string_1() As Byte
      Dim test(1) As Long
    'Request
    sndbin_2(0) = &HFA
    sndbin_2(1) = &HAF
    sndbin_2(2) = "&H" + Hex(SID) 'ID
    sndbin_2(3) = &HF   'Flag 0b00001111
    sndbin_2(4) = &H2A  'Address
    sndbin_2(5) = &HC   'Length
    sndbin_2(6) = &H0   'Memory Write(No)
    sndbin_2(7) = &H0
    For j = 2 To 6
        sndbin_2(7) = sndbin_2(7) Xor sndbin_2(j)
    Next
    ec.Binary = sndbin_2()
    
    'Receive
    QPC_wait_ms (5)
    Recbin() = ec.Binary
    Dim start_point As Integer
    start_point = UBound(Recbin()) 'エラー処理用の初期値
    For i = 0 To UBound(Recbin()) - 2
        If (Recbin(i) = &HFD) And (Recbin(i + 1) = &HDF) And (Recbin(i + 2) = ("&H" + Hex(SID))) Then
            start_point = i
            i = UBound(Recbin()) 'end
        End If
    Next
    
    'Set Values
    If UBound(Recbin()) > start_point + 17 Then 'エラー処理
        If Recbin(start_point + 8) < 128 Then
            return_single_buf(0) = CDec(Recbin(start_point + 8)) * 256 + CDec(Recbin(start_point + 7))
        Else
            return_single_buf(0) = (CDec(Recbin(start_point + 8) - 255)) * 256 + CDec(Recbin(start_point + 7) - 256)
        End If
        
        If Recbin(start_point + 14) < 128 Then
            return_single_buf(1) = CDec(Recbin(start_point + 14)) * 256 + CDec(Recbin(start_point + 13))
        Else
            return_single_buf(1) = (CDec(Recbin(start_point + 14) - 255)) * 256 + CDec(Recbin(start_point + 13) - 256)
        End If
        
        If Recbin(start_point + 16) < 128 Then
            return_single_buf(2) = CDec(Recbin(start_point + 16)) * 256 + CDec(Recbin(start_point + 15))
        Else
            return_single_buf(2) = (CDec(Recbin(start_point + 16) - 255)) * 256 + CDec(Recbin(start_point + 15) - 256)
        End If
        
        If Recbin(start_point + 18) < 128 Then
            return_single_buf(3) = CDec(Recbin(start_point + 18)) * 256 + CDec(Recbin(start_point + 17))
        Else
            return_single_buf(3) = (CDec(Recbin(start_point + 18) - 255)) * 256 + CDec(Recbin(start_point + 17) - 256)
        End If
        'return_single_buf(0) = Recbin(start_point + 8) * 256 + Recbin(start_point + 7)
        'return_single_buf(1) = Recbin(start_point + 14) * 256 + Recbin(start_point + 13)
        'return_single_buf(2) = Recbin(start_point + 16) * 256 + Recbin(start_point + 15)
        'return_single_buf(3) = Recbin(start_point + 18) * 256 + Recbin(start_point + 17)
        return_single_buf(4) = 0
    Else
        return_single_buf(0) = 0
        return_single_buf(1) = 0
        return_single_buf(2) = 0
        return_single_buf(3) = 0
        return_single_buf(4) = 1
    End If
    
    'Ending
    'ec.COMn = 0
End Sub

Sub get_return(samples)
    Application.ScreenUpdating = False
    Shtname = ActiveSheet.Name
    Call com_init2
    Dim id(servosend) As Integer
    Dim pos(servosend) As Currency
    Dim tor(servosend) As Currency
    Dim tem(servosend) As Currency
    Dim vol(servosend) As Currency
    Dim err(servosend) As Byte
    Dim err_0(servosend) As Byte
    QPC_wait_ms (50)
    For j = 0 To samples - 1
        For i = 0 To servosend - 1
            err_0(i) = 0
        Next
        For i = 0 To servosend - 1
            id(i) = Sheets(Shtname).Cells(i + 8, 2).Value
            Call get_single_buf(id(i))
            If (return_single_buf(4) = 0) Then
                pos(i) = pos(i) + (return_single_buf(0) / samples)
                tor(i) = tor(i) + (return_single_buf(1) / samples)
                tem(i) = tem(i) + (return_single_buf(2) / samples)
                vol(i) = vol(i) + (return_single_buf(3) / samples)
            Else
                err_0(i) = err_0(i) + 1
                If (err_0(i) < 5) Then '何回までリトライするか
                    i = i - 1
                Else
                    err(i) = 1
                End If
            End If
        Next
    Next
    
    ec.COMn = 0
    For i = 0 To servosend - 1
        If err(i) = 0 Then
            return_pos(i) = CInt(pos(i))
            return_tor(i) = CInt(tor(i))
            return_tem(i) = CInt(tem(i))
            return_vol(i) = CInt(vol(i))
        Else
            return_pos(i) = CInt(pos(i))
            return_tor(i) = CInt(tor(i))
            return_tem(i) = CInt(tem(i))
            return_vol(i) = CInt(vol(i))
            MsgBox "ID = " & id(i) & " から正しいリターンを得られないようです．", vbInformation
        End If
        return_err(i) = err(i)
    Next
End Sub

Sub view_pos()
    Application.ScreenUpdating = False
    Shtname = ActiveSheet.Name
    Range("C8:C30").Select
    Selection.ClearContents
    Call get_return(5)
    For i = 0 To servosend - 1
        If Sheets(Shtname).Cells(i + 8, 5).Value = "" Or Sheets(Shtname).Cells(i + 8, 5).Value = 0 Then
            Sheets(Shtname).Cells(i + 8, 3) = return_pos(i) / 10
        Else
            Sheets(Shtname).Cells(i + 8, 3) = -1 * return_pos(i) / 10
        End If
        'Sheets(Shtname).Cells(I + 8, 3) = return_pos(I) / 10
        If (return_err(i) = 0) Then
            'Cells(I+8, 3).Font.ColorIndex = 5
        Else
            'Cells(I+8, 3).Font.ColorIndex = 5
        End If
    Next
End Sub

Sub view_tor()
    Application.ScreenUpdating = False
    Shtname = ActiveSheet.Name
    Range("C8:C30").Select
    Selection.ClearContents
    Call get_return(5)
    For i = 0 To servosend - 1
        Sheets(Shtname).Cells(i + 8, 3) = return_tor(i)

        If (return_err(i) = 0) Then
            'Cells(I+8, 3).Font.ColorIndex = 5
        Else
            'Cells(I+8, 3).Font.ColorIndex = 5
        End If
    Next
End Sub

Sub view_tem()
    Application.ScreenUpdating = False
    Shtname = ActiveSheet.Name
    Range("C8:C30").Select
    Selection.ClearContents
    Call get_return(5)
    For i = 0 To servosend - 1
        Sheets(Shtname).Cells(i + 8, 3) = 20 + ((955 - return_tem(i)) * 100 / 39)
        'Sheets(Shtname).Cells(I + 8, 3) = return_pos(I) / 10
        If (return_err(i) = 0) Then
            'Cells(I+8, 3).Font.ColorIndex = 5
        Else
            'Cells(I+8, 3).Font.ColorIndex = 5
        End If
    Next
End Sub





Sub Play_nyx_1(frame)
    Dim sndbin(sendbyte) As Byte                         '送信バイトセット　6 + 5*26 +1 = 137
    Dim shtnum As String
    Dim wtim As Integer
    Shtname = ActiveSheet.Name
    posnum = frame + 8
'-------------------------------------
    'ec.COMn = comnum                                     'COMを開きます
    'ec.Setting = "115200,n,8,1"
'-----------------------
    sndbin(0) = &HFA
    sndbin(1) = &HAF
    sndbin(2) = &H0
    sndbin(3) = &H0
    sndbin(4) = &H1E
    sndbin(5) = &H5
    sndbin(6) = "&H" + Hex(servosend)
    
    For i = 2 To 6
    sndbin(sendbyte) = sndbin(sendbyte) Xor sndbin(i)
    Next
    
    STIM = Sheets(Shtname).Cells(5, posnum).Value
    
    TIMH = Int(STIM / 256)
    TIML = Int(STIM - 256 * TIMH)
    
    wtim = (Cells(5, posnum).Value * 10) + (Cells(6, posnum).Value * 10)
    If (wtim < 0) Then
        wtim = 0
    End If
    
    For i = cellstart To servosend + 7
        SID = Sheets(Shtname).Cells(i, 2).Value
        If Sheets(Shtname).Cells(i, 5).Value = "" Or Sheets(Shtname).Cells(i, 5).Value = 0 Then
            sdeg = Int((Sheets(Shtname).Cells(i, posnum).Value + Sheets(Shtname).Cells(i, 6).Value + Sheets(Shtname).Cells(i, 7).Value))
        Else
            sdeg = -1 * Int((Sheets(Shtname).Cells(i, posnum).Value + Sheets(Shtname).Cells(i, 6).Value + Sheets(Shtname).Cells(i, 7).Value))
        End If
        Dim deg_16bit As Long
        If (sdeg < 0) Then
            deg_16bit = (65536 - CInt(-1 * sdeg * 10))
            POSH = Int(deg_16bit / 256)
            POSL = Int(deg_16bit - 256 * POSH)
        Else
            deg_16bit = CInt(sdeg * 10)
            POSH = Int(deg_16bit / 256)
            POSL = Int(deg_16bit - 256 * POSH)
        End If
        'POSH = Int(sdeg / 256)
        'POSL = Int(sdeg - 256 * POSH)
        
        'If (sdeg < 0) Then
        '    POSH = Int(256 + POSH)
        'End If

        sndbin((i - 8) * 5 + 7) = "&H" + Hex(SID)
        sndbin(sendbyte) = sndbin(sendbyte) Xor sndbin((i - 8) * 5 + 7)
        sndbin((i - 8) * 5 + 8) = "&H" + Hex(POSL)
        sndbin(sendbyte) = sndbin(sendbyte) Xor sndbin((i - 8) * 5 + 8)
        sndbin((i - 8) * 5 + 9) = "&H" + Hex(POSH)
        sndbin(sendbyte) = sndbin(sendbyte) Xor sndbin((i - 8) * 5 + 9)
        sndbin((i - 8) * 5 + 10) = "&H" + Hex(TIML)
        sndbin(sendbyte) = sndbin(sendbyte) Xor sndbin((i - 8) * 5 + 10)
        sndbin((i - 8) * 5 + 11) = "&H" + Hex(TIMH)
        sndbin(sendbyte) = sndbin(sendbyte) Xor sndbin((i - 8) * 5 + 11)

    Next
    
    ec.Binary = sndbin()
    
'-----------------------------------

    QPC_wait_ms (wtim)

    
'-----------------------------------
    sndbin(sendbyte) = 0
    'ec.COMn = 0
    Exit Sub
End Sub


Sub ButtonPlayTest_2()
    Application.ScreenUpdating = False
    Call QPC_start_counting
    Call servo_on

    Dim sndbin(sendbyte) As Byte                         '送信バイトセット　6 + 5*26 +1 = 137
    Dim a1_pl, a2_pl, a3_pl, b_pl, c_pl, Lop As Integer
    Dim Shtname As String
    Shtname = ActiveSheet.Name
    Lop = Sheets(Shtname).Cells(2, 1).Value

    ec.COMn = comnum                                     'COMを開きます
    ec.Setting = "115200,n,8,1"
    QPC_wait_ms (10)

    For k = 9 To 28
        If Sheets(Shtname).Cells(7, k).Value = 1 Then
        a1_pl = k
        a2_pl = k
        k = 28
    End If
    
    Next
   
     For k = a1_pl + 1 To 28
        If Sheets(Shtname).Cells(7, k).Value = 2 Then
            a2_pl = k
            k = 28
        End If
    Next
   
    For k = (a1_pl - 8) To (a2_pl - 8)
        If Not Sheets(Shtname).Cells(7, k + 8).Value = 3 Then
            Play_nyx_1 (k)
        End If
        
        If Sheets(Shtname).Cells(1, k + 8) <> "" And Lop >= 1 And k > Sheets(Shtname).Cells(1, k + 8).Value And Sheets(Shtname).Cells(1, k + 8).Value >= 0 Then
            Lop = Lop - 1
            k = Sheets(Shtname).Cells(1, k + 8).Value - 1
        End If
    Next
    
   Play_nyx_1 (0)
    QPC_wait_ms (10)
   Application.ScreenUpdating = True
   ec.COMn = 0
End Sub





