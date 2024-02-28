Attribute VB_Name = "nyx_orientalmotor"
Sub oriental_test()

    Dim baud_rate As Long
    Dim com_port As Integer
    Dim sending_packet_2(15) As Byte
    Dim Shtname As String
    Dim tmp As Integer
    Dim tmp_H As Integer
    Dim tmp_L As Integer
    Dim ec_set As String
    
    Application.ScreenUpdating = False
    Call QPC_start_counting
            
    Shtname = ActiveSheet.Name
    baud_rate = Sheets("COM").Cells(2, 2).Value
    com_port = Sheets("COM").Cells(1, 2).Value

    ec.COMn = com_port 'COM‚ðŠJ‚«‚Ü‚·
    ec_set = SPrintF("%d,e,8,1", baud_rate)
    ec.Setting = ec_set '"500000,n,8,1"

    Dim k As Integer
    For k = 1 To 1000
        If (k Mod 2) = 0 Then
            sending_packet_2(0) = &H0
            sending_packet_2(1) = &H10
            sending_packet_2(2) = &H0
            sending_packet_2(3) = &H5C   'Packet Length
            sending_packet_2(4) = &H0   'instruction"
            sending_packet_2(5) = &H2   'Adress of Torque on/off
            sending_packet_2(6) = &H4   'data 1
            sending_packet_2(7) = &H0   'data 2
            sending_packet_2(8) = &H0   'data 2
            sending_packet_2(9) = &H0   'data 2
            sending_packet_2(10) = &H64   'data 2
            sending_packet_2(11) = &HF3   'data 2
            sending_packet_2(12) = &HD1   'data 2
        Else
            sending_packet_2(0) = &H0
            sending_packet_2(1) = &H10
            sending_packet_2(2) = &H0
            sending_packet_2(3) = &H5C   'Packet Length
            sending_packet_2(4) = &H0   'instruction"
            sending_packet_2(5) = &H2   'Adress of Torque on/off
            sending_packet_2(6) = &H4   'data 1
            sending_packet_2(7) = &H0   'data 2
            sending_packet_2(8) = &H0   'data 2
            sending_packet_2(9) = &H0   'data 2
            sending_packet_2(10) = &HC8   'data 2
            sending_packet_2(11) = &HF3   'data 2
            sending_packet_2(12) = &HAC   'data 2
        End If
        ec.Binary = sending_packet_2()
        QPC_wait_ms (5)
    Next

    
    
    ec.COMn = 0
    Application.ScreenUpdating = True
End Sub
