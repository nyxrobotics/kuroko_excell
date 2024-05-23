Attribute VB_Name = "ServoDynamixel"
'Reference
'Protocol 2.0: https://emanual.robotis.com/docs/en/dxl/protocol2/
'Control Table: https://www.besttechnology.co.jp/modules/knowledge/?X%20Series%20Control%20table
'CRC: https://emanual.robotis.com/docs/en/dxl/crc/

'Parameters
Public MOTOR_TOTAL As Long

'Internal variables
Public TARGET_ID() As Long
Public TARGET_POS() As Currency 'unit: [rad]
Public TARGET_VEL() As Currency 'unit: [rad/sec]
Public MEASURED_POS() As Currency 'unit: [rad]

'Received packet buffer
Public Const RX_RING_BUFFER_SIZE As Long = 1000
Public RX_RING_BUFFER(RX_RING_BUFFER_SIZE - 1) As Byte
Public RX_READ_POINT As Long
Public RX_WRITE_POINT As Long

Sub dynamixelSetMotorNum(ByVal input_num As Long)
    MOTOR_TOTAL = input_num
    ReDim TARGET_ID(MOTOR_TOTAL - 1)
    ReDim TARGET_POS(MOTOR_TOTAL - 1)
    ReDim TARGET_VEL(MOTOR_TOTAL - 1)
    ReDim MEASURED_POS(MOTOR_TOTAL - 1)
    For i = 0 To MOTOR_TOTAL - 1
        TARGET_ID(i) = 0
        TARGET_POS(i) = 0
        TARGET_VEL(i) = 0
        MEASURED_POS(i) = 0
    Next
End Sub

Sub dynamixelSetTargetID(ByVal input_num As Long, ByVal input_id As Long)
    If input_num < MOTOR_TOTAL Then
        TARGET_ID(input_num) = input_id
    End If
End Sub

Sub dynamixelSetTargetPos(ByVal input_num As Long, ByVal input_pos As Currency)
    If input_num < MOTOR_TOTAL Then
        TARGET_POS(input_num) = input_pos
    End If
End Sub

Sub dynamixelSetTargetVel(ByVal input_num As Long, ByVal input_vel As Currency)
    If input_num < MOTOR_TOTAL Then
        TARGET_VEL(input_num) = input_vel
    End If
End Sub

Function dynamixelTorqueOnPacket() As Byte()
    Dim send_packet(12) As Byte
    send_packet(0) = &HFF 'Header
    send_packet(1) = &HFF 'Header
    send_packet(2) = &HFD 'Header
    send_packet(3) = &H0  'Reserved
    send_packet(4) = &HFE 'ID (0xFE: Broadcast)
    send_packet(5) = &H6  'Length Low (Length = the number of Parameters + 3)
    send_packet(6) = &H0  'Length High
    send_packet(7) = &H3  'Instruction (3: Write)
    send_packet(8) = &H40 'Address Low (0x40: Torque Enable)
    send_packet(9) = &H0  'Address High
    send_packet(10) = &H1 'Parameters (1: ON)
    Dim crc As Long
    crc = dynamixelChecksum(send_packet)
    send_packet(11) = crc And &HFF&       'CRC Low
    send_packet(12) = (crc \ &H100&) And &HFF& 'CRC High
    dynamixelTorqueOnPacket = send_packet
End Function

Function dynamixelTorqueOffPacket() As Byte()
    Dim send_packet(12) As Byte
    send_packet(0) = &HFF 'Header
    send_packet(1) = &HFF 'Header
    send_packet(2) = &HFD 'Header
    send_packet(3) = &H0  'Reserved
    send_packet(4) = &HFE 'ID (0xFE: Broadcast)
    send_packet(5) = &H6  'Length Low
    send_packet(6) = &H0  'Length High
    send_packet(7) = &H3  'Instruction
    send_packet(8) = &H40 'Address Low (0x40: Torque Enable)
    send_packet(9) = &H0  'Address High
    send_packet(10) = &H0 'Parameters (0: OFF)
    Dim crc As Long
    crc = dynamixelChecksum(send_packet)
    send_packet(11) = crc And &HFF&       'CRC Low
    send_packet(12) = (crc \ &H100&) And &HFF& 'CRC High
    dynamixelTorqueOffPacket = send_packet
End Function

Function dynamixelRequestPosePacket(id As Long) As Byte()
    Dim send_packet(13) As Byte
    send_packet(0) = &HFF 'Header
    send_packet(1) = &HFF 'Header
    send_packet(2) = &HFD 'Header
    send_packet(3) = &H0  'Reserved
    send_packet(4) = id   'ID
    send_packet(5) = &H7  'Length Low
    send_packet(6) = &H0  'Length High
    send_packet(7) = &H2  'Instruction
    send_packet(8) = &H84 'Address Low (0x84: Read Position)
    send_packet(9) = &H0  'Address High
    send_packet(10) = &H4 'Length Low
    send_packet(11) = &H0 'Length High
    Dim crc As Long
    crc = dynamixelChecksum(send_packet)
    send_packet(12) = crc And &HFF&       'CRC Low
    send_packet(13) = (crc \ &H100&) And &HFF& 'CRC High
    dynamixelRequestPosePacket = send_packet
End Function

Function dynamixelSyncWritePosePacket() As Byte()
    Dim packet_length As Long
    Dim data_length As Long
    packet_length = 14 + MOTOR_TOTAL * 5
    data_length = 7 + MOTOR_TOTAL * 5
    Dim send_packet() As Byte
    ReDim send_packet(packet_length - 1)
    send_packet(0) = &HFF 'Header
    send_packet(1) = &HFF 'Header
    send_packet(2) = &HFD 'Header
    send_packet(3) = &H0  'Reserved
    send_packet(4) = &HFE 'ID (0xFE: Broadcast)
    send_packet(5) = data_length And &HFF&              'Length Low
    send_packet(6) = (data_length \ &H100&) And &HFF&   'Length High
    send_packet(7) = &H83 'Instruction
    send_packet(8) = &H74 'Address Low (0x74:Goal Position)
    send_packet(9) = &H0  'Address High
    send_packet(10) = &H4 'Length Low
    send_packet(11) = &H0 'Length High
    Dim i As Long
    Dim target_pose_float As Currency
    Dim target_pose_int As Long
    
    For i = 0 To MOTOR_TOTAL - 1
        target_pose_float = TARGET_POS(i)
        If target_pose_float < 0 Then
            target_pose_float = target_pose_float + 2 * WorksheetFunction.Pi()
        End If
        target_pose_int = Fix(target_pose_float * 2048 / WorksheetFunction.Pi())
        send_packet(i * 5 + 12) = TARGET_ID(i) 'ID
        send_packet(i * 5 + 13) = target_pose_int And &HFF&
        send_packet(i * 5 + 14) = (target_pose_int \ &H100&) And &HFF&
        send_packet(i * 5 + 15) = (target_pose_int \ &H10000) And &HFF&
        send_packet(i * 5 + 16) = (target_pose_int \ &H1000000) And &HFF&
    Next i
    Dim crc As Long
    crc = dynamixelChecksum(send_packet)
    send_packet(packet_length - 2) = crc And &HFF&   'CRC Low
    send_packet(packet_length - 1) = (crc \ &H100&) And &HFF& 'CRC High
    dynamixelSyncWritePosePacket = send_packet
End Function

Function dynamixelSyncRequestPosePacket() As Byte()
    Dim packet_length As Long
    Dim data_length As Long
    packet_length = 14 + MOTOR_TOTAL
    data_length = 7 + MOTOR_TOTAL
    Dim send_packet() As Byte
    ReDim send_packet(packet_length - 1)
    send_packet(0) = &HFF 'Header
    send_packet(1) = &HFF 'Header
    send_packet(2) = &HFD 'Header
    send_packet(3) = &H0  'Reserved
    send_packet(4) = &HFE 'ID (0xFE: Broadcast)
    send_packet(5) = data_length And &HFF&              'Length Low
    send_packet(6) = (data_length \ &H100&) And &HFF&   'Length High
    send_packet(7) = &H82 'Instruction
    send_packet(8) = &H84 'Address Low (0x84: Read Position)
    send_packet(9) = &H0  'Address High
    send_packet(10) = &H4 'Length Low
    send_packet(11) = &H0 'Length High
    Dim i As Long
    For i = 0 To MOTOR_TOTAL - 1
        send_packet(i + 12) = TARGET_ID(i) 'ID
    Next i
    Dim crc As Long
    crc = dynamixelChecksum(send_packet)
    send_packet(packet_length - 2) = crc And &HFF&   'CRC Low
    send_packet(packet_length - 1) = (crc \ &H100&) And &HFF& 'CRC High
    dynamixelSyncRequestPosePacket = send_packet
End Function

Sub dynamixelClearRxBuffer()
    RX_READ_POINT = 0
    RX_WRITE_POINT = 0
End Sub

Sub dynamixelUpdateRxBuffer()
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim received() As Byte
    received() = ec.Binary
    Dim received_size As Long
    received_size = UBound(received) + 1
    
    For i = 0 To received_size - 1
        RX_RING_BUFFER(RX_WRITE_POINT) = received(i)
        RX_WRITE_POINT = (RX_WRITE_POINT + 1) Mod RX_RING_BUFFER_SIZE
        If RX_WRITE_POINT = RX_READ_POINT Then
            RX_READ_POINT = (RX_READ_POINT + 1) Mod RX_RING_BUFFER_SIZE
        End If
    Next i
ErrorHandler:
        Exit Sub
End Sub

Function dynamixelGetRxDataSize() As Long
    Dim rx_data_size As Long
    rx_data_size = RX_WRITE_POINT - RX_READ_POINT
    If rx_data_size < 0 Then
        rx_data_size = rx_data_size + RX_RING_BUFFER_SIZE
    End If
    dynamixelGetRxDataSize = rx_data_size
End Function

Function dynamixelGetRxSingleByte(point As Long) As Byte
    Dim target_point As Long
    target_point = point Mod RX_RING_BUFFER_SIZE
    dynamixelGetRxSingleByte = RX_RING_BUFFER(target_point)
End Function

Function dynamixelRxSingleByteAvailable(point As Long) As Boolean
    Dim available As Boolean
    Dim target_point As Long
    available = False
    If RX_READ_POINT = RX_WRITE_POINT Then
        dynamixelRxSingleByteAvailable = False
        Exit Function
    End If
    target_point = point Mod RX_RING_BUFFER_SIZE
    If RX_READ_POINT < RX_WRITE_POINT Then
        If target_point < RX_WRITE_POINT Then
            available = True
        End If
    Else
        If target_point >= RX_READ_POINT Or target_point < RX_WRITE_POINT Then
            available = True
        End If
    End If
    dynamixelRxSingleByteAvailable = available
End Function

Sub dynamixelDecodeRxBuffer()
    ' Only read Position packet
    Dim i As Long, j As Long, k As Long
    Dim rx_data_size As Long
    Dim header_addr_candidate As Long
    Dim length_candidate As Long
    Dim footer_addr_candidate As Long
    Dim checksum_candidate As Long
    Dim checksum_expected As Long
    Dim single_status_packet() As Byte
    Dim status_id As Long
    Dim status_pose As Long
    Dim status_radian As Currency
    Const minimum_packet_size As Integer = 15
    
    rx_data_size = dynamixelGetRxDataSize()
    While (rx_data_size >= minimum_packet_size)
        For i = 0 To rx_data_size - 3
            ' Only if there are more than minimum_packet_size bytes including headers, go to the next process
            If rx_data_size - i < minimum_packet_size Then Exit Sub
            ' Find the 0xFF 0xFF 0xFD data and make it the header_candidate of the status packet
            If dynamixelGetRxSingleByte((RX_READ_POINT + i) Mod RX_RING_BUFFER_SIZE) = &HFF And _
               dynamixelGetRxSingleByte((RX_READ_POINT + i + 1) Mod RX_RING_BUFFER_SIZE) = &HFF And _
               dynamixelGetRxSingleByte((RX_READ_POINT + i + 2) Mod RX_RING_BUFFER_SIZE) = &HFD Then
                ' Check the length of the packet
                header_addr_candidate = (RX_READ_POINT + i) Mod RX_RING_BUFFER_SIZE
                length_candidate = dynamixelGetRxSingleByte((header_addr_candidate + 5) Mod RX_RING_BUFFER_SIZE) + _
                                   dynamixelGetRxSingleByte((header_addr_candidate + 6) Mod RX_RING_BUFFER_SIZE) * &H100
                footer_addr_candidate = (header_addr_candidate + 6 + length_candidate) Mod RX_RING_BUFFER_SIZE
                If dynamixelRxSingleByteAvailable(footer_addr_candidate) Then
                    ReDim single_status_packet(6 + length_candidate)
                    For j = 0 To 6 + length_candidate
                        single_status_packet(j) = dynamixelGetRxSingleByte((header_addr_candidate + j) Mod RX_RING_BUFFER_SIZE)
                    Next j
                    checksum_candidate = dynamixelGetRxSingleByte(footer_addr_candidate) * &H100& + _
                                         dynamixelGetRxSingleByte((footer_addr_candidate - 1 + RX_RING_BUFFER_SIZE) Mod RX_RING_BUFFER_SIZE)
                    checksum_expected = dynamixelChecksum(single_status_packet)
                    
                    If checksum_candidate = dynamixelChecksum(single_status_packet) Then
                        ' Received status packet
                        status_id = single_status_packet(4)
                        If status_id <> &HFD And length_candidate = 8 Then
                            status_pose = dynamixelGetRxSingleByte((header_addr_candidate + 9) Mod RX_RING_BUFFER_SIZE) + _
                                          dynamixelGetRxSingleByte((header_addr_candidate + 10) Mod RX_RING_BUFFER_SIZE) * &H100 + _
                                          dynamixelGetRxSingleByte((header_addr_candidate + 11) Mod RX_RING_BUFFER_SIZE) * &H10000 + _
                                          dynamixelGetRxSingleByte((header_addr_candidate + 12) Mod RX_RING_BUFFER_SIZE) * &H1000000
                            status_radian = status_pose * WorksheetFunction.Pi() / 2048
                            If status_radian > WorksheetFunction.Pi() Then
                                status_radian = status_radian - 2 * WorksheetFunction.Pi()
                            End If
                            For k = 0 To MOTOR_TOTAL - 1
                                If TARGET_ID(k) = status_id Then
                                    MEASURED_POS(k) = status_radian
                                    Exit For
                                End If
                            Next k
                        End If
                        RX_READ_POINT = (RX_READ_POINT + UBound(single_status_packet) + i + 1) Mod RX_RING_BUFFER_SIZE
                        Exit For
                    End If
                End If
            End If
        Next i
        rx_data_size = dynamixelGetRxDataSize()
    Wend
End Sub

Function dynamixelReadMeasuredPose(input_num As Long) As Currency
    dynamixelReadMeasuredPose = MEASURED_POS(input_num)
End Function

Function dynamixelChecksum(ByRef data_in() As Byte) As Long
    Dim crc As Long
    Dim crc_accum As Long
    Dim data_blk_size As Long
    crc_accum = 0
    data_blk_size = UBound(data_in) - 1
    crc = update_crc(crc_accum, data_in, data_blk_size)
    dynamixelChecksum = crc
End Function

'&H8000 is an Integer Constant, so the Sign bit is set making it -32768 and that Integer is copied to the Long value
'&H8000& is a Long Constant, so bit 15 of the Long is set making it 32768 and that Long is copied to the Long value
'Reference:  https://www.vbforums.com/showthread.php?847437-amp-H-values-and-Long-variable-type-in-Excel-VBA&p=5170541&viewfull=1#post5170541
Function update_crc(crc_accum As Long, data_blk_ptr() As Byte, data_blk_size As Long) As Long
    Dim i As Long, j As Long
    Dim crc_table(0 To 255) As Long
    crc_table(0) = &H0&: crc_table(1) = &H8005&: crc_table(2) = &H800F&: crc_table(3) = &HA&: crc_table(4) = &H801B&: crc_table(5) = &H1E&: crc_table(6) = &H14&: crc_table(7) = &H8011&
    crc_table(8) = &H8033&: crc_table(9) = &H36&: crc_table(10) = &H3C&: crc_table(11) = &H8039&: crc_table(12) = &H28&: crc_table(13) = &H802D&: crc_table(14) = &H8027&: crc_table(15) = &H22&
    crc_table(16) = &H8063&: crc_table(17) = &H66&: crc_table(18) = &H6C&: crc_table(19) = &H8069&: crc_table(20) = &H78&: crc_table(21) = &H807D&: crc_table(22) = &H8077&: crc_table(23) = &H72&
    crc_table(24) = &H50&: crc_table(25) = &H8055&: crc_table(26) = &H805F&: crc_table(27) = &H5A&: crc_table(28) = &H804B&: crc_table(29) = &H4E&: crc_table(30) = &H44&: crc_table(31) = &H8041&
    crc_table(32) = &H80C3&: crc_table(33) = &HC6&: crc_table(34) = &HCC&: crc_table(35) = &H80C9&: crc_table(36) = &HD8&: crc_table(37) = &H80DD&: crc_table(38) = &H80D7&: crc_table(39) = &HD2&
    crc_table(40) = &HF0&: crc_table(41) = &H80F5&: crc_table(42) = &H80FF&: crc_table(43) = &HFA&: crc_table(44) = &H80EB&: crc_table(45) = &HEE&: crc_table(46) = &HE4&: crc_table(47) = &H80E1&
    crc_table(48) = &HA0&: crc_table(49) = &H80A5&: crc_table(50) = &H80AF&: crc_table(51) = &HAA&: crc_table(52) = &H80BB&: crc_table(53) = &HBE&: crc_table(54) = &HB4&: crc_table(55) = &H80B1&
    crc_table(56) = &H8093&: crc_table(57) = &H96&: crc_table(58) = &H9C&: crc_table(59) = &H8099&: crc_table(60) = &H88&: crc_table(61) = &H808D&: crc_table(62) = &H8087&: crc_table(63) = &H82&
    crc_table(64) = &H8183&: crc_table(65) = &H186&: crc_table(66) = &H18C&: crc_table(67) = &H8189&: crc_table(68) = &H198&: crc_table(69) = &H819D&: crc_table(70) = &H8197&: crc_table(71) = &H192&
    crc_table(72) = &H1B0&: crc_table(73) = &H81B5&: crc_table(74) = &H81BF&: crc_table(75) = &H1BA&: crc_table(76) = &H81AB&: crc_table(77) = &H1AE&: crc_table(78) = &H1A4&: crc_table(79) = &H81A1&
    crc_table(80) = &H1E0&: crc_table(81) = &H81E5&: crc_table(82) = &H81EF&: crc_table(83) = &H1EA&: crc_table(84) = &H81FB&: crc_table(85) = &H1FE&: crc_table(86) = &H1F4&: crc_table(87) = &H81F1&
    crc_table(88) = &H81D3&: crc_table(89) = &H1D6&: crc_table(90) = &H1DC&: crc_table(91) = &H81D9&: crc_table(92) = &H1C8&: crc_table(93) = &H81CD&: crc_table(94) = &H81C7&: crc_table(95) = &H1C2&
    crc_table(96) = &H140&: crc_table(97) = &H8145&: crc_table(98) = &H814F&: crc_table(99) = &H14A&: crc_table(100) = &H815B&: crc_table(101) = &H15E&: crc_table(102) = &H154&: crc_table(103) = &H8151&
    crc_table(104) = &H8173&: crc_table(105) = &H176&: crc_table(106) = &H17C&: crc_table(107) = &H8179&: crc_table(108) = &H168&: crc_table(109) = &H816D&: crc_table(110) = &H8167&: crc_table(111) = &H162&
    crc_table(112) = &H8123&: crc_table(113) = &H126&: crc_table(114) = &H12C&: crc_table(115) = &H8129&: crc_table(116) = &H138&: crc_table(117) = &H813D&: crc_table(118) = &H8137&: crc_table(119) = &H132&
    crc_table(120) = &H110&: crc_table(121) = &H8115&: crc_table(122) = &H811F&: crc_table(123) = &H11A&: crc_table(124) = &H810B&: crc_table(125) = &H10E&: crc_table(126) = &H104&: crc_table(127) = &H8101&
    crc_table(128) = &H8303&: crc_table(129) = &H306&: crc_table(130) = &H30C&: crc_table(131) = &H8309&: crc_table(132) = &H318&: crc_table(133) = &H831D&: crc_table(134) = &H8317&: crc_table(135) = &H312&
    crc_table(136) = &H330&: crc_table(137) = &H8335&: crc_table(138) = &H833F&: crc_table(139) = &H33A: crc_table(140) = &H832B&: crc_table(141) = &H32E&: crc_table(142) = &H324&: crc_table(143) = &H8321&
    crc_table(144) = &H360&: crc_table(145) = &H8365&: crc_table(146) = &H836F&: crc_table(147) = &H36A: crc_table(148) = &H837B&: crc_table(149) = &H37E&: crc_table(150) = &H374&: crc_table(151) = &H8371&
    crc_table(152) = &H8353&: crc_table(153) = &H356&: crc_table(154) = &H35C&: crc_table(155) = &H8359: crc_table(156) = &H348&: crc_table(157) = &H834D&: crc_table(158) = &H8347&: crc_table(159) = &H342&
    crc_table(160) = &H3C0&: crc_table(161) = &H83C5&: crc_table(162) = &H83CF&: crc_table(163) = &H3CA: crc_table(164) = &H83DB&: crc_table(165) = &H3DE&: crc_table(166) = &H3D4&: crc_table(167) = &H83D1&
    crc_table(168) = &H83F3&: crc_table(169) = &H3F6&: crc_table(170) = &H3FC&: crc_table(171) = &H83F9: crc_table(172) = &H3E8&: crc_table(173) = &H83ED&: crc_table(174) = &H83E7&: crc_table(175) = &H3E2&
    crc_table(176) = &H83A3&: crc_table(177) = &H3A6&: crc_table(178) = &H3AC&: crc_table(179) = &H83A9: crc_table(180) = &H3B8&: crc_table(181) = &H83BD&: crc_table(182) = &H83B7&: crc_table(183) = &H3B2&
    crc_table(184) = &H390&: crc_table(185) = &H8395&: crc_table(186) = &H839F&: crc_table(187) = &H39A: crc_table(188) = &H838B&: crc_table(189) = &H38E&: crc_table(190) = &H384&: crc_table(191) = &H8381&
    crc_table(192) = &H280&: crc_table(193) = &H8285&: crc_table(194) = &H828F&: crc_table(195) = &H28A: crc_table(196) = &H829B&: crc_table(197) = &H29E&: crc_table(198) = &H294&: crc_table(199) = &H8291&
    crc_table(200) = &H82B3&: crc_table(201) = &H2B6&: crc_table(202) = &H2BC&: crc_table(203) = &H82B9: crc_table(204) = &H2A8&: crc_table(205) = &H82AD&: crc_table(206) = &H82A7&: crc_table(207) = &H2A2&
    crc_table(208) = &H82E3&: crc_table(209) = &H2E6&: crc_table(210) = &H2EC&: crc_table(211) = &H82E9: crc_table(212) = &H2F8&: crc_table(213) = &H82FD&: crc_table(214) = &H82F7&: crc_table(215) = &H2F2&
    crc_table(216) = &H2D0&: crc_table(217) = &H82D5&: crc_table(218) = &H82DF&: crc_table(219) = &H2DA: crc_table(220) = &H82CB&: crc_table(221) = &H2CE&: crc_table(222) = &H2C4&: crc_table(223) = &H82C1&
    crc_table(224) = &H8243&: crc_table(225) = &H246&: crc_table(226) = &H24C&: crc_table(227) = &H8249: crc_table(228) = &H258&: crc_table(229) = &H825D&: crc_table(230) = &H8257&: crc_table(231) = &H252&
    crc_table(232) = &H270&: crc_table(233) = &H8275&: crc_table(234) = &H827F&: crc_table(235) = &H27A: crc_table(236) = &H826B&: crc_table(237) = &H26E&: crc_table(238) = &H264&: crc_table(239) = &H8261&
    crc_table(240) = &H220&: crc_table(241) = &H8225&: crc_table(242) = &H822F&: crc_table(243) = &H22A: crc_table(244) = &H823B&: crc_table(245) = &H23E&: crc_table(246) = &H234&: crc_table(247) = &H8231&
    crc_table(248) = &H8213&: crc_table(249) = &H216&: crc_table(250) = &H21C&: crc_table(251) = &H8219: crc_table(252) = &H208&: crc_table(253) = &H820D&: crc_table(254) = &H8207&: crc_table(255) = &H202&
    For j = 0 To data_blk_size - 1
        i = ((crc_accum \ &H100&) Xor data_blk_ptr(j)) And &HFF&
        crc_accum = ((crc_accum And &HFF&) * &H100&) Xor crc_table(i)
        crc_accum = crc_accum And &HFFFF&
    Next j
    update_crc = crc_accum
End Function

