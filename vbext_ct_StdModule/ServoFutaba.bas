Attribute VB_Name = "ServoFutaba"
'Parameters
Public MOTOR_NUM As Integer

'Internal variables
Public TARGET_ID() As Integer
Public TARGET_POS() As Currency 'unit: [rad]
Public TARGET_VEL() As Currency 'unit: [rad/sec]
Public MEASURED_POS() As Currency 'unit: [rad]

Sub futabaSetMotorNum(ByVal input_num As Integer)
    MOTOR_NUM = input_num
    ReDim TARGET_ID(MOTOR_NUM)
    ReDim TARGET_POS(MOTOR_NUM)
    ReDim TARGET_VEL(MOTOR_NUM)
    ReDim MEASURED_POS(MOTOR_NUM)
    For i = 0 To MOTOR_NUM
        TARGET_ID(i) = 255
        TARGET_POS(i) = 0
        TARGET_VEL(i) = 0
        MEASURED_POS(i) = 0
    Next
End Sub

Function futabaTorqueOnPacket() As Byte()
    Dim send_packet(8)
    send_packet(0) = &HFA 'Header1
    send_packet(1) = &HAF 'Header2
    send_packet(2) = &HFF 'ID:all
    send_packet(3) = &H0
    send_packet(4) = &H24
    send_packet(5) = &H1
    send_packet(6) = &H1
    send_packet(7) = &H1
    send_packet(8) = 0 'Check Sum
    For i = 2 To 7
        send_packet(8) = send_packet(8) Xor send_packet(i)
    torqueOn = send_packet()
End Function

Function futabaTorqueOffPacket() As Byte()
    Dim send_packet(8)
    send_packet(0) = &HFA
    send_packet(1) = &HAF
    send_packet(2) = &HFF
    send_packet(3) = &H0
    send_packet(4) = &H24
    send_packet(5) = &H1
    send_packet(6) = &H1
    send_packet(7) = &H2
    send_packet(8) = 0 'Check Sum
    For i = 2 To 7
        send_packet(8) = send_packet(8) Xor send_packet(i)
    Next
    torqueOff = send_packet()
End Function

Function futabaChecksum(ByRef data_in() As Byte) As Byte
    Dim checksum As Byte
    checksum = 0
    For i = 2 To UBound(data_in) - 1
        checksum = checksum Xor data_in(i)
    Next
    futabaChecksum = checksum
End Function

