Attribute VB_Name = "DynamixelMotionPlayer"


Sub dynamixelTorqueOn()
    Call comComfig
    Call comOpen
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOnPacket()
    ec.Binary = send_packet()
    'Test to read return packet
    qpcWaitMs (50)
    Call dynamixelUpdateRxBuffer
End Sub

Sub dynamixelTorqueOff()
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOffPacket()
    ec.Binary = send_packet()
    Call comClose
End Sub


