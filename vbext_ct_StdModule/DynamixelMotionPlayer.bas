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
    Call dynamixelDecodeRxBuffer
End Sub

Sub dynamixelTorqueOff()
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOffPacket()
    ec.Binary = send_packet()
    Call comClose
End Sub

Sub dynamixelSendPose()
    Dim buttonOrShapeName As String
    Dim result As String
    Dim buttom_row As Long
    Dim button_col As Long

    buttonOrShapeName = Application.Caller
    result = getButtonCenterCell(buttonOrShapeName)
    'MsgBox result
    
    button_row = GetRowFromResult(result)
    button_col = GetColumnFromResult(result)
    
    MsgBox "Row: " & button_row & ", Column: " & button_col
End Sub

