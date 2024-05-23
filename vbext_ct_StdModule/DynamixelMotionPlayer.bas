Attribute VB_Name = "DynamixelMotionPlayer"
Sub dynamixelTorqueOn()
    Call comComfig
    Call comOpen
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOnPacket()
    ec.Binary = send_packet()
End Sub

Sub dynamixelTorqueOff()
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOffPacket()
    ec.Binary = send_packet()
    Call comClose
End Sub

Sub dynamixelGetPose()
    Call dynamixelClearRxBuffer
    Call qpcInit
    Call comComfig
    Call comOpen
    Dim sheet_name As String
    Dim num_motors As Long
    Dim id As Long
    Dim current_pose As Currency
    Dim send_packet() As Byte
    Dim i As Long
    sheet_name = ActiveSheet.Name
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelSetMotorNum(num_motors)
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
    Next i
    send_packet() = dynamixelSyncRequestPosePacket()
    ec.Binary = send_packet()
    Call sleepSend(send_packet)
    Call sleepReceive(15 * num_motors)
    Call qpcWaitMs(15 + num_motors)
    Call dynamixelUpdateRxBuffer
    Call dynamixelDecodeRxBuffer
    For i = 0 To num_motors - 1
        Sheets(sheet_name).Cells(i + 8, 3) = dynamixelReadMeasuredPose(i)
    Next i
End Sub

Sub dynamixelGetPoseEach()
    Call dynamixelClearRxBuffer
    Call qpcInit
    Call comComfig
    Call comOpen
    Dim sheet_name As String
    Dim num_motors As Long
    Dim id As Long
    Dim current_pose As Currency
    Dim send_packet() As Byte
    Dim i As Long
    sheet_name = ActiveSheet.Name
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelSetMotorNum(num_motors)
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        send_packet() = dynamixelRequestPosePacket(id)
        ec.Binary = send_packet()
        Call sleepSend(send_packet)
        Call sleepReceive(15)
        Call qpcWaitMs(15)
        Call dynamixelUpdateRxBuffer
        Call dynamixelDecodeRxBuffer
        Sheets(sheet_name).Cells(i + 8, 3) = dynamixelReadMeasuredPose(i)
    Next i
End Sub

Sub dynamixelSendPose()
    Dim buttonOrShapeName As String
    Dim result As String
    Dim buttom_row As Long
    Dim button_col As Long
    'Get the pose of the button
    buttonOrShapeName = Application.Caller
    result = getButtonCenterCell(buttonOrShapeName)
    button_row = GetRowFromResult(result)
    button_col = GetColumnFromResult(result)
    'MsgBox "Row: " & button_row & ", Column: " & button_col
    Dim i As Long
    Dim num_motors As Long
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelSetMotorNum(num_motors)
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        Call dynamixelSetTargetPos(i, 0)
    Next i
    Dim send_packet() As Byte
    send_packet() = dynamixelSyncWritePosePacket()
    Call qpcInit
    Call comComfig
    Call comOpen
    ec.Binary = send_packet()
End Sub
