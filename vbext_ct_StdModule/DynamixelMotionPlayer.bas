Attribute VB_Name = "DynamixelMotionPlayer"
Sub dynamixelTorqueOn()
    Call comComfig
    Call comOpen
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOnPacket()
    ec.Binary = send_packet()
End Sub

Sub dynamixelTorqueOff()
    Call comComfig
    Call comOpen
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOffPacket()
    ec.Binary = send_packet()
End Sub

Sub dynamixelClose()
    Call comClose
End Sub

Sub dynamixelGetPose()
    Dim i As Long
    Dim sheet_name As String
    Dim num_motors As Long
    Dim id As Long
    Dim current_pos_deg As Currency
    Dim is_updated As Boolean
    Dim is_reverse As Long
    Dim send_packet() As Byte
    Call dynamixelClearRxBuffer
    Call qpcInit
    Call comComfig
    Call comOpen
    sheet_name = ActiveSheet.Name
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelSetMotorNum(num_motors)
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        dynamixelClearMeasuredPoseIsUpdated (i)
    Next i
    send_packet() = dynamixelSyncRequestPosePacket()
    ec.Binary = send_packet()
    Call sleepSend(send_packet)
    Call sleepReceive(15 * num_motors)
    Call qpcWaitMs(15 + num_motors)
    Call dynamixelUpdateRxBuffer
    Call dynamixelDecodeRxBuffer
    For i = 0 To num_motors - 1
        current_pos_deg = rad2Deg(dynamixelReadMeasuredPose(i))
        is_reverse = Sheets(sheet_name).Cells(i + 8, 5).Value
        If is_reverse Then current_pos_deg = -current_pos_deg
        Sheets(sheet_name).Cells(i + 8, 3) = current_pos_deg
        is_updated = dynamixelMeasuredPoseIsUpdated(i)
        If is_updated Then
            Sheets(sheet_name).Cells(i + 8, 3).Font.Color = RGB(0, 0, 255)
        Else
            Sheets(sheet_name).Cells(i + 8, 3).Font.Color = RGB(255, 0, 0)
        End If
    Next i
End Sub

Sub dynamixelGetPoseEach()
    Dim i As Long
    Dim sheet_name As String
    Dim num_motors As Long
    Dim id As Long
    Dim current_pos_deg As Currency
    Dim is_updated As Boolean
    Dim is_reverse As Long
    Dim send_packet() As Byte
    sheet_name = ActiveSheet.Name
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelClearRxBuffer
    Call qpcInit
    Call comComfig
    Call comOpen
    Call dynamixelSetMotorNum(num_motors)
    For i = 0 To num_motors - 1
        dynamixelClearMeasuredPoseIsUpdated (i)
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        send_packet() = dynamixelRequestPosePacket(id)
        ec.Binary = send_packet()
        Call sleepSend(send_packet)
        Call sleepReceive(15)
        Call qpcWaitMs(15)
        Call dynamixelUpdateRxBuffer
        Call dynamixelDecodeRxBuffer
        current_pos_deg = rad2Deg(dynamixelReadMeasuredPose(i))
        is_reverse = Sheets(sheet_name).Cells(i + 8, 5).Value
        If is_reverse Then current_pos_deg = -current_pos_deg
        Sheets(sheet_name).Cells(i + 8, 3) = current_pos_deg
        is_updated = dynamixelMeasuredPoseIsUpdated(i)
        If is_updated Then
            Sheets(sheet_name).Cells(i + 8, 3).Font.Color = RGB(0, 0, 255)
        Else
            Sheets(sheet_name).Cells(i + 8, 3).Font.Color = RGB(255, 0, 0)
        End If
    Next i
End Sub

Sub dynamixelPoseButton()
    Dim button_result As String
    Dim button_name As String
    Dim buttom_row As Long
    Dim button_col As Long
    Dim animation_frame_num As Long
    'Get the pose of the button
    button_name = Application.Caller
    button_result = getButtonCenterCell(button_name)
    button_row = GetRowFromResult(button_result)
    button_col = GetColumnFromResult(button_result)
    animation_frame_num = button_col - 8
    Call dynamixelSendPoseWithInterval(animation_frame_num, 0.5)
End Sub

Sub dynamixelSendPose(ByVal input_frame_num As Long)
    Dim i As Long
    Dim sheet_name As String
    Dim num_motors As Long
    Dim target_pos_rad As Currency
    Dim is_reverse As Long
    Dim send_packet() As Byte
    Call qpcInit
    Call comComfig
    Call comOpen
    sheet_name = ActiveSheet.Name
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelSetMotorNum(num_motors)
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        is_reverse = Sheets(sheet_name).Cells(i + 8, 5).Value
        target_pos_rad = deg2Rad(Sheets(sheet_name).Cells(i + 8, input_frame_num + 8).Value)
        If is_reverse Then target_pos_rad = -target_pos_rad
        Call dynamixelSetTargetPos(i, target_pos_rad)
    Next i
    send_packet() = dynamixelSyncWritePosPacket()
    ec.Binary = send_packet()
End Sub

Sub dynamixelSendPoseWithVel(ByVal input_frame_num As Long, ByVal input_vel As Currency)
    Dim i As Long
    Dim sheet_name As String
    Dim num_motors As Long
    Dim target_pos_rad As Currency
    Dim is_reverse As Long
    Dim send_packet() As Byte
    Call qpcInit
    Call comComfig
    Call comOpen
    sheet_name = ActiveSheet.Name
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelSetMotorNum(num_motors)
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        is_reverse = Sheets(sheet_name).Cells(i + 8, 5).Value
        target_pos_rad = deg2Rad(Sheets(sheet_name).Cells(i + 8, input_frame_num + 8).Value)
        If is_reverse Then target_pos_rad = -target_pos_rad
        Call dynamixelSetTargetPos(i, target_pos_rad)
        Call dynamixelSetTargetVel(i, input_vel)
    Next i
    send_packet() = dynamixelSyncWritePosVelPacket()
    ec.Binary = send_packet()
End Sub

Sub dynamixelSendPoseWithInterval(ByVal input_frame_num As Long, ByVal input_interval As Currency)
    Dim i As Long
    Dim sheet_name As String
    Dim num_motors As Long
    Dim current_pos_rad As Currency
    Dim target_pos_rad As Currency
    Dim target_vel_rad_per_sec As Currency
    Dim is_reverse As Long
    Dim send_packet() As Byte
    Call qpcInit
    Call comComfig
    Call comOpen
    sheet_name = ActiveSheet.Name
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelSetMotorNum(num_motors)
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        dynamixelClearMeasuredPoseIsUpdated (i)
    Next i
    send_packet() = dynamixelSyncRequestPosePacket()
    ec.Binary = send_packet()
    Call sleepSend(send_packet)
    Call sleepReceive(15 * num_motors)
    Call qpcWaitMs(15 + num_motors)
    Call dynamixelUpdateRxBuffer
    Call dynamixelDecodeRxBuffer
    
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        is_reverse = Sheets(sheet_name).Cells(i + 8, 5).Value
        target_pos_rad = deg2Rad(Sheets(sheet_name).Cells(i + 8, input_frame_num + 8).Value)
        If is_reverse Then target_pos_rad = -target_pos_rad
        Call dynamixelSetTargetPos(i, target_pos_rad)
        current_pos_rad = dynamixelReadMeasuredPose(i)
        target_vel_rad_per_sec = (target_pos_rad - current_pos_rad) / input_interval
        Call dynamixelSetTargetVel(i, target_vel_rad_per_sec)
    Next i
    send_packet() = dynamixelSyncWritePosVelPacket()
    ec.Binary = send_packet()
End Sub


Sub dynamixelSendAnimationFrame(ByVal input_frame_num As Long)
End Sub

