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

Sub dynamixelPlayButton()
    Dim i As Long
    Dim j As Long
    Dim sheet_name As String
    Dim motion_start_frame As Long
    Dim loop_start_frame As Long
    Dim loop_end_frame As Long
    Dim motion_end_frame As Long
    Dim loop_count As Long
    motion_start_frame = 0
    loop_start_frame = 0
    loop_end_frame = 0
    motion_end_frame = 0
    sheet_name = ActiveSheet.Name
    loop_count = Sheets(sheet_name).Cells(2, 1).Value
    For i = 0 To 100
        If 1 = Sheets(sheet_name).Cells(7, i + 8).Value Then
            motion_start_frame = i
            Exit For
        End If
    Next i
    For i = 0 To 100
        If "" <> Sheets(sheet_name).Cells(1, i + 8).Value Then
            loop_end_frame = i
            loop_start_frame = Sheets(sheet_name).Cells(1, i + 8).Value
            Exit For
        End If
    Next i
    For i = 0 To 100
        If 2 = Sheets(sheet_name).Cells(7, i + 8).Value Then
            motion_end_frame = i
            Exit For
        End If
    Next i
    'Play motion
    Call dynamixelSendPoseWithInterval(motion_start_frame, 0.5)
    Call qpcWait(1#)
    If loop_end_frame <> 0 Then
        For i = motion_start_frame + 1 To loop_end_frame
            If 0 = Sheets(sheet_name).Cells(7, i + 8).Value Then
                Call dynamixelSendAnimationFrame(i - 1, i)
            End If
        Next i
        For j = 1 To loop_count
            Call dynamixelSendAnimationFrame(loop_end_frame, loop_start_frame)
            For i = loop_start_frame + 1 To loop_end_frame
                If 0 = Sheets(sheet_name).Cells(7, i + 8).Value Then
                    Call dynamixelSendAnimationFrame(i - 1, i)
                End If
            Next i
        Next j
        For i = loop_end_frame + 1 To motion_end_frame
            If 0 = Sheets(sheet_name).Cells(7, i + 8).Value Or 2 = Sheets(sheet_name).Cells(7, i + 8).Value Then
                Call dynamixelSendAnimationFrame(i - 1, i)
            End If
        Next i
    ElseIf motion_end_frame > motion_start_frame Then
        For i = motion_start_frame + 1 To motion_end_frame
            If 0 = Sheets(sheet_name).Cells(7, i + 8).Value Or 2 = Sheets(sheet_name).Cells(7, i + 8).Value Then
                Call dynamixelSendAnimationFrame(i - 1, i)
            End If
        Next i
    End If
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
        target_pos_rad = deg2Rad(Sheets(sheet_name).Cells(i + 8, 6).Value)                                      'Home
        target_pos_rad = target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, 7).Value)                     'Offset
        target_pos_rad = target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, input_frame_num + 8).Value)   'Target
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
        target_pos_rad = deg2Rad(Sheets(sheet_name).Cells(i + 8, 6).Value)                                      'Home
        target_pos_rad = target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, 7).Value)                     'Offset
        target_pos_rad = target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, input_frame_num + 8).Value)   'Target
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
        target_pos_rad = deg2Rad(Sheets(sheet_name).Cells(i + 8, 6).Value)                                      'Home
        target_pos_rad = target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, 7).Value)                     'Offset
        target_pos_rad = target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, input_frame_num + 8).Value)   'Target
        If is_reverse Then target_pos_rad = -target_pos_rad
        Call dynamixelSetTargetPos(i, target_pos_rad)
        current_pos_rad = dynamixelReadMeasuredPose(i)
        target_vel_rad_per_sec = (target_pos_rad - current_pos_rad) / input_interval
        Call dynamixelSetTargetVel(i, target_vel_rad_per_sec)
    Next i
    send_packet() = dynamixelSyncWritePosVelPacket()
    ec.Binary = send_packet()
End Sub

'Available only in frame 1 and above
'Calculate the amount of travel based on the difference from one previous frame
'Calculate the speed required to arrive on time
Sub dynamixelSendAnimationFrame(ByVal input_previous_frame_num As Long, ByVal input_frame_num As Long)
    Dim i As Long
    Dim sheet_name As String
    Dim num_motors As Long
    Dim target_pos_rad As Currency
    Dim target_vel_rad_per_sec As Currency
    Dim previous_target_pos_rad As Currency
    Dim travel_interval As Currency
    Dim wait_interval As Currency
    Dim is_reverse As Long
    Dim send_packet() As Byte
    Call qpcInit
    Call comComfig
    Call comOpen
    sheet_name = ActiveSheet.Name
    num_motors = motionPlayerConfigGetNumMotors()
    Call dynamixelSetMotorNum(num_motors)
    travel_interval = Sheets(sheet_name).Cells(5, input_frame_num + 8).Value * 0.01
    wait_interval = travel_interval + (Sheets(sheet_name).Cells(6, input_frame_num + 8).Value * 0.01)
    For i = 0 To num_motors - 1
        id = motionPlayerConfigGetID(i)
        Call dynamixelSetTargetID(i, id)
        is_reverse = Sheets(sheet_name).Cells(i + 8, 5).Value
        previous_target_pos_rad = deg2Rad(Sheets(sheet_name).Cells(i + 8, 6).Value)                                                         'Home
        previous_target_pos_rad = previous_target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, 7).Value)                               'Offset
        previous_target_pos_rad = previous_target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, input_previous_frame_num + 8).Value)    'Target
        target_pos_rad = deg2Rad(Sheets(sheet_name).Cells(i + 8, 6).Value)                                      'Home
        target_pos_rad = target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, 7).Value)                     'Offset
        target_pos_rad = target_pos_rad + deg2Rad(Sheets(sheet_name).Cells(i + 8, input_frame_num + 8).Value)   'Target
        If is_reverse Then
            target_pos_rad = -target_pos_rad
            previous_target_pos_rad = -previous_target_pos_rad
        End If
        Call dynamixelSetTargetPos(i, target_pos_rad)
        target_vel_rad_per_sec = (target_pos_rad - current_pos_rad) / travel_interval
        Call dynamixelSetTargetVel(i, target_vel_rad_per_sec)
    Next i
    send_packet() = dynamixelSyncWritePosVelPacket()
    ec.Binary = send_packet()
    Call qpcWait(wait_interval)
End Sub

