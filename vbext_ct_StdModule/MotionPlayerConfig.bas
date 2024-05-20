Attribute VB_Name = "MotionPlayerConfig"
Function motionPlayerConfigGetNumMotors() As Long
    Dim sheet_name As String
    Dim num_motors As Long
    Dim i As Long
    Dim current_row As Long
    sheet_name = ActiveSheet.Name
    num_motors = 0
    For i = 0 To 254
        If (Sheets(sheet_name).Cells(i + 8, 2).Value > 0) Then
            num_motors = num_motors + 1
        Else
            Exit For
        End If
    Next
    motionPlayerConfigGetNumMotors = num_motors
End Function

Function motionPlayerConfigGetID(motor_num As Long) As Long
    Dim sheet_name As String
    Dim id As Long
    sheet_name = ActiveSheet.Name
    id = 0
    If (Sheets(sheet_name).Cells(motor_num + 8, 2).Value > 0) Then
        id = Sheets(sheet_name).Cells(motor_num + 8, 2).Value
    End If
    motionPlayerConfigGetID = id
End Function

Function motionPlayerConfigGetTargetPose(motor_num As Long, frame_num As Long) As Long
    Dim sheet_name As String
    Dim target_pose As Long
    sheet_name = ActiveSheet.Name
    target_pose = 0
    If (Sheets(sheet_name).Cells(motor_num + 8, frame_num + 8).Value <> 0) Then
        target_pose = Sheets(sheet_name).Cells(motor_num + 8, 2).Value
    End If
    motionPlayerConfigGetID = target_pose
End Function

Sub motionPlayerConfigSetCurrentPose(motor_num As Long, current_pose As Currency)
    Dim sheet_name As String
    sheet_name = ActiveSheet.Name
    Sheets(sheet_name).Cells(motor_num + 8, 2) = current_pose
End Sub
