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

Function getButtonCenterCell(buttonOrShapeName As String) As String
    Dim button As Object
    Dim shape As shape
    Dim result As String

    ' Try to get the button (Form Control Button or ActiveX Control Button)
    On Error Resume Next
    Set button = ActiveSheet.Buttons(buttonOrShapeName)
    If button Is Nothing Then
        Set button = ActiveSheet.OLEObjects(buttonOrShapeName).Object
    End If
    On Error GoTo 0

    ' If button is found, use GetControlButtonCenterCell
    If Not button Is Nothing Then
        result = getControlButtonCenterCell(button)
    Else
        ' If button is not found, try to get the shape
        On Error Resume Next
        Set shape = ActiveSheet.Shapes(buttonOrShapeName)
        On Error GoTo 0

        If Not shape Is Nothing Then
            result = getShapeButtonCenterCell(shape)
        Else
            result = "Button or shape not found."
        End If
    End If

    getButtonCenterCell = result
End Function

Function getControlButtonCenterCell(button As Object) As String
    Dim topLeftCell As Range
    Dim bottomRightCell As Range
    Dim centerRow As Long
    Dim centerColumn As Long
    Dim buttonCenterY As Double
    Dim buttonCenterX As Double

    ' Determine the type of the button
    If TypeOf button Is MSForms.CommandButton Or TypeOf button Is OLEObject Then
        ' For ActiveX controls
        With button
            Set topLeftCell = .topLeftCell
            Set bottomRightCell = .bottomRightCell
        End With
    ElseIf TypeOf button Is button Then
        ' For Form control buttons
        With button
            Set topLeftCell = .topLeftCell
            Set bottomRightCell = .bottomRightCell
        End With
    Else
        getControlButtonCenterCell = "The button type is not supported."
        Exit Function
    End If

    ' Calculate the exact center of the button
    buttonCenterY = (topLeftCell.Top + bottomRightCell.Top + bottomRightCell.Height) / 2
    buttonCenterX = (topLeftCell.Left + bottomRightCell.Left + bottomRightCell.Width) / 2

    ' Find the cell that contains the center point
    centerRow = Cells(1, 1).row
    centerColumn = Cells(1, 1).column

    Do While buttonCenterY > Cells(centerRow, 1).Top + Cells(centerRow, 1).Height
        centerRow = centerRow + 1
    Loop

    Do While buttonCenterX > Cells(1, centerColumn).Left + Cells(1, centerColumn).Width
        centerColumn = centerColumn + 1
    Loop

    ' Return the result
    getControlButtonCenterCell = "Row: " & centerRow & ", Column: " & centerColumn
End Function

Function getShapeButtonCenterCell(shape As shape) As String
    Dim ws As Worksheet
    Dim shapeCenterX As Double
    Dim shapeCenterY As Double
    Dim centerRow As Long
    Dim centerColumn As Long

    ' Get the worksheet containing the shape
    Set ws = shape.Parent

    ' Calculate the exact center of the shape
    shapeCenterX = shape.Left + (shape.Width / 2)
    shapeCenterY = shape.Top + (shape.Height / 2)

    ' Find the cell that contains the center point
    centerRow = ws.Cells(1, 1).row
    centerColumn = ws.Cells(1, 1).column

    Do While shapeCenterY > ws.Cells(centerRow, 1).Top + ws.Cells(centerRow, 1).Height
        centerRow = centerRow + 1
    Loop

    Do While shapeCenterX > ws.Cells(1, centerColumn).Left + ws.Cells(1, centerColumn).Width
        centerColumn = centerColumn + 1
    Loop

    ' Return the result
    getShapeButtonCenterCell = "Row: " & centerRow & ", Column: " & centerColumn
End Function

Function GetRowFromResult(result As String) As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim rowString As String
    startPos = InStr(result, "Row: ") + Len("Row: ")
    endPos = InStr(result, ", Column")
    rowString = Mid(result, startPos, endPos - startPos)
    GetRowFromResult = CLng(rowString)
End Function

Function GetColumnFromResult(result As String) As Long
    Dim startPos As Long
    Dim colString As String
    startPos = InStr(result, "Column: ") + Len("Column: ")
    colString = Mid(result, startPos)
    GetColumnFromResult = CLng(colString)
End Function
