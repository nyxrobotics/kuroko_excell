Attribute VB_Name = "DynamixelMotionPlayer"

Public COM_NUM As Integer
Public COM_BAUD As Long
Public COM_PARITY As String
Public COM_LENGTH As Integer
Public COM_STOP As Integer

Sub comComfig()
    COM_NUM = 8
    COM_BAUD = 57600
    COM_PARITY = "n"
    COM_LENGTH = 8
    COM_STOP = 1
End Sub

Sub comOpen()
    Call qpcInit
    Dim ec_config As String
    ec_config = SPrintF("%d,%s,%d,%d", COM_BAUD, COM_PARITY, COM_LENGTH, COM_STOP)
    ec.COMn = COM_NUM 'COMを開きます
    ec.Setting = ec_config
    qpcWaitMs (20)
End Sub

Sub comClose()
    ec.COMn = 0
    Application.ScreenUpdating = False
End Sub

Sub dynamixelTorqueOn()
    Call comComfig
    Call comOpen
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOnPacket()
    ec.Binary = send_packet()
    qpcWaitMs (50)
    Call dynamixelUpdateRxBuffer
End Sub

Sub dynamixelTorqueOff()
    Dim send_packet() As Byte
    send_packet() = dynamixelTorqueOffPacket()
    ec.Binary = send_packet()
    Call comClose
End Sub


