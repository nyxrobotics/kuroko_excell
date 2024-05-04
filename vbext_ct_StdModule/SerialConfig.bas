Attribute VB_Name = "SerialConfig"
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
End Sub


Sub sendFreePacket(ByRef data_in() As Byte)
    Dim parity_bit As Currency
    Dim one_byte_time As Currency
    Dim send_time As Currency
    Call comClose
    Call comOpen
    parity_bit = 0
    If COM_PARITY <> "n" Then
        parity_bit = 1
    End If
    one_byte_time = (1 + CCur(COM_LENGTH) + parity_bit + CCur(COM_STOP)) / CCur(COM_BAUD)
    send_time = one_byte_time * CCur(UBound(data_in))
    ec.Binary = data_in()
    qpcWait (send_time)
    qpcWaitMs (20)
    Dim received() As Byte
    received() = ec.Binary
    Dim received_size As Long
    received_size = UBound(received)
End Sub
