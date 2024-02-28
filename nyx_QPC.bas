Attribute VB_Name = "nyx_QPC"
'�W�����W���[����(General)(Declarations)�֋L�q���܂�

#If VBA7 Then
Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Boolean
Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Boolean
#Else
Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Boolean
#End If

'CPU�̎��s���x
Dim gQPC_frequency As Currency
Dim gQPC_time_start As Currency
Dim gQPC_time_now As Currency
Dim gQPC_time_prev As Currency
Dim gQPC_time_stop As Currency




Sub QPC_start_counting()
    '���g������(��b������̃J�E���g�l������)
    Call QueryPerformanceFrequency(gQPC_frequency)
    
    '�J�E���^�[�̎擾,1ms�P��
    Dim count As Currency
    Call QueryPerformanceCounter(count)
    gQPC_time_start = count / (gQPC_frequency / 1000)
    gQPC_time_now = gQPC_time_start
    gQPC_time_prev = gQPC_time_start
    
End Sub

Sub QPC_wait_ms(interval As Currency)
    Dim count As Currency
    '�J�E���^�[�̎擾,1ms�P��
    Call QueryPerformanceCounter(count)
    gQPC_time_now = count / (gQPC_frequency / 1000)

    Do While (gQPC_time_now - gQPC_time_prev < interval)
        '�J�E���^�[�̎擾,1ms�P��
        Call QueryPerformanceCounter(count)
        gQPC_time_now = count / (gQPC_frequency / 1000)
    Loop
    
    gQPC_time_prev = gQPC_time_now
End Sub

Sub QPC_test()
    Call QPC_start_counting
    Call QPC_wait_ms(1000)
End Sub
