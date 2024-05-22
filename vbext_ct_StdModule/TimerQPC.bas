Attribute VB_Name = "TimerQPC"
'標準モジュールの(General)(Declarations)へ記述します

#If VBA7 Then
Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Boolean
Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Boolean
#Else
Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Boolean
#End If

'CPUの実行速度
Dim QPC_FEAQ As Currency
Dim TIME_START As Currency

Sub qpcInit()
    '周波数所得(一秒当たりのカウント値を所得)
    Call QueryPerformanceFrequency(QPC_FEAQ)
    'カウンターの取得
    Call QueryPerformanceCounter(TIME_START)
End Sub

Sub qpcRestart()
    Call QueryPerformanceCounter(TIME_START)
End Sub

Sub qpcWaitMs(interval As Currency)
    Dim time_now As Currency
    'カウンターの取得
    Call QueryPerformanceCounter(time_now)
    '経過時間を1ms単位に変換
    Dim time_passed_ms As Currency
    time_passed_ms = (time_now - TIME_START) / (QPC_FEAQ / 1000)
    '指定した時間[ms]経過するまで待機
    Do While (time_passed_ms < interval)
        Call QueryPerformanceCounter(time_now)
        time_passed_ms = (time_now - TIME_START) / (QPC_FEAQ / 1000)
    Loop
    TIME_START = time_now
End Sub


Sub qpcWait(interval As Currency)
    Dim time_now As Currency
    Call QueryPerformanceCounter(time_now)
    Dim time_passed As Currency
    time_passed = (time_now - TIME_START) / QPC_FEAQ
    Do While (time_passed < interval)
        Call QueryPerformanceCounter(time_now)
        time_passed = (time_now - TIME_START) / QPC_FEAQ
    Loop
    TIME_START = time_now
End Sub
