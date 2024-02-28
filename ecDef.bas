Attribute VB_Name = "ecDef"
Option Explicit

'   EasyComm
'   Module ecDef.bas
'   Copyright(c) 2000 T.Kinoshita
'   Copyright(c) 2001 T.Kinoshita
'   Copyright(c) 2002 T.Kinoshita
'   Copyright(c) 2003 T.Kinoshita

'   2001.10.23 ReadFile,WriteFileの修正

'----------------------
'   変数・定数の定義
'----------------------
Private Const Version As String = "1.71"

Type PortSetting
    Handle As Long          ' ファイルハンドル
    Delimiter As String     ' デリミタ
            ' .Delimiter = ""           : CR    （デフォルト）
            ' .Delimiter = "CR"         : CR
            ' .Delimiter = "LF"         : LF
            ' .Delimiter = "CRLF"       : CRLF
            ' .Delimiter = "LFCR"       : LFCR
    ' Version 1.51で追加
    LineInTimeOut As Long   ' AsciiLineプロパティの読み出しタイムアウト(mS)
    ' ここまで
End Type

Public Const ecMaxPort As Long = 50&                ' 最大ポート番号
Public ecH(0 To ecMaxPort) As PortSetting           ' ポートのハンドル，デリミタ
Public Const ecMinimumInBuffer As Long = 2048&      ' 最小受信バッファサイズ
Public Const ecMinimumOutBuffer As Long = 2048&     ' 最小送信バッファサイズ
Public Cn As Long                                   ' 処理対象のポート番号


'標準設定
Public Const ecSetting As String = "9600,n,8,1"     '標準通信条件
Public Const ecReadIntervalTimeout = 1000&          '次の文字入力まで１秒でタイムアウト
Public Const ecReadTotalTimeoutConstant = 0&
Public Const ecReadTotalTimeoutMultiplier = 0&
Public Const ecWriteTotalTimeoutConstant = 0&
Public Const ecWriteTotalTimeoutMultiplier = 1000&  '１文字送信するのに１秒でタイムアウト

Public Const ecInBufferSize As Long = 2048&
    '入力バッファに制限があるときはその値を標準設定値とするが，無制限の時はecInBufferSizeを標準にする．
Public Const ecXonLim As Long = 256&                'バッファしきい値
Public Const ecXoffLim As Long = 256&               'バッファ最小空き容量
Public Const ecOutBufferSize As Long = 2048&
    '出力バッファに制限があるときはその値を標準設定値とするが，無制限の時はecOutBufferSizeを標準にする．



'----------------------
'   APIの宣言
'----------------------

'======================
'CreateFile
'======================
'   ファイルオープンapi
'   成功時 = ハンドル，失敗時 = INVALID_HANDLE_VALUE
'   lpSecurityAttributes は SECURITY_ATTRIBUTESへのポインタ
'   使わないので long で定義して Null(&0)を渡す。（ByValに変更）
Public Declare PtrSafe Function _
    CreateFile Lib "kernel32" Alias "CreateFileA" ( _
        ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) _
As Long
'パラメータの意味
'   lpFileName             ポートの論理名
'       "COM1","COM2",...
'   dwDesiredAccess        アクセスモード
'      読み出し専用 GENERIC_READ
'      書き込み専用 GENERIC_WRITE
'      双方向      GENERIC_READ Or GENERIC_WRITE
'      通常のシリアルポートでは双方向を指定
'   dwShareMode            共有フラグ。
'       シリアルポートは共有できないので &0 を渡す。
'   lpSecurityAttributes   セキュリティ属性に指定。
'       子プロセスに継承しないので，デフォルトの属性を指定するように &0 を渡す。
'   dwCreationDisposition  ファイルを開く方法を指定する。
'       ポートは新規作成するのではなく，既存なのでOPEN_EXISTINGを渡す。
'   dwFlagsAndAttributes   ファイル属性の指定。
'       シリアルポートでは &0，またはFILE_FLAG_OVERLAPPEDを渡す。
'       FILE_FLAG_OVERLAPPEDの場合，ポートは非同期I/Oモードで開かれる。
'   hTemplateFile
'       ポートのオープンには関係ないので &0を渡す。
'補足説明
'   デバッグ中はプログラムの中断によりポートが開かれたままになってしまうことがある。
'   開かれたポートのハンドルがわからないと閉じることができないため，ポートのハンド
'   ルをシートに書き込むようなマクロを加えておくとよい。


'======================
'CloseHandle
'======================
'   ファイルのクローズ
'   成功時 <>0，失敗時 = 0
Public Declare PtrSafe Function _
    CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) _
As Long


'======================
'GetCommState
'======================
'   ポートの設定値をDCBに読み出す
'   成功時 <>0，失敗時 = 0
Public Declare PtrSafe Function _
    GetCommState Lib "kernel32" ( _
        ByVal nCid As Long, _
        ByRef lfDCB As DCB) _
As Long

'======================
'SetCommState
'======================
'   DCBの内容を設定する
'   成功時 <>0，失敗時 = 0
Public Declare PtrSafe Function _
    SetCommState Lib "kernel32" ( _
        ByVal hCommDev As Long, _
        ByRef lfDCB As DCB) _
As Long


'======================
'BuildCommDCB
'======================
'文字列によるポートの設定
Public Declare PtrSafe Function _
    BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" ( _
    ByVal lpDef As String, _
    ByRef lfDCB As DCB) _
As Long

'======================
'GetCommTimeouts
'======================
'タイムアウトの読み出し
Public Declare PtrSafe Function _
    GetCommTimeouts Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lfCOMMTIMEOUTS As COMMTIMEOUTS) _
As Long

'======================
'SetCommTimeouts
'======================
'タイムアウトの設定
Public Declare PtrSafe Function _
    SetCommTimeouts Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lfCOMMTIMEOUTS As COMMTIMEOUTS) _
As Long

'======================
'PurgeComm
'======================
'バッファのクリア
Public Declare PtrSafe Function _
    PurgeComm Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal dwFlags As Long) _
As Long

'======================
'ClearCommError
'======================
'バッファの状態を取得
Public Declare PtrSafe Function _
    ClearCommError Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lpErrors As Long, _
        ByRef lpStat As COMSTAT) _
As Long

'======================
'SetupComm
'======================
'バッファサイズの指定
Public Declare PtrSafe Function _
    SetupComm Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal dwInQueue As Long, _
        ByVal dwOutQueue As Long) _
As Long

'======================
'GetCommProperties
'======================
'ポートの仕様の取得
Public Declare PtrSafe Function _
    GetCommProperties Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lfCOMMPROP As COMMPROP) _
As Long

'======================
'WriteFile
'======================
'ポート出力API
'lpBufferは，バイナリコードを扱うことがあるのでStringではなくAnyで宣言する
'lpOverlappedは使わないときにNullを渡すのでLongまたはAny
'---Vers 1.00
'Public Declare PtrSafe Function WriteFile Lib "kernel32" ( _
'    ByVal hFile As Long, _
'    ByRef lpBuffer As Any, _
'    ByVal nNumberOfBytesToWrite As Long, _
'    ByRef lecActiveOfBytesWritten As Long, _
'    ByVal lpOverlapped As Long) _
'As Long
'---Vers1.01

'Public Declare PtrSafe Function WriteFile Lib "kernel32" ( _
'    ByVal hFile As Long, _
'    ByRef lpBuffer As Any, _
'    ByVal nNumberOfBytesToWrite As Long, _
'    ByRef lecActiveOfBytesWritten As Long, _
'    ByRef lpOverlapped As Long) _
'As Long

'win32api.txtではlpOverlappedがByRefで定義されているがByValの誤り
'定義内容
'Declare PtrSafe Function WriteFile Lib "kernel32" ( _
'    ByVal hFile As Long, _
'    lpBuffer As Any, _
'    ByVal nNumberOfBytesToWrite As Long, _
'    lpNumberOfBytesWritten As Long, _
'    lpOverlapped As OVERLAPPED) _
'As Long
Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) _
As Long

'======================
'ReadFile
'======================
'ポート入力API
'lpBufferは，バイナリコードを扱うことがあるのでStringではなくAnyで宣言する
'lpOverlappedは使わないときにNullを渡すのでLongまたはAny
'---Vers1.00
'Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
'    ByVal hFile As Long, _
'    ByRef lpBuffer As Any, _
'    ByVal nNumberOfBytesToRead As Long, _
'    ByRef lecActiveOfBytesRead As Long, _
'    ByVal lpOverlapped As Long) _
'As Long
'---Vers1.01

'win32api.txtではlpOverlappedがByRefで定義されているがByValの誤り
'定義内容
'Declare PtrSafe Function ReadFile Lib "kernel32" ( _
'    ByVal hFile As Long, _
'    lpBuffer As Any, _
'    ByVal nNumberOfBytesToRead As Long, _
'    lpNumberOfBytesRead As Long, _
'    lpOverlapped As OVERLAPPED) _
'As Long

Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lecActiveOfBytesRead As Long, _
    ByVal lpOverlapped As Long) _
As Long


'---

' API
Public Const INVALID_HANDLE_VALUE = -1
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3

'  PURGE function flags.
Public Const PURGE_TXCLEAR = &H4     '  送信バッファクリア
Public Const PURGE_RXCLEAR = &H8     '  受信バッファクリア

'======================
'Sleep
'======================
'停止タイマー関数
'指定時間（ミリ秒），実行を中断する．
Public Declare PtrSafe Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long)

'======================
'EscapeCommFunction
'======================
'RTS,DTRの強制制御
Public Declare PtrSafe Function _
    EscapeCommFunction Lib "kernel32" ( _
        ByVal nCid As Long, _
        ByVal nFunc As Long) _
As Long

' Escape Functions
Public Const SETRTS = 3 '  Set RTS high
Public Const CLRRTS = 4 '  Set RTS low
Public Const SETDTR = 5 '  Set DTR high
Public Const CLRDTR = 6 '  Set DTR low


'======================
'GetCommModemStatus
'======================
'RTS,DTRの状態の読み取り
Public Declare PtrSafe Function _
    GetCommModemStatus Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpModemStat As Long) _
As Long

'  Modem Status Flags
Public Const MS_CTS_ON = &H10&
Public Const MS_DSR_ON = &H20&
Public Const MS_RING_ON = &H40&
Public Const MS_RLSD_ON = &H80&

'======================
'SetCommBreak
'======================
'Break信号の送信
Public Declare PtrSafe Function SetCommBreak Lib "kernel32" ( _
    ByVal nCid As Long) _
As Long

'======================
'ClearCommBreak
'======================
'Break信号の送信中止
Public Declare PtrSafe Function ClearCommBreak Lib "kernel32" ( _
    ByVal nCid As Long) _
As Long

'======================
'GetTickCount
'======================
'Windows 起動からの経過時間をミリ秒単位で取得します．
'API内では，経過時間は符号なしの長整数 DWORD 型で保存されています．
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () _
As Long

'======================
'GetLocalTime
'======================
'現在のローカル時間をmS単位まで取得します．
Public Declare PtrSafe Sub GetLocalTime Lib "kernel32" ( _
    lpSystemTime As SYSTEMTIME)

'======================
'GetTempPath
'======================
Public Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) _
As Long


'----------------------
'   構造体の定義
'----------------------

'DCB構造体
Type DCB
    DCBlength As Long
    BaudRate As Long
    fBitFields As Long 'See Comments in Win32API.Txt
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer 'Reserved; Do Not Use
End Type

'DCB型変数の定義
Public fDCB As DCB

'   メンバの意味
'DCBlength
'   DCB構造体のバイトサイズ
'BaudRate
'   ボーレート
'fBitFields
'   各ビットで機能を指定する
'   各ビットが1のときの意味は次の通り
' bit1   fBinary
'   バイナリモードが使用可能
'   Win32 APIは非バイナリモード転送をサポートしないのでこのメンバーは常に１
' bit2   fParity
'   パリティチェックを使用
' bit3   fOutxCtsFlow
'   CTSが出力フロー制御で監視される
' bit4   fOutxDsrFlow
'   DSRが出力フロー制御で監視される
' bit5,6 fDtrControl
'   DTRによるフロー制御を2ビットで指定
Public Const DTR_CONTROL_DISABLE = &H0      'DTRラインをOFF
Public Const DTR_CONTROL_ENABLE = &H1       'DTRラインをON
Public Const DTR_CONTROL_HANDSHAKE = &H2    'DTRによるハンドシェーク
' bit7   fDsrSensitivity
'   DSRがOFFの間に受信したデータを無視
' bit8   fTXContinueOnXoff
'   受信バッファがフルになりXoffChar文字を送信した後に送信を止める
' bit9   fOutX
'   送信時にXON/XOFFフローを有効にする
' bit10  fInX
'   送信時にXON/XOFFフローを有効にする
' bit11  fErrorChar
'   パリティエラーが発生したときに，文字をErrorCharに置き換える
' bit12  fNull
'   ヌル文字（値０のデータ）を破棄する
' bit13,14  fRtsControl
'   2ビットでRTSのフロー制御を指定
Public Const RTS_CONTROL_DISABLE = &H0      'RTSをOFF
Public Const RTS_CONTROL_ENABLE = &H1       'RTSをON
Public Const RTS_CONTROL_HANDSHAKE = &H2    'RTSによるハンドシェーク*1
Public Const RTS_CONTROL_TOGGLE = &H3       'RTSによるハンドシェーク*2
'       *1 受信バッファが 3/4以上埋まるとRTSがON，1/2以下になるとOFF
'       *2 受信バッファにデータが残っていればRTSがON，ゼロならばOFF
' bit15  fAbortOnError
'   エラーが起こったときには読み書きを終了
' bit16  fDummy2
'   未使用

'wReserved
'   未使用。ゼロをセットしなければならない
'XonLim
'   受信バッファのデータが何バイト以上になったらXONを送るかを指定
'XoffLim
'   受信バッファのデータが何バイト未満になったらXONを送るかを指定
'ByteSize
'   データのビット数
'Parity
'   パリティの方式
Public Const NOPARITY = 0       'パリティなし
Public Const ODDPARITY = 1      '奇数パリティ
Public Const EVENPARITY = 2     '偶数パリティ
Public Const MARKPARITY = 3     '常にマーク
Public Const SPACEPARITY = 4    '常にスペース
'StopBits
'   ストップビットの数
Public Const ONESTOPBIT = 0     '1 bit
Public Const ONE5STOPBITS = 1   '1.5 bit
Public Const TWOSTOPBITS = 2    '2 bit
'XonChar
'   XONの送信文字
'XoffChar
'   XOFFの送信文字
'ErrorChar
'   パリティエラー発生時に置き換える文字
'EofChar
'   非バイナリモードのときにこの文字を受信するとデータ終了をみなす
'   ただしWin32 APIでは非バイナリモードをサポートしないので無意味
'EvtChar
'   この文字を受信するとイベントが発生
'wReserved1
'   未使用


'COMMTIMEOUT構造体
Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

'COMMTIMEOUTS型変数の定義
Public fCOMMTIMEOUTS As COMMTIMEOUTS


'COMSTAT構造体の定義
Type COMSTAT
    fBitFields As Long 'See Comment in Win32API.Txt
    cbInQue As Long
    cbOutQue As Long
End Type
' The eight actual COMSTAT bit-sized data fields within the four bytes of fBitFields can be manipulated by bitwise logical And/Or operations.
' FieldName     Bit #     Description
' ---------     -----     ---------------------------
' fCtsHold        1       Tx waiting for CTS signal
' fDsrHold        2       Tx waiting for DSR signal
' fRlsdHold       3       Tx waiting for RLSD signal
' fXoffHold       4       Tx waiting, XOFF char rec'd
' fXoffSent       5       Tx waiting, XOFF char sent
' fEof            6       EOF character sent
' fTxim           7       character waiting for Tx
' fReserved       8       reserved (25 bits)

'COMSTAT型変数の定義
Public fCOMSTAT As COMSTAT

'COMMPROP構造体の定義
Type COMMPROP
    wPacketLength As Integer
    wPacketVersion As Integer
    dwServiceMask As Long
    dwReserved1 As Long
    dwMaxTxQueue As Long
    dwMaxRxQueue As Long
    dwMaxBaud As Long
    dwProvSubType As Long
    dwProvCapabilities As Long
    dwSettableParams As Long
    dwSettableBaud As Long
    wSettableData As Integer
    wSettableStopParity As Integer
    dwCurrentTxQueue As Long
    dwCurrentRxQueue As Long
    dwProvSpec1 As Long
    dwProvSpec2 As Long
    wcProvChar(1) As Integer
End Type

'COMMPROP型変数の定義
Public fCOMMPROP As COMMPROP

'ローカルタイム構造体
Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'ローカルタイム型変数の定義
Public fSYSTEMTIME As SYSTEMTIME

'  Settable baud rates in the provider.
Public Const BAUD_075 = &H1&
Public Const BAUD_110 = &H2&
Public Const BAUD_134_5 = &H4&
Public Const BAUD_150 = &H8&
Public Const BAUD_300 = &H10&
Public Const BAUD_600 = &H20&
Public Const BAUD_1200 = &H40&
Public Const BAUD_1800 = &H80&
Public Const BAUD_2400 = &H100&
Public Const BAUD_4800 = &H200&
Public Const BAUD_7200 = &H400&
Public Const BAUD_9600 = &H800&
Public Const BAUD_14400 = &H1000&
Public Const BAUD_19200 = &H2000&
Public Const BAUD_38400 = &H4000&
Public Const BAUD_56K = &H8000&
Public Const BAUD_128K = &H10000
Public Const BAUD_115200 = &H20000
Public Const BAUD_57600 = &H40000
Public Const BAUD_USER = &H10000000


'  Provider capabilities flags.
Public Const PCF_DTRDSR = &H1&
Public Const PCF_RTSCTS = &H2&
Public Const PCF_RLSD = &H4&
Public Const PCF_PARITY_CHECK = &H8&
Public Const PCF_XONXOFF = &H10&
Public Const PCF_SETXCHAR = &H20&
Public Const PCF_TOTALTIMEOUTS = &H40&
Public Const PCF_INTTIMEOUTS = &H80&
Public Const PCF_SPECIALCHARS = &H100&
Public Const PCF_16BITMODE = &H200&

'  Comm provider settable parameters.
Public Const SP_PARITY = &H1&
Public Const SP_BAUD = &H2&
Public Const SP_DATABITS = &H4&
Public Const SP_STOPBITS = &H8&
Public Const SP_HANDSHAKING = &H10&
Public Const SP_PARITY_CHECK = &H20&
Public Const SP_RLSD = &H40&

'  Settable Data Bits
Public Const DATABITS_5 = &H1&
Public Const DATABITS_6 = &H2&
Public Const DATABITS_7 = &H4&
Public Const DATABITS_8 = &H8&
Public Const DATABITS_16 = &H10&
Public Const DATABITS_16X = &H20&

'  Settable Stop and Parity bits.
Public Const STOPBITS_10 = &H1&
Public Const STOPBITS_15 = &H2&
Public Const STOPBITS_20 = &H4&
Public Const PARITY_NONE = &H100&
Public Const PARITY_ODD = &H200&
Public Const PARITY_EVEN = &H400&
Public Const PARITY_MARK = &H800&
Public Const PARITY_SPACE = &H1000&



