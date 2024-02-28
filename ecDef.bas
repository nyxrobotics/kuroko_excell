Attribute VB_Name = "ecDef"
Option Explicit

'   EasyComm
'   Module ecDef.bas
'   Copyright(c) 2000 T.Kinoshita
'   Copyright(c) 2001 T.Kinoshita
'   Copyright(c) 2002 T.Kinoshita
'   Copyright(c) 2003 T.Kinoshita

'   2001.10.23 ReadFile,WriteFile�̏C��

'----------------------
'   �ϐ��E�萔�̒�`
'----------------------
Private Const Version As String = "1.71"

Type PortSetting
    Handle As Long          ' �t�@�C���n���h��
    Delimiter As String     ' �f���~�^
            ' .Delimiter = ""           : CR    �i�f�t�H���g�j
            ' .Delimiter = "CR"         : CR
            ' .Delimiter = "LF"         : LF
            ' .Delimiter = "CRLF"       : CRLF
            ' .Delimiter = "LFCR"       : LFCR
    ' Version 1.51�Œǉ�
    LineInTimeOut As Long   ' AsciiLine�v���p�e�B�̓ǂݏo���^�C���A�E�g(mS)
    ' �����܂�
End Type

Public Const ecMaxPort As Long = 50&                ' �ő�|�[�g�ԍ�
Public ecH(0 To ecMaxPort) As PortSetting           ' �|�[�g�̃n���h���C�f���~�^
Public Const ecMinimumInBuffer As Long = 2048&      ' �ŏ���M�o�b�t�@�T�C�Y
Public Const ecMinimumOutBuffer As Long = 2048&     ' �ŏ����M�o�b�t�@�T�C�Y
Public Cn As Long                                   ' �����Ώۂ̃|�[�g�ԍ�


'�W���ݒ�
Public Const ecSetting As String = "9600,n,8,1"     '�W���ʐM����
Public Const ecReadIntervalTimeout = 1000&          '���̕������͂܂łP�b�Ń^�C���A�E�g
Public Const ecReadTotalTimeoutConstant = 0&
Public Const ecReadTotalTimeoutMultiplier = 0&
Public Const ecWriteTotalTimeoutConstant = 0&
Public Const ecWriteTotalTimeoutMultiplier = 1000&  '�P�������M����̂ɂP�b�Ń^�C���A�E�g

Public Const ecInBufferSize As Long = 2048&
    '���̓o�b�t�@�ɐ���������Ƃ��͂��̒l��W���ݒ�l�Ƃ��邪�C�������̎���ecInBufferSize��W���ɂ���D
Public Const ecXonLim As Long = 256&                '�o�b�t�@�������l
Public Const ecXoffLim As Long = 256&               '�o�b�t�@�ŏ��󂫗e��
Public Const ecOutBufferSize As Long = 2048&
    '�o�̓o�b�t�@�ɐ���������Ƃ��͂��̒l��W���ݒ�l�Ƃ��邪�C�������̎���ecOutBufferSize��W���ɂ���D



'----------------------
'   API�̐錾
'----------------------

'======================
'CreateFile
'======================
'   �t�@�C���I�[�v��api
'   ������ = �n���h���C���s�� = INVALID_HANDLE_VALUE
'   lpSecurityAttributes �� SECURITY_ATTRIBUTES�ւ̃|�C���^
'   �g��Ȃ��̂� long �Œ�`���� Null(&0)��n���B�iByVal�ɕύX�j
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
'�p�����[�^�̈Ӗ�
'   lpFileName             �|�[�g�̘_����
'       "COM1","COM2",...
'   dwDesiredAccess        �A�N�Z�X���[�h
'      �ǂݏo����p GENERIC_READ
'      �������ݐ�p GENERIC_WRITE
'      �o����      GENERIC_READ Or GENERIC_WRITE
'      �ʏ�̃V���A���|�[�g�ł͑o�������w��
'   dwShareMode            ���L�t���O�B
'       �V���A���|�[�g�͋��L�ł��Ȃ��̂� &0 ��n���B
'   lpSecurityAttributes   �Z�L�����e�B�����Ɏw��B
'       �q�v���Z�X�Ɍp�����Ȃ��̂ŁC�f�t�H���g�̑������w�肷��悤�� &0 ��n���B
'   dwCreationDisposition  �t�@�C�����J�����@���w�肷��B
'       �|�[�g�͐V�K�쐬����̂ł͂Ȃ��C�����Ȃ̂�OPEN_EXISTING��n���B
'   dwFlagsAndAttributes   �t�@�C�������̎w��B
'       �V���A���|�[�g�ł� &0�C�܂���FILE_FLAG_OVERLAPPED��n���B
'       FILE_FLAG_OVERLAPPED�̏ꍇ�C�|�[�g�͔񓯊�I/O���[�h�ŊJ�����B
'   hTemplateFile
'       �|�[�g�̃I�[�v���ɂ͊֌W�Ȃ��̂� &0��n���B
'�⑫����
'   �f�o�b�O���̓v���O�����̒��f�ɂ��|�[�g���J���ꂽ�܂܂ɂȂ��Ă��܂����Ƃ�����B
'   �J���ꂽ�|�[�g�̃n���h�����킩��Ȃ��ƕ��邱�Ƃ��ł��Ȃ����߁C�|�[�g�̃n���h
'   �����V�[�g�ɏ������ނ悤�ȃ}�N���������Ă����Ƃ悢�B


'======================
'CloseHandle
'======================
'   �t�@�C���̃N���[�Y
'   ������ <>0�C���s�� = 0
Public Declare PtrSafe Function _
    CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) _
As Long


'======================
'GetCommState
'======================
'   �|�[�g�̐ݒ�l��DCB�ɓǂݏo��
'   ������ <>0�C���s�� = 0
Public Declare PtrSafe Function _
    GetCommState Lib "kernel32" ( _
        ByVal nCid As Long, _
        ByRef lfDCB As DCB) _
As Long

'======================
'SetCommState
'======================
'   DCB�̓��e��ݒ肷��
'   ������ <>0�C���s�� = 0
Public Declare PtrSafe Function _
    SetCommState Lib "kernel32" ( _
        ByVal hCommDev As Long, _
        ByRef lfDCB As DCB) _
As Long


'======================
'BuildCommDCB
'======================
'������ɂ��|�[�g�̐ݒ�
Public Declare PtrSafe Function _
    BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" ( _
    ByVal lpDef As String, _
    ByRef lfDCB As DCB) _
As Long

'======================
'GetCommTimeouts
'======================
'�^�C���A�E�g�̓ǂݏo��
Public Declare PtrSafe Function _
    GetCommTimeouts Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lfCOMMTIMEOUTS As COMMTIMEOUTS) _
As Long

'======================
'SetCommTimeouts
'======================
'�^�C���A�E�g�̐ݒ�
Public Declare PtrSafe Function _
    SetCommTimeouts Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lfCOMMTIMEOUTS As COMMTIMEOUTS) _
As Long

'======================
'PurgeComm
'======================
'�o�b�t�@�̃N���A
Public Declare PtrSafe Function _
    PurgeComm Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal dwFlags As Long) _
As Long

'======================
'ClearCommError
'======================
'�o�b�t�@�̏�Ԃ��擾
Public Declare PtrSafe Function _
    ClearCommError Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lpErrors As Long, _
        ByRef lpStat As COMSTAT) _
As Long

'======================
'SetupComm
'======================
'�o�b�t�@�T�C�Y�̎w��
Public Declare PtrSafe Function _
    SetupComm Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal dwInQueue As Long, _
        ByVal dwOutQueue As Long) _
As Long

'======================
'GetCommProperties
'======================
'�|�[�g�̎d�l�̎擾
Public Declare PtrSafe Function _
    GetCommProperties Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lfCOMMPROP As COMMPROP) _
As Long

'======================
'WriteFile
'======================
'�|�[�g�o��API
'lpBuffer�́C�o�C�i���R�[�h���������Ƃ�����̂�String�ł͂Ȃ�Any�Ő錾����
'lpOverlapped�͎g��Ȃ��Ƃ���Null��n���̂�Long�܂���Any
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

'win32api.txt�ł�lpOverlapped��ByRef�Œ�`����Ă��邪ByVal�̌��
'��`���e
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
'�|�[�g����API
'lpBuffer�́C�o�C�i���R�[�h���������Ƃ�����̂�String�ł͂Ȃ�Any�Ő錾����
'lpOverlapped�͎g��Ȃ��Ƃ���Null��n���̂�Long�܂���Any
'---Vers1.00
'Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
'    ByVal hFile As Long, _
'    ByRef lpBuffer As Any, _
'    ByVal nNumberOfBytesToRead As Long, _
'    ByRef lecActiveOfBytesRead As Long, _
'    ByVal lpOverlapped As Long) _
'As Long
'---Vers1.01

'win32api.txt�ł�lpOverlapped��ByRef�Œ�`����Ă��邪ByVal�̌��
'��`���e
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
Public Const PURGE_TXCLEAR = &H4     '  ���M�o�b�t�@�N���A
Public Const PURGE_RXCLEAR = &H8     '  ��M�o�b�t�@�N���A

'======================
'Sleep
'======================
'��~�^�C�}�[�֐�
'�w�莞�ԁi�~���b�j�C���s�𒆒f����D
Public Declare PtrSafe Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long)

'======================
'EscapeCommFunction
'======================
'RTS,DTR�̋�������
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
'RTS,DTR�̏�Ԃ̓ǂݎ��
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
'Break�M���̑��M
Public Declare PtrSafe Function SetCommBreak Lib "kernel32" ( _
    ByVal nCid As Long) _
As Long

'======================
'ClearCommBreak
'======================
'Break�M���̑��M���~
Public Declare PtrSafe Function ClearCommBreak Lib "kernel32" ( _
    ByVal nCid As Long) _
As Long

'======================
'GetTickCount
'======================
'Windows �N������̌o�ߎ��Ԃ��~���b�P�ʂŎ擾���܂��D
'API���ł́C�o�ߎ��Ԃ͕����Ȃ��̒����� DWORD �^�ŕۑ�����Ă��܂��D
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () _
As Long

'======================
'GetLocalTime
'======================
'���݂̃��[�J�����Ԃ�mS�P�ʂ܂Ŏ擾���܂��D
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
'   �\���̂̒�`
'----------------------

'DCB�\����
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

'DCB�^�ϐ��̒�`
Public fDCB As DCB

'   �����o�̈Ӗ�
'DCBlength
'   DCB�\���̂̃o�C�g�T�C�Y
'BaudRate
'   �{�[���[�g
'fBitFields
'   �e�r�b�g�ŋ@�\���w�肷��
'   �e�r�b�g��1�̂Ƃ��̈Ӗ��͎��̒ʂ�
' bit1   fBinary
'   �o�C�i�����[�h���g�p�\
'   Win32 API�͔�o�C�i�����[�h�]�����T�|�[�g���Ȃ��̂ł��̃����o�[�͏�ɂP
' bit2   fParity
'   �p���e�B�`�F�b�N���g�p
' bit3   fOutxCtsFlow
'   CTS���o�̓t���[����ŊĎ������
' bit4   fOutxDsrFlow
'   DSR���o�̓t���[����ŊĎ������
' bit5,6 fDtrControl
'   DTR�ɂ��t���[�����2�r�b�g�Ŏw��
Public Const DTR_CONTROL_DISABLE = &H0      'DTR���C����OFF
Public Const DTR_CONTROL_ENABLE = &H1       'DTR���C����ON
Public Const DTR_CONTROL_HANDSHAKE = &H2    'DTR�ɂ��n���h�V�F�[�N
' bit7   fDsrSensitivity
'   DSR��OFF�̊ԂɎ�M�����f�[�^�𖳎�
' bit8   fTXContinueOnXoff
'   ��M�o�b�t�@���t���ɂȂ�XoffChar�����𑗐M������ɑ��M���~�߂�
' bit9   fOutX
'   ���M����XON/XOFF�t���[��L���ɂ���
' bit10  fInX
'   ���M����XON/XOFF�t���[��L���ɂ���
' bit11  fErrorChar
'   �p���e�B�G���[�����������Ƃ��ɁC������ErrorChar�ɒu��������
' bit12  fNull
'   �k�������i�l�O�̃f�[�^�j��j������
' bit13,14  fRtsControl
'   2�r�b�g��RTS�̃t���[������w��
Public Const RTS_CONTROL_DISABLE = &H0      'RTS��OFF
Public Const RTS_CONTROL_ENABLE = &H1       'RTS��ON
Public Const RTS_CONTROL_HANDSHAKE = &H2    'RTS�ɂ��n���h�V�F�[�N*1
Public Const RTS_CONTROL_TOGGLE = &H3       'RTS�ɂ��n���h�V�F�[�N*2
'       *1 ��M�o�b�t�@�� 3/4�ȏ㖄�܂��RTS��ON�C1/2�ȉ��ɂȂ��OFF
'       *2 ��M�o�b�t�@�Ƀf�[�^���c���Ă����RTS��ON�C�[���Ȃ��OFF
' bit15  fAbortOnError
'   �G���[���N�������Ƃ��ɂ͓ǂݏ������I��
' bit16  fDummy2
'   ���g�p

'wReserved
'   ���g�p�B�[�����Z�b�g���Ȃ���΂Ȃ�Ȃ�
'XonLim
'   ��M�o�b�t�@�̃f�[�^�����o�C�g�ȏ�ɂȂ�����XON�𑗂邩���w��
'XoffLim
'   ��M�o�b�t�@�̃f�[�^�����o�C�g�����ɂȂ�����XON�𑗂邩���w��
'ByteSize
'   �f�[�^�̃r�b�g��
'Parity
'   �p���e�B�̕���
Public Const NOPARITY = 0       '�p���e�B�Ȃ�
Public Const ODDPARITY = 1      '��p���e�B
Public Const EVENPARITY = 2     '�����p���e�B
Public Const MARKPARITY = 3     '��Ƀ}�[�N
Public Const SPACEPARITY = 4    '��ɃX�y�[�X
'StopBits
'   �X�g�b�v�r�b�g�̐�
Public Const ONESTOPBIT = 0     '1 bit
Public Const ONE5STOPBITS = 1   '1.5 bit
Public Const TWOSTOPBITS = 2    '2 bit
'XonChar
'   XON�̑��M����
'XoffChar
'   XOFF�̑��M����
'ErrorChar
'   �p���e�B�G���[�������ɒu�������镶��
'EofChar
'   ��o�C�i�����[�h�̂Ƃ��ɂ��̕�������M����ƃf�[�^�I�����݂Ȃ�
'   ������Win32 API�ł͔�o�C�i�����[�h���T�|�[�g���Ȃ��̂Ŗ��Ӗ�
'EvtChar
'   ���̕�������M����ƃC�x���g������
'wReserved1
'   ���g�p


'COMMTIMEOUT�\����
Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

'COMMTIMEOUTS�^�ϐ��̒�`
Public fCOMMTIMEOUTS As COMMTIMEOUTS


'COMSTAT�\���̂̒�`
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

'COMSTAT�^�ϐ��̒�`
Public fCOMSTAT As COMSTAT

'COMMPROP�\���̂̒�`
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

'COMMPROP�^�ϐ��̒�`
Public fCOMMPROP As COMMPROP

'���[�J���^�C���\����
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

'���[�J���^�C���^�ϐ��̒�`
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



