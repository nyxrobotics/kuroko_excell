Attribute VB_Name = "ec"
Option Explicit
'   EasyComm
'   Module ec.bas

'�|�|�|�|�|�|EasyComm ���p�K��|�|�|�|�|�|
'   ���L�����ɓ��ӂ��ꂽ���݂̂����p���������D
'   �i�d�������b�������������Ɨ��L���܂��j
'
'�i�P�j������P�̂Ŕ̔����邱�Ƃ͏o���܂���D
'�i�Q�j�����𓮍삳�������ʁC�����Ȃ鑹�Q���������Ă��C�؉����͈�ؐӔC�𕉂��܂���D
'�i�R�j�����͌l�C��Ђ��킸�C�؉����̋��������R�ɔz�z���C�_�E�����[�h�ł��܂��D
'�i�S�j������Ǝ��̃A�v���P�[�V�����ɑg�ݍ��񂾏ꍇ�C�؉����̋��������R�ɔz�z�C�̔�
'       ���邱�Ƃ��o���܂��D
'�i�T�j�����́C�A�v���P�[�V�����̈ꕔ�Ƃ��ė��p����ꍇ�Ɍ����āC�؉����̋��������R
'       �ɉ������C�܂��͈ꕔ�����p���ė��p�C�z�z�C�̔����邱�Ƃ��o���܂��D
'
'   Version 1.00    Aug 25,2000  Created by T.Kinoshita
'   Version 1.10    Nov 21,2000
'   Version 1.20    Jan 14,2001
'       DCD,RI,Break�v���p�e�B�ǉ�
'   Version 1.30    Mar 31,2001
'       ���M�o�b�t�@�̏����ݒ�l��ύX
'       �f�t�H���g�̑���M�o�b�t�@�̐ݒ���@��ύX
'       Xerror=1�ŁCSTOP�̌�ɂ��ׂẴ|�[�g�����悤�ɃR�[�h��ǉ�
'   Version 1.31    May 17,2001
'       WAITmS�̃o�O�C���i���c�l�̂��w�E�ɂ��j
'   Version 1.40    Jul 25,2001
'       Binary�v���p�e�B�̏������݂��̃f�[�^�^�Ή����g��
'       ���[�h�[���̃G���[�������́C���ׂẴ|�[�g�����悤�Ɏd�l�ύX
'       �B�����\�b�h��CLoseAll��CloseAll�ɕύX
'       ���O�t�@�C���̔p�~
'       �f�t�H���g�̃n���h�V�F�[�L���O��"N"�ɕύX
'       �f���~�^�w��p��������擾���邽�߂̓ǂݎ���p�v���p�e�BDELIMS��ǉ�
'       �n���h�V�F�[�L���O�w��p��������擾���邽�߂̓ǂݎ���p�v���p�e�BHANDSHAKEs��ǉ�
'       �f�t�H���g�̐ݒ��ύX
'
'   Version 1.50    Oct 23,2001
'       Windows2000�ɑΉ��iReadFile,WriteFile�̒�`���C��
'
'   Version 1.51    Jan 17,2002
'       tomot����̂��ӌ����Q�l�ɁCAsciiLine�̓ǂݏo���^�C���A�E�g��ǉ����܂����D
'   Version 1.51a   Jan 17,2002
'       COMn�v���p�e�B�Ƀ[���ȉ��̒l���������Ƃ��C�|�[�g�Z�b�e�B���O�ϐ������Z�b�g����悤�ɕύX
'       COMnClose�v���p�e�B�ŕ����|�[�g�ɑ΂��Ă��C�|�[�g�Z�b�e�B���O�ϐ������Z�b�g����p�ɕύX
'   Version 1.60    Jul 9,2002
'       Version�v���p�e�B�ǉ�(Private��Public�ɕύX)
'       DoseSeconds�v���p�e�B�̒ǉ�
'       �ő哯���I�[�v���|�[�g����20����50�Ɋg��
'       �t�@�C���n���h�����L������e���|�����t�@�C��ecPort.tmp�̎g�p
'       AsciiLine�v���p�e�B�̎�M�ŁC�ݒ�ɂ���Ă͕�����̖�����chr(0)������o�O���C��
'   Version 1.61    Jul 23,2002
'       shin����̂��w�E�ɂ��C�e���|�����t�@�C���̃t�H���_��ThisWorkbook.Path�ɂ��Ă������߁C
'       Excel�ȊO�̃A�v���Ŏg�p�s�\�ɂȂ��Ă����o�O���C���D
'       Windows�̃e���|�����t�@�C���ɍ쐬����悤�ɂ����D
'   Version 1.62    Jul 31,2002
'       ���肳��̂��w�E�ɂ��Excel97�ł̕s����C��
'   Version 1.62a   Aug 10,2002
'       ���p�K���ύX
'   Version 1.70    Mar 6,2003
'       Binary�v���p�e�B�ǂݏo�����̃o�C�g�����w�肷��BinaryBytes�v���p�e�B��ǉ��D
'       Ascii�v���p�e�B�ǂݏo�����̃o�C�g�����w�肷��AsciiBytes�v���p�e�B��ǉ��D
'       ��������f�t�H���g�Ń[��(���o�[�W�����݊�)
'       InBufferClear���\�b�h��ǉ�(���\�b�h�͏��߂�)
'   Version 1.71    Mar 7,2003
'       �����f�Ӄo�[�W����
'       BinaryBytes,AsciiBytes�v���p�e�B�̃o�O�C��
'       �w��o�C�g���̃f�[�^���o�b�t�@�ɖ������́C��M�����܂Œ�~��ԂɂȂ�o�O�̏C��
'
Public Const Version As String = "1.71"
'
'   Copyright(c) 2000 T.Kinoshita
'   Copyright(c) 2001 T.Kinoshita
'   Copyright(c) 2002 T.Kinoshita
'   Copyright(c) 2003 T.Kinoshita


'===================================
' �ϐ��̒�`
'===================================

'-----------------------------------
'   AsciiBytes�v���p�e�B(Version1.70����ǉ�)
'-----------------------------------
'Ascii�v���p�e�B�ǂݏo�����Ɏ�M�o�b�t�@������o���o�C�g�����w�肵�܂�
'�f�t�H���g�̓[���ŁC�[���ȉ��̎��͎�M�o�b�t�@�̃f�[�^�����ׂĎ擾���܂��D
Public AsciiBytes As Long

'-----------------------------------
'   BinaryBytes�v���p�e�B(Version1.70����ǉ�)
'-----------------------------------
'Ascii�v���p�e�B�ǂݏo�����Ɏ�M�o�b�t�@������o���o�C�g�����w�肵�܂�
'�f�t�H���g�̓[���ŁC�[���ȉ��̎��͎�M�o�b�t�@�̃f�[�^�����ׂĎ擾���܂��D
Public BinaryBytes As Long

'-----------------------------------
'   Xerror�v���p�e�B
'-----------------------------------
'�G���[���[�h���w�肷��v���p�e�B
'�[��(�K��j�̎���EasyComm�W���̃G���[���[�h�ŁC�G���[����������ƒ�~���܂��D
'Version1.4����́C���̃��[�h�ŃG���[����������Ƃ��ׂẴ|�[�g�����悤�Ɏd�l�ύX�D
'�P�̎��̓g���b�v�\�ȃG���[�𔭐��D
'�������C�P�ɐݒ肵�ăG���[�����������Ƃ��C�g���b�v���[�e�B�����Ȃ��ƃv���O�������I���C�ĕϐ���
'���Z�b�g����̂ŁC���łɊJ���ꂽ�|�[�g����邱�Ƃ��ł��Ȃ��Ȃ邱�Ƃ�����܂��D
'�Q�̎��́C�G���[�𖳎����܂��D
'Public�錾���邱�Ƃɂ���āC�ǂݏ����\�ȃv���p�e�B�ɂ��Ă��܂��D
Public Xerror As Long

'�f���~�^�w�蕶����
Type DelimType
    Cr      As String   ' CR
    Lf      As String   ' LF
    CrLf    As String   ' CR + LF
    LfCr    As String   ' LF + CR
End Type

'�n���h�V�F�[�L���O�w�蕶����
Type HandShakingType
    No      As String   ' �Ȃ�
    XonXoff As String   ' Xon/Off
    RTSCTS  As String   ' RTS/CTS
    DTRDSR  As String   ' DTR/DSR
End Type

'===================================
'   �v���p�e�B�C���\�b�h�̒�`
'===================================

'-----------------------------------
'   InBufferClear���\�b�h
'-----------------------------------
'COMn�v���p�e�B�Ŏw�肳��Ă���|�[�g�̎�M�o�b�t�@�̃f�[�^���N���A���܂�.
'Ascii�v���p�e�B�̓ǂݏo���ő�p���Ă��܂������CAsciiBytes�v���p�e�B�̒ǉ��ɔ����CInBufferClear���\�b�h��ǉ����܂����D
Public Sub InBufferClear()
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10280, _
                    Description:="InBufferClear " & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Sub
        End Select
    End If
    If ecDef.PurgeComm(ecH(Cn).Handle, ecDef.PURGE_RXCLEAR) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ��M�o�b�t�@�̃N���A�Ɏ��s���܂����D
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10281, _
                    Description:="InBufferClear " & Chr$(&HA) & "��M�o�b�t�@�̃N���A�Ɏ��s���܂���"
                Exit Sub
        End Select
    End If

End Sub

'-----------------------------------
'   COMn�v���p�e�B(�ǂݏo���C��������)
'-----------------------------------
'����̑ΏۂƂȂ��|�[�g�ԍ����w��C�܂��͎擾���܂��D
'COMn�v���p�e�B�Ƀ[���ȉ��̐�����������ƁC���ׂẴ|�[�g����܂��D
'ec.COMn = 0 �� ec.COMnClose = 0 �́C������������܂��D
'�v���O�����̏I�����ɁC�����ꂩ�̃R�[�h�ł��ׂẴ|�[�g����Ă��������D

'�ǂݏo���v���V�[�W��
Public Property Get COMn() As Long
    COMn = Cn
End Property

'�������݃v���V�[�W��
Public Property Let COMn(PortNumber As Long)
    Dim rv As Long
    Dim PortName As String
    Dim SettingFlag As Boolean          ' �W���ݒ�p�t���O
    Dim i As Integer
    Dim Fnumb As Integer                ' �L�^�t�@�C���̃t�@�C���ԍ�
    Dim RxBuffer As Long                ' ��M�o�b�t�@�T�C�Y�̐ݒ�p�ϐ�
    Dim TxBuffer As Long                ' ���M�o�b�t�@�T�C�Y�̐ݒ�p�ϐ�
    Dim FileNumber As Integer           ' �I�[�v���\�ȃt�@�C���ԍ�
    Dim Handle As Long
    Dim Fpath As String * 260
    FileNumber = FreeFile()

    If PortNumber <= 0 Then
        '���ׂẴ|�[�g����܂�
        ec.CloseAll

      ElseIf PortNumber > ecMaxPort Then

        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �w�肳�ꂽ�|�[�g�ԍ��͋��e�͈͂ɂ���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10010, _
                    Description:="COMn - Write" & Chr$(&HA) & "�w�肳�ꂽ�|�[�g�ԍ�(" & PortNumber & ")�͋��e�͈͂ɂ���܂���"
                End

        End Select
        
      Else
        
        '���ɊJ����Ă��邩�ǂ����`�F�b�N
        If ecH(PortNumber).Handle > 0 Then
            '���ɊJ����Ă���
            Cn = PortNumber     '�����Ώۂ̃|�[�g�������ɐݒ�
            Exit Property       '����
        End If

        '�|�[�g�I�[�v������
        
        PortName = "\\.\COM" & Trim(Str(PortNumber))
        rv = CreateFile(PortName, GENERIC_READ Or GENERIC_WRITE, _
                        0&, 0&, OPEN_EXISTING, 0&, 0&)

        GetTempPath 260, Fpath
        If rv <> INVALID_HANDLE_VALUE Then
            ' �L�^
            Open Left(Fpath, InStr(Fpath, vbNullChar) - 1) & "ecPort.tmp" For Random Access Read Write As #FileNumber Len = Len(Handle)
            Put #FileNumber, PortNumber, rv
            Close #FileNumber
            
            ecH(PortNumber).Handle = rv     ' �n���h�����X�g�̍X�V
        End If
        Cn = PortNumber                     ' �A�N�e�B�u�|�[�g�ԍ��̐ݒ�
        ' ���s�����Ƃ��͋L�^����Ă���n���h�����g���ă|�[�g����Ă���Ē���
        If rv = INVALID_HANDLE_VALUE Then
            Open Left(Fpath, InStr(Fpath, vbNullChar) - 1) & "ecPort.tmp" For Random Access Read Write As #FileNumber Len = Len(Handle)
            Get #FileNumber, PortNumber, Handle
            If Handle > 0 Then
                ecDef.CloseHandle Handle
                rv = CreateFile(PortName, GENERIC_READ Or GENERIC_WRITE, _
                                0&, 0&, OPEN_EXISTING, 0&, 0&)
                If rv <> INVALID_HANDLE_VALUE Then
                    ' �L�^
                    Put #FileNumber, PortNumber, rv
                    ecH(PortNumber).Handle = rv     ' �n���h�����X�g�̍X�V
                  Else
                    rv = 0&
                    Put #FileNumber, PortNumber, rv
                    ecH(PortNumber).Handle = 0      ' �n���h�����X�g�̍X�V
                End If
            End If
            Close #FileNumber
        End If

        If rv = INVALID_HANDLE_VALUE Then
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' �w�肳�ꂽ�|�[�g�ԍ����J���܂���ł���
                    End             ' �v���O�������I�����܂��D

                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10011, _
                        Description:="COMn - Write" & Chr$(&HA) & "�w�肳�ꂽ�|�[�g�ԍ�(" & PortNumber & ")���J���܂���ł���"
                    End
            
            End Select
        End If

        '--------------------------
        '�J�����|�[�g�̕W���ݒ�
        SettingFlag = True
        
        '�|�[�g�̋@�\���擾
        If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' �|�[�g�̏����擾�ł��܂���ł���
                    End             ' �v���O�������I�����܂��D
    
                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10013, _
                        Description:="InCount - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̏����擾�ł��܂���ł���"
                    Exit Property
            
            End Select
        End If

        ' ��M�o�b�t�@�T�C�Y�̐ݒ�
        RxBuffer = ecDef.fCOMMPROP.dwMaxRxQueue
        If RxBuffer = 0 Then    ' �������̎�
            RxBuffer = ecDef.ecInBufferSize     ' �K��̍ő�l�ɐ�������
        End If

        ' ���M�o�b�t�@�T�C�Y�̐ݒ�
        TxBuffer = ecDef.fCOMMPROP.dwMaxTxQueue
        If TxBuffer = 0 Then    ' �������̎�
            TxBuffer = ecDef.ecOutBufferSize    ' �K��̍ő�l�ɐ�������
        End If

        ' ���s
        If SetupComm(ecH(PortNumber).Handle, RxBuffer, TxBuffer) = False Then
            SettingFlag = False
        End If

        ' DCB�̃Z�b�g
        ' �W���ݒ�l�͈�ʓI�Ȃ��̂Ȃ̂ŁC�T�|�[�g�̗L�����m�F���Ȃ��D
        If GetCommState(ecH(PortNumber).Handle, fDCB) <> False Then
            If BuildCommDCB(ecSetting, fDCB) <> False Then
                fDCB.fBitFields = &H1011
                fDCB.XonLim = ecXonLim
                fDCB.XoffLim = ecXoffLim
                If SetCommState(ecH(PortNumber).Handle, fDCB) = False Then
                    SettingFlag = False
                End If
              Else
                SettingFlag = False
            End If
          Else
            SettingFlag = False
        End If
        
        '�o�b�t�@�̃N���A
        If PurgeComm(ecH(PortNumber).Handle, PURGE_TXCLEAR) = False Then
            SettingFlag = False
        End If
        If PurgeComm(ecH(PortNumber).Handle, PURGE_RXCLEAR) = False Then
            SettingFlag = False
        End If
    
        '�^�C���A�E�g�̐ݒ�
        If GetCommTimeouts(ecH(PortNumber).Handle, fCOMMTIMEOUTS) <> False Then
            '�^�C���A�E�g�l�̕W���ݒ�
            With fCOMMTIMEOUTS
                .ReadIntervalTimeout = ecReadIntervalTimeout
                .ReadTotalTimeoutConstant = ecReadTotalTimeoutConstant
                .ReadTotalTimeoutMultiplier = ecReadTotalTimeoutMultiplier
                .WriteTotalTimeoutConstant = ecWriteTotalTimeoutConstant
                .WriteTotalTimeoutMultiplier = ecWriteTotalTimeoutMultiplier
            End With
            If SetCommTimeouts(ecH(PortNumber).Handle, fCOMMTIMEOUTS) = False Then
                SettingFlag = False
            End If
          Else
            SettingFlag = False
        End If
        If SettingFlag = False Then
            
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' �|�[�g�̕W���ݒ肪�o���܂���ł���
                    End             ' �v���O�������I�����܂��D
                
                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10012, _
                        Description:="COMn - Write" & Chr$(&HA) & "�w�肳�ꂽ�|�[�g�ԍ�(" & PortNumber & ")�̕W���ݒ肪�o���܂���ł���"
                        Exit Property
            
            End Select
            
        End If
    End If
    
End Property

'-----------------------------------
'   COMnClose�v���p�e�B
'-----------------------------------
'�w�肵���ԍ��̃|�[�g����܂��D
'COMnClose�v���p�e�B�Ƀ[���ȉ��̐�����������ƁC���ׂẴ|�[�g����܂��D
'ec.COMn = 0 �� ec.COMnClose = 0 �́C������������܂��D
'�v���O�����̏I�����ɁC�����ꂩ�̃R�[�h�ł��ׂẴ|�[�g����Ă��������D
Public Property Let COMnClose(PortNumber As Long)
    Dim rv As Long
    Dim SheetObj As Object
    Dim i As Integer
    Dim Fnumb As Integer
    Dim FileNumber As Integer           ' �I�[�v���\�ȃt�@�C���ԍ�
    Dim Handle As Long
    Dim Fpath As String * 260
    FileNumber = FreeFile()

    If PortNumber <= 0 Then
        '���ׂẴ|�[�g����܂�
        '�G���[��Ԃ��܂���
        ec.CloseAll
      Else
        '�w�肳�ꂽ�|�[�g����܂�
        '�G���[��Ԃ��܂��D
        rv = CloseHandle(ecH(PortNumber).Handle)

       '�ŏI����
        If rv = False Then
            '���s
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' �w�肳�ꂽ�|�[�g�ԍ��̃N���[�Y�Ɏ��s���܂���
                    End             ' �v���O�������I�����܂��D
                
                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10020, _
                        Description:="COMnClose - Write" & Chr$(&HA) & "�w�肳�ꂽ�|�[�g�ԍ�(" & PortNumber & ")�̃N���[�Y�Ɏ��s���܂���"
            
            End Select
          
          Else
            '����
            ecH(PortNumber).Handle = 0          '�n���h�����X�g�̍X�V
            ecH(PortNumber).Delimiter = ""      '�W���f���~�^
            ecH(PortNumber).LineInTimeOut = 0   'AsciiLineTimeOut�̃��Z�b�g
            GetTempPath 260, Fpath
            Open Left(Fpath, InStr(Fpath, vbNullChar) - 1) & "ecPort.tmp" For Random Access Read Write As #FileNumber Len = Len(Handle)
            rv = 0&
            Put #FileNumber, PortNumber, rv
            Close #FileNumber
        End If
    End If
End Property

'-----------------------------------
'   Setting�v���p�e�B(�������ݓǂݏo��)
'-----------------------------------
'����̑ΏۂƂȂ��Ă���|�[�g�̒ʐM�����̐ݒ�C�܂��͓ǂݏo�����s�Ȃ��܂��D

'�ǂݏo��
Public Property Get Setting() As String

    Dim ModeStr As String

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10038, _
                    Description:="Setting - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
    
    End If

    If GetCommState(ecH(Cn).Handle, fDCB) = False Then     'DCB�̓ǂݏo��
        
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���ݐݒ�l(DCB)�̓ǂݍ��݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10030, _
                    Description:="Setting - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̌��ݐݒ�l(DCB)�̓ǂݍ��݂Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If

    ModeStr = "baud=" & Trim(Str(fDCB.BaudRate))
    ModeStr = ModeStr & " parity="
    If fDCB.fBitFields And &H2 <> 0 Then
        Select Case fDCB.Parity
            Case Is = NOPARITY
                ModeStr = ModeStr & "N"
            Case Is = ODDPARITY
                ModeStr = ModeStr & "O"
            Case Is = EVENPARITY
                ModeStr = ModeStr & "E"
            Case Is = MARKPARITY
                ModeStr = ModeStr & "M"
            Case Is = SPACEPARITY
                ModeStr = ModeStr & "S"
        End Select
      Else
        ModeStr = ModeStr & "N"
    End If
    
    ModeStr = ModeStr & " data=" & Trim(Str(fDCB.ByteSize))

    ModeStr = ModeStr & " stop="
    Select Case fDCB.StopBits
        Case Is = ONESTOPBIT
            ModeStr = ModeStr & "1"
        Case Is = ONE5STOPBITS
            ModeStr = ModeStr & "1.5"
        Case Is = TWOSTOPBITS
            ModeStr = ModeStr & "2"
    End Select
    
    Setting = ModeStr

End Property

'��������
Public Property Let Setting(Mode As String)
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10039, _
                    Description:="Setting - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
        
    End If

    ' DCB�̓ǂݏo��
    If GetCommState(ecH(Cn).Handle, fDCB) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���ݐݒ�l(DCB)�̓ǂݍ��݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10031, _
                    Description:="Setting - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̌��ݐݒ�l(DCB)�̓ǂݍ��݂Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If

    '������̕ϊ�
    If BuildCommDCB(Mode, fDCB) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ݒ蕶����̉�͂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10032, _
                    Description:="Setting - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̐ݒ蕶����(" & Mode & ")�̉�͂Ɏ��s���܂���"
                Exit Property
        
        End Select

    End If
    
    'DCB�̏�������
    If SetCommState(ecH(Cn).Handle, fDCB) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���ݐݒ�l(DCB)�̏������݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10033, _
                    Description:="Setting - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̌��ݐݒ�l(DCB)�̏������݂Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If
    
    '�o�b�t�@�̃N���A
    '�����܂Ńp�X���Ă��Ȃ���o�b�t�@�̃N���A�ŃG���[����������\���͂Ȃ��Ǝv����̂ŁC
    '�G���[�`�F�b�N�͍s���܂���D
    PurgeComm ecH(Cn).Handle, PURGE_TXCLEAR + PURGE_RXCLEAR

End Property


'-----------------------------------
'   HandShaking�v���p�e�B(�ǂݏo����������)
'-----------------------------------
'COMn�v���p�e�B�Ŏw�肳��Ă���|�[�g�̃t���[���������ݒ�C�܂��͎擾���܂�

'�ǂݏo��
Public Property Get HandShaking() As String
    Dim Flow As String
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10188, _
                    Description:="HandShaking - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If
    
    ' DCB�̓ǂݏo��
    If GetCommState(ecH(Cn).Handle, fDCB) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���ݐݒ�l(DCB)�̓ǂݍ��݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10180, _
                    Description:="HandShaking - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̏�Ԃ��擾�ł��܂���ł����D"
                Exit Property
        
        End Select

    End If
    
    Flow = ""

    'RTS�n���h�V�F�[�N�̃`�F�b�N
    If (fDCB.fBitFields And &H2000) <> 0 Then
        Flow = ec.HANDSHAKEs.RTSCTS
    End If

    'XON/OFF�n���h�V�F�[�N�̃`�F�b�N
    If (fDCB.fBitFields And &H300) <> 0 Then
        Flow = Flow & ec.HANDSHAKEs.XonXoff
    End If
    
    'DTR/DSR�n���h�V�F�[�N�̃`�F�b�N
    If (fDCB.fBitFields And &H20) <> 0 Then
        Flow = Flow & ec.HANDSHAKEs.DTRDSR
    End If
    
    If Flow = "" Then
        Flow = ec.HANDSHAKEs.No
    End If

    HandShaking = Flow

End Property

'��������
Public Property Let HandShaking(Flow As String)
    Dim f As String

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10189, _
                    Description:="HandShaking - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
        
    End If
    
    f = Left(UCase(Flow), 1)        '�擪�P�����̂ݗL��
    
    ' DCB�̓ǂݏo��
    If GetCommState(ecH(Cn).Handle, fDCB) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���ݐݒ�l(DCB)�̓ǂݍ��݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10181, _
                    Description:="HandShaking - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̏�Ԃ��擾�ł��܂���ł����D"
                Exit Property
        
        End Select
        
    End If

    With fDCB
        .fBitFields = .fBitFields And &HCC83            '�Y���r�b�g�̃��Z�b�g
        Select Case f
            Case Is = "N"
                '�t���[�Ȃ�
                .fBitFields = .fBitFields Or &H1010     '�Y���r�b�g�̃Z�b�g
            Case Is = "X"
                'Xon/Off�ɂ��n���h�V�F�C�N
                .fBitFields = .fBitFields Or &H1310     '�Y���r�b�g�̃Z�b�g
            Case Is = "R"
                'RTS�ɂ��n���h�V�F�[�N
                .fBitFields = .fBitFields Or &H2014     '�Y���r�b�g�̃Z�b�g
            Case Is = "D"
                'DTR�ɂ��n���h�V�F�[�N
                .fBitFields = .fBitFields Or &H1068     '�Y���r�b�g�̃Z�b�g
        End Select
    End With

    ' DCB�̏�������
    If SetCommState(ecH(Cn).Handle, fDCB) = False Then

        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �n���h�V�F�[�N���������߂܂���ł���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10182, _
                    Description:="HandShaking - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̃n���h�V�F�[�N���������߂܂���ł����D"
                Exit Property
        
        End Select
    End If
End Property

'-----------------------------------
'   InBuffer�v���p�e�B(�ǂݏo����������)
'-----------------------------------
'COMn�v���p�e�B�Ŏw�肳��Ă���|�[�g�̎�M�o�b�t�@�̏�Ԃ�ݒ�C�܂��͎擾���܂��D

'�ǂݏo��
'��M�o�b�t�@�ɂ��܂�����M�f�[�^�̃o�C�g�����擾���܂��D
Public Property Get InBuffer() As Long
    Dim Er As Long
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10068, _
                    Description:="InBuffer - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If
    
    If ClearCommError(ecH(Cn).Handle, Er, fCOMSTAT) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ��M�o�b�t�@�̃f�[�^�o�C�g���̓ǂݎ��Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10060, _
                    Description:="InBuffer - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̎�M�o�b�t�@�̃f�[�^�o�C�g���̓ǂݎ��Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If
    InBuffer = fCOMSTAT.cbInQue
End Property

'�������ݏ���
'��M�o�b�t�@�T�C�Y���w�肵�܂��D
'�ŏ��ݒ�l��ecMinimumInBuffer�Ŏw�肳��܂��D
'ecMinimumInBuffer�́CecDef���Œ�`����Ă��܂��D
Public Property Let InBuffer(BufferSize As Long)
    Dim InBuff As Long      '���݂̎�M�o�b�t�@�T�C�Y
    Dim OutBuff As Long     '���݂̎�M�o�b�t�@�T�C�Y

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10069, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If

    '�w��T�C�Y�����̃`�F�b�N
    If BufferSize < ecMinimumInBuffer Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �o�b�t�@�̒l�����Ȃ����܂��D
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10065, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "�o�b�t�@�� " & ecMinimumInBuffer & "�ȏ�ɐݒ肵�Ă�������"
                Exit Property
        
        End Select
        
    End If
    
    '�|�[�g�̐ݒ�l�̎擾
    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ��M�o�b�t�@�T�C�Y�̓ǂݎ��Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10061, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̎�M�o�b�t�@�T�C�Y�̓ǂݎ��Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If

    '����l�̃`�F�b�N
    If fCOMMPROP.dwMaxRxQueue <> 0 Then     '�[���̎��͏���Ȃ�
        '����l����
        If BufferSize > fCOMMPROP.dwMaxRxQueue Then
            '����𒴂����ݒ�
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' �|�[�g�̍ő�o�b�t�@�T�C�Y���z�����ݒ�����悤�Ƃ��܂����D
                    End             ' �v���O�������I�����܂��D
    
                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10066, _
                        Description:="InBuffer - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̎�M�o�b�t�@�̏���l (" & fCOMMPROP.dwMaxRxQueue & "���z����ݒ�͏o���܂���D"
                    Exit Property
            
            End Select
        End If
    End If
    
    InBuff = fCOMMPROP.dwCurrentRxQueue     '��M�o�b�t�@�̐ݒ��ǂݏo��
    OutBuff = fCOMMPROP.dwCurrentTxQueue    '���M�o�b�t�@�̐ݒ��ǂݏo��
    
    '�V�����T�C�Y�̐ݒ�
    If SetupComm(ecH(Cn).Handle, BufferSize, OutBuff) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ��M�o�b�t�@�T�C�Y�̏������݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10062, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̎�M�o�b�t�@�T�C�Y�̏������݂Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If
        
    ' DCB�̓ǂݏo��
    If GetCommState(ecH(Cn).Handle, fDCB) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���ݐݒ�l(DCB)�̓ǂݍ��݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10063, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̌��ݐݒ�l(DCB)�̓ǂݍ��݂Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If
    
    If SetCommState(ecH(Cn).Handle, fDCB) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���ݐݒ�l(DCB)�̏������݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10064, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̐ݒ�l(DCB)�̏������݂Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If

End Property

'-----------------------------------
'   OutBuffer�v���p�e�B(�������ݐ�p)
'-----------------------------------
'COMn�v���p�e�B�Ŏw�肳��Ă���|�[�g�̑��M�o�b�t�@�̏�Ԃ�ݒ�C�܂��͎擾���܂��D
'�ŏ��ݒ�l��ecMinimumOutBuffer�Ŏw�肳��܂��D
'ecMinimumOutBuffer�́CecDef���Œ�`����Ă��܂��D
Public Property Let OutBuffer(BufferSize As Long)
    
    Dim InBuff As Long      '���݂̎�M�o�b�t�@�T�C�Y
    Dim OutBuff As Long     '���݂̑��M�o�b�t�@�T�C�Y

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10079, _
                    Description:="OutBuffer - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If
    
    '�w��T�C�Y�̃`�F�b�N
    If BufferSize < ecMinimumOutBuffer Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���M�o�b�t�@�̐ݒ�l�����������܂�
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10073, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "�o�b�t�@�� " & ecMinimumOutBuffer & "�ȏ�ɐݒ肵�Ă�������"
                Exit Property
        
        End Select
        
    End If
    
    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���M�o�b�t�@�T�C�Y�̓ǂݎ��Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10071, _
                    Description:="OutBuffer - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̑��M�o�b�t�@�T�C�Y�̓ǂݎ��Ɏ��s���܂���"
                Exit Property
        
        End Select
    End If
    
    '����l�̃`�F�b�N
    If fCOMMPROP.dwMaxTxQueue <> 0 Then     '�[���̎��͏���Ȃ�
        '����l����
        If BufferSize > fCOMMPROP.dwMaxTxQueue Then
            '����𒴂����ݒ�
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' �|�[�g�̍ő�o�b�t�@�T�C�Y���z�����ݒ�����悤�Ƃ��܂����D
                    End             ' �v���O�������I�����܂��D
    
                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10076, _
                        Description:="InBuffer - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̎�M�o�b�t�@�̏���l (" & fCOMMPROP.dwMaxRxQueue & "���z����ݒ�͏o���܂���D"
                    Exit Property
            
            End Select
        End If
    End If
    
    InBuff = fCOMMPROP.dwCurrentRxQueue     '��M�o�b�t�@�̐ݒ��ǂݏo��
    OutBuff = fCOMMPROP.dwCurrentTxQueue    '���M�o�b�t�@�̐ݒ��ǂݏo��
    
    '�V�����T�C�Y�̐ݒ�
    If SetupComm(ecH(Cn).Handle, InBuff, BufferSize) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ���M�o�b�t�@�T�C�Y�̏������݂Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10072, _
                    Description:="OutBuffer - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̑��M�o�b�t�@�T�C�Y�̏������݂Ɏ��s���܂���"
                Exit Property
        
        End Select
        
    End If

End Property

'-----------------------------------
'   Ascii�v���p�e�B
'-----------------------------------
'�ǂݏo��
Public Property Get Ascii() As String
    Dim ReadBytes As Long       '��M�o�C�g��
    Dim ReadedBytes As Long     '�ǂݍ��߂��o�C�g��
    Dim bdata() As Byte         '�o�C�i���z��
    Dim Er As Long              ' error value
    Dim ErrorFlag As Boolean    '�G���[�t���O

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10108, _
                    Description:="Ascii - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
    End If

    If ClearCommError(ecH(Cn).Handle, Er, fCOMSTAT) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ��M�o�b�t�@�̏�Ԃ��擾�ł��܂���ł���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10100, _
                    Description:="Ascii - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̎�M�o�b�t�@�̏�Ԃ��擾�ł��܂���ł����D"
                Exit Property
        
        End Select
    
    End If
    
    '--------- Version1.70�ŕύX
'    ReadBytes = fCOMSTAT.cbInQue       '��M�o�C�g���̎擾
    If ec.AsciiBytes <= 0 Then
        ' ���o�[�W�����݊�
        ReadBytes = fCOMSTAT.cbInQue    '��M�o�C�g���̎擾
      Else
        '--------- Version1.71�ŏC��
'        ReadBytes = ec.AsciiBytes
        If ec.AsciiBytes <= fCOMSTAT.cbInQue Then
            ReadBytes = ec.AsciiBytes
          Else
            ReadBytes = fCOMSTAT.cbInQue    '��M�o�C�g���̎擾
        End If
        '--------- �����܂�
    End If
    '--------- �����܂�


    If ReadBytes = 0 Then       '�f�[�^�o�b�t�@����̂Ƃ�
        Ascii = ""
        Exit Property
    End If

    ReDim bdata(ReadBytes - 1)
    
    ErrorFlag = False
    
    If ReadFile(ecH(Cn).Handle, bdata(0), ReadBytes, ReadedBytes, 0&) = False Then ErrorFlag = True
    If ReadBytes <> ReadedBytes Then ErrorFlag = True
    If ErrorFlag Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �������M�Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10101, _
                    Description:="Ascii - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")����̕������M�Ɏ��s���܂����D"
                Exit Property
        
        End Select
    End If
    Ascii = StrConv(bdata, vbUnicode)

End Property

'��������
Public Property Let Ascii(TxD As String)
    Dim WriteBytes As Long      '���M�o�C�g��
    Dim WrittenBytes As Long    '�o�b�t�@�ɏ������߂��o�C�g��
    Dim bdata() As Byte         '�o�C�i���z��
    Dim ErrorFlag As Boolean    '�G���[�t���O
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10109, _
                    Description:="Ascii - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If
    
    bdata() = StrConv(TxD, vbFromUnicode)   'ANSI�ɕϊ�
    WriteBytes = UBound(bdata()) + 1
    
    ErrorFlag = False

    If WriteFile(ecH(Cn).Handle, bdata(0), WriteBytes, WrittenBytes, 0&) = False Then
        ErrorFlag = True
    End If

    If WriteBytes <> WrittenBytes Then ErrorFlag = True
    If ErrorFlag Then
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �����񑗐M�Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10102, _
                    Description:="Ascii - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")����̕����񑗐M�Ɏ��s���܂����D"
                Exit Property
        
        End Select
    End If
End Property

'-----------------------------------
'   AsciiLineTimeOut�v���p�e�B
'-----------------------------------
'Version1.51�Œǉ�
'AsciiLine�v���p�e�B�̓ǂݏo�����̃^�C���A�E�g��ݒ�C�܂��͐ݒ�l��ǂݏo���܂��D
'AsciiLine�v���p�e�B��ǂݏo�������_����v�����CAsciiLineTimeOut(mS)���z���Ă��f��
'�~�^����M�ł��Ȃ������Ƃ��͏����𒆎~���C����܂łɎ�M�������������̂܂ܕԂ��܂��D
'�G���[�͔������܂���D
'AsciiLineTimeOut�v���p�e�B�̓|�[�g���Ƃɐݒ肵�܂��D
'�����l�̓[���ł����C�[���ȉ��̒l���ݒ肳��Ă���ƃ^�C���A�E�g�͔������܂���D
'��������
Public Property Let AsciiLineTimeOut(TimeOut As Long)
    If Cn = 0 Then Exit Property
    ecH(Cn).LineInTimeOut = TimeOut
End Property
'�ǂݏo��
Public Property Get AsciiLineTimeOut() As Long
    If Cn = 0 Then
        AsciiLineTimeOut = 0
        Exit Property
      Else
        AsciiLineTimeOut = ecH(Cn).LineInTimeOut
    End If
End Property

'-----------------------------------
'   AsciiLine�v���p�e�B
'-----------------------------------
'�f���~�^�܂ł̕��������M
'�f���~�^��Delimiter�v���p�e�B�ŁC�|�[�g���Ƃɐݒ�E�ǂݏo�����\�ł��D
'Version1.51���C�ǂݏo���^�C���A�E�g���T�|�[�g���܂����D
'�^�C���A�E�g�̓|�[�g���Ƃɐݒ肷��K�v������܂��D
'AsciiLine�v���p�e�B��ǂݏo�����Ƃ�����C�f���~�^����M����܂ł̎��Ԃ��ݒ�l(mS)��
'�z����ƃ^�C���A�E�g���������C�������I�����܂��D
'�������^�C���A�E�g���������Ă��G���[�ɂ͂Ȃ炸�C�����܂łɓǂݍ��񂾕���������̂܂�
'�Ԃ��܂��D
'�^�C���A�E�g�l��AsciiLine�̓ǂݏo�����̂ݗL���ł��D
'
'�ǂݏo��
Public Property Get AsciiLine() As String
    Dim ReadBytes As Long       '��M�o�C�g��
    Dim ReadedBytes As Long     '�ǂݍ��߂��o�C�g��
    Dim bdata As Byte           '�o�C�g�ϐ�
    Dim Bstr() As Byte          '�o�C�g�z��
    Dim Er As Long              ' error value
    Dim n As Long               '�������J�E���g
    Dim DelimStr As String      '�f���~�^
    Dim ErrorFlag As Boolean    '�G���[�t���O
    
    '----Version 1.51
    Dim TimeOutFlag As Boolean  ' �^�C���A�E�g�t���O
    Dim STARTmS As Double       ' �J�n����(mS)
    Dim NOWmS As Double         ' ���݂̎���(mS)

    STARTmS = GetTickCount      ' �J�n���̎��Ԃ��擾
    If STARTmS < 0 Then
        STARTmS = STARTmS + 4294967296#
    End If
    TimeOutFlag = False
    '----�����܂�

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10218, _
                    Description:="AsciiLine - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If

    ' �ݒ�l�̐��K��
    DelimStr = Trim(UCase(StrConv(ecH(Cn).Delimiter, vbNarrow)))
    Select Case DelimStr
        Case Is = "CR"
            ecH(Cn).Delimiter = "CR"
        Case Is = "LF"
            ecH(Cn).Delimiter = "LF"
        Case Is = "CRLF"
            ecH(Cn).Delimiter = "CRLF"
        Case Is = "LFCR"
            ecH(Cn).Delimiter = "LFCR"
        Case Else
            ecH(Cn).Delimiter = "CR"    ' �K��l�ȊO�͂b�q�Ƃ݂Ȃ��܂��D
            DelimStr = "CR"
    End Select


    n = 0   ' �����J�E���^�̃��Z�b�g

    Do
        '�f�[�^��M�܂ő҂�
        Do
            If ClearCommError(ecH(Cn).Handle, Er, fCOMSTAT) = False Then
                '�G���[����
                Select Case Xerror
                    Case Is = 0 '-----�W���G���[
                        ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                        Stop            ' ��M�o�b�t�@�̏�Ԃ��擾�ł��܂���ł���
                        End             ' �v���O�������I�����܂��D

                    Case Is = 1 '-----�g���b�v�\�G���[
                        err.Raise _
                            Number:=10210, _
                            Description:="AsciiLine - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̎�M�o�b�t�@�̏�Ԃ��擾�ł��܂���ł����D"
                        Exit Property
                
                End Select
                
            End If
        
            ReadBytes = fCOMSTAT.cbInQue
            DoEvents

            If ecH(Cn).LineInTimeOut > 0 Then   ' �^�C���A�E�g���[���ȉ��Ȃ�΃X�L�b�v
                ' ���݂̎��Ԃ��擾
                NOWmS = GetTickCount
                If NOWmS < 0 Then
                    NOWmS = NOWmS + 4294967296#
                End If
                '�^�C���A�E�g�̃`�F�b�N
                If STARTmS + ecH(Cn).LineInTimeOut <= NOWmS Then
                    ' �^�C���A�E�g����
                    TimeOutFlag = True
                    Exit Do
                End If
            End If
        Loop While ReadBytes = 0

        If TimeOutFlag Then Exit Do

        '�P������M
        ErrorFlag = False
        
        If ReadFile(ecH(Cn).Handle, bdata, 1&, ReadedBytes, 0&) = False Then ErrorFlag = True
        If ReadedBytes <> 1 Then ErrorFlag = True
        If ErrorFlag Then
            '���s
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' �������M�Ɏ��s���܂���
                    End             ' �v���O�������I�����܂��D

                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10211, _
                        Description:="AsciiLine - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")����̕������M�Ɏ��s���܂����D"
                    Exit Property
            
            End Select
        End If

        '��M�����̉��
        Select Case bdata
            Case Is = &HD   'Cr����M
                Select Case DelimStr
                    Case Is = "CR"
                        Exit Do                     ' �f���~�^����M
                    Case Is = "LF", "CRLF"
                        ReDim Preserve Bstr(n)      ' �ȑO�̃f�[�^���c�����܂܍Ē�`
                        Bstr(n) = bdata             ' ��M�����ɉ�����
                        n = n + 1
                    Case Is = "LFCR"
                        If n >= 1 Then
                            If Bstr(n - 1) = &HA Then   ' ��O�̕�����Lf?
                                n = n - 1               ' �f���~�^�����𖳌���
                                Exit Do                 ' �f���~�^����M
                            End If
                          Else
                            ReDim Preserve Bstr(n)  ' �ȑO�̃f�[�^���c�����܂܍Ē�`
                            Bstr(n) = bdata         ' ��M�����ɉ�����
                            n = n + 1
                        End If
                End Select
            
            Case Is = &HA   'Lf����M
                
                Select Case DelimStr
                    Case Is = "CR", "LFCR"
                        ReDim Preserve Bstr(n)      ' �ȑO�̃f�[�^���c�����܂܍Ē�`
                        Bstr(n) = bdata             ' ��M�����ɉ�����
                        n = n + 1
                    Case Is = "LF"                  ' �f���~�^����M
                        Exit Do
                    Case Is = "CRLF"
                        If n >= 1 Then
                            If Bstr(n - 1) = &HD Then   ' ��O�̕�����Cr?
                                n = n - 1               ' �f���~�^�����𖳌���
                                Exit Do                 ' �f���~�^����M
                            End If
                          Else
                            ReDim Preserve Bstr(n)  ' �ȑO�̃f�[�^���c�����܂܍Ē�`
                            Bstr(n) = bdata         ' ��M�����ɉ�����
                            n = n + 1
                        End If
                End Select
                
            Case Else   '�ʏ�̕���
                ReDim Preserve Bstr(n)          ' �ȑO�̃f�[�^���c�����܂܍Ē�`
                Bstr(n) = bdata
                n = n + 1
        End Select
    Loop

    If n > 0 Then
        ReDim Preserve Bstr(n - 1)
        AsciiLine = StrConv(Bstr, vbUnicode)
      Else
        AsciiLine = ""
    End If

End Property

'�w�肳�ꂽ������Ƀf���~�^��t�����đ��M
'��������
Public Property Let AsciiLine(TxD As String)
    Dim WriteBytes As Long      '���M�o�C�g��
    Dim WrittenBytes As Long    '�o�b�t�@�ɏ������߂��o�C�g��
    Dim bdata() As Byte         '�o�C�i���z��
    Dim DelimStr As String      '�f���~�^
    Dim Td As String            '���M������
    Dim ErrorFlag As Boolean    '�G���[�t���O
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10219, _
                    Description:="AsciiLine - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If

    ' �ݒ�l�̐��K���ƕ�����Ƀf���~�^��t��
    
    DelimStr = Trim(UCase(StrConv(ecH(Cn).Delimiter, vbNarrow)))
    Select Case DelimStr
        Case Is = "CR"
            ecH(Cn).Delimiter = "CR"
            Td = TxD & Chr(&HD)
        Case Is = "LF"
            ecH(Cn).Delimiter = "LF"
            Td = TxD & Chr(&HA)
        Case Is = "CRLF"
            ecH(Cn).Delimiter = "CRLF"
            Td = TxD & Chr(&HD) & Chr(&HA)
        Case Is = "LFCR"
            ecH(Cn).Delimiter = "LFCR"
            Td = TxD & Chr(&HA) & Chr(&HD)
        Case Else
            ecH(Cn).Delimiter = "CR"        ' �K��l�ȊO�͂b�q�Ƃ݂Ȃ��܂��D
            Td = TxD & Chr(&HD)
    End Select
    
    bdata() = StrConv(Td, vbFromUnicode)   'ANSI�ɕϊ�
    
    WriteBytes = UBound(bdata()) + 1
    
    ErrorFlag = False
    
    If WriteFile(ecH(Cn).Handle, bdata(0), WriteBytes, WrittenBytes, 0&) = False Then ErrorFlag = True
    If WriteBytes <> WrittenBytes Then ErrorFlag = True
    If ErrorFlag Then
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �����񑗐M�Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10212, _
                    Description:="AsciiLine - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")����̕����񑗐M�Ɏ��s���܂����D"
                Exit Property
        
        End Select
    End If
End Property


'-----------------------------------
'   Binary�v���p�e�B
'-----------------------------------
'��������
Public Property Let Binary(ByteData As Variant)
    Dim WriteBytes As Long          '���M�o�C�g��
    Dim WrittenBytes As Long        '�ǂݍ��߂��o�C�g��
    Dim bdata() As Byte             '�o�C�i���z��
    Dim Er As Long                  ' error value
    Dim ErrorFlag As Boolean        '�G���[�t���O
    Dim i As Long
    Dim j As Long
    Dim C As Long

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10169, _
                    Description:="Binary - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
        
    End If

    '�����̌^�ɂ�鏈���̕���
    Select Case TypeName(ByteData)
        Case "Byte"
            ReDim bdata(0)
            bdata(0) = ByteData
        Case "Integer", "Long", "Single", "Double"
            If ByteData < 0 Or ByteData > 255 Then
                '�I�[�o�[�t���[�G���[
                Select Case Xerror
                    Case Is = 0 '-----�W���G���[
                        ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                        Stop            ' �����̒l��0�`255�͈̔͂ɂ���܂���D
                        End             ' �v���O�������I�����܂��D
        
                    Case Is = 1 '-----�g���b�v�\�G���[
                        err.Raise _
                            Number:=10162, _
                            Description:="Binary - Write" & Chr$(&HA) & "�����̒l��0�`255�͈̔͂ɂ���܂���"
                        Exit Property

                End Select
            End If
            If ByteData <> Int(ByteData) Then
                ' �񐮐��G���[
                Select Case Xerror
                    Case Is = 0 '-----�W���G���[
                        ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                        Stop            ' �����̒l�������ł͂���܂���D
                        End             ' �v���O�������I�����܂��D
        
                    Case Is = 1 '-----�g���b�v�\�G���[
                        err.Raise _
                            Number:=10163, _
                            Description:="Binary - Write" & Chr$(&HA) & "�����������ł͂���܂���"
                        Exit Property
                
                End Select
            End If
            ReDim bdata(0)
            bdata(0) = CByte(ByteData)
    
        Case "String"
            '������̓��j�R�[�h�̂܂ܑ��M
            ReDim bdata(LenB(ByteData) - 1)
            For i = 0 To LenB(ByteData) - 1
                bdata(i) = AscB(MidB(ByteData, i + 1, 1))
            Next i
        
        Case "Byte()"
            bdata = ByteData
    
        Case "Integer()", "Long()", "Single()", "Double()", "Variant()"
            ReDim bdata(UBound(ByteData))   ' �f�[�^�z��̍Đ錾
            ' �o�C�g�z��ւ̑���ƃI�[�o�[�t���[�C�����̃`�F�b�N
            For i = 0 To UBound(ByteData)
                If ByteData(i) < 0 Or ByteData(i) > 255 Then
                '�I�[�o�[�t���[�G���[
                    Select Case Xerror
                        Case Is = 0 '-----�W���G���[
                            ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                            Stop            ' �����̒l��0�`255�͈̔͂ɂ���܂���D
                            End             ' �v���O�������I�����܂��D
            
                        Case Is = 1 '-----�g���b�v�\�G���[
                            err.Raise _
                                Number:=10162, _
                                Description:="Binary - Write" & Chr$(&HA) & "�����̒l��0�`255�͈̔͂ɂ���܂���"
                            Exit Property
    
                    End Select
                End If
                If ByteData(i) <> Int(ByteData(i)) Then
                    ' �񐮐��G���[
                    Select Case Xerror
                        Case Is = 0 '-----�W���G���[
                            ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                            Stop            ' �����̒l�������ł͂���܂���D
                            End             ' �v���O�������I�����܂��D
            
                        Case Is = 1 '-----�g���b�v�\�G���[
                            err.Raise _
                                Number:=10163, _
                                Description:="Binary - Write" & Chr$(&HA) & "�����������ł͂���܂���"
                            Exit Property
                    
                    End Select
                End If
                ' �f�[�^�̃Z�b�g
                bdata(i) = ByteData(i)
            Next i
        
        Case "String()"
            '�����z��̓��j�R�[�h�̂܂ܑ��M
            C = 0   ' ���݂̃g�[�^��������
            For j = 0 To UBound(ByteData)
                If LenB(ByteData(j)) > 0 Then
                    ReDim Preserve bdata(C + LenB(ByteData(j)) - 1)
                    For i = 0 To LenB(ByteData(j)) - 1
                        bdata(C + i) = AscB(MidB(ByteData(j), i + 1, 1))
                    Next i
                    C = C + LenB(ByteData(j))
                End If
            Next j

        Case Else
            ' ��Ή��^
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' �w�肳�ꂽ�^�ɂ͑Ή����Ă��܂���D
                    End             ' �v���O�������I�����܂��D
    
                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10164, _
                        Description:="Binary - Write" & Chr$(&HA) & "�w�肳�ꂽ�^�ɂ͑Ή����Ă��܂���"
                    Exit Property
            
            End Select
    
    End Select
    
    WriteBytes = UBound(bdata) + 1
    
    ErrorFlag = False
    
    If WriteFile(ecH(Cn).Handle, bdata(0), WriteBytes, WrittenBytes, 0&) = False Then ErrorFlag = True
    If WriteBytes <> WrittenBytes Then ErrorFlag = True
    If ErrorFlag Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �o�C�i���[�f�[�^�̑��M�Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10160, _
                    Description:="Binary - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")����̃o�C�i���[�f�[�^�̑��M�Ɏ��s���܂����D"
                Exit Property
        
        End Select
    End If
End Property

'�ǂݏo��
'Version1.70���CBinaryBytes�v���p�e�B�Ŏ�M�o�b�t�@����ǂݏo���o�C�g�����w��ł���悤�ɂȂ�܂����D
'�������CBytes�v���p�e�B���[���ȉ��̎��́C���ׂẴf�[�^���擾���܂�(���o�[�W�����݊�)
Public Property Get Binary() As Variant
    Dim ReadBytes As Long       '��M�o�C�g��
    Dim ReadedBytes As Long     '�ǂݍ��߂��o�C�g��
    Dim bdata() As Byte         '�o�C�i���z��
    Dim Er As Long              ' error value
    Dim ErrorFlag As Boolean    '�G���[�t���O
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10179, _
                    Description:="Binary - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
        
    End If
    
    ' �|�[�g�̏�Ԃ��擾
    If ClearCommError(ecH(Cn).Handle, Er, fCOMSTAT) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' ��M�o�b�t�@�̏�Ԃ��擾�ł��܂���ł���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10172, _
                    Description:="Binary - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̎�M�o�b�t�@�̏�Ԃ��擾�ł��܂���ł����D"
                Exit Property
        
        End Select
        
    End If

    '--------- Version1.70�ŕύX
'    ReadBytes = fCOMSTAT.cbInQue       '��M�o�C�g���̎擾
    If ec.BinaryBytes <= 0 Then
        ' ���o�[�W�����݊�
        ReadBytes = fCOMSTAT.cbInQue    '��M�o�C�g���̎擾
      Else
        '--------- Version1.71�ŏC��
'        ReadBytes = ec.BinaryBytes
        If ec.AsciiBytes <= fCOMSTAT.cbInQue Then
            ReadBytes = ec.BinaryBytes
          Else
            ReadBytes = fCOMSTAT.cbInQue    '��M�o�C�g���̎擾
        End If
        '--------- �����܂�
    End If
    '--------- �����܂�
    
    If ReadBytes = 0 Then           '�f�[�^�o�b�t�@����̂Ƃ�
        Binary = 0                  ' 0 ��Ԃ�
        Exit Property
    End If
    
    ReDim bdata(ReadBytes - 1)

    ErrorFlag = False

    If ReadFile(ecH(Cn).Handle, bdata(0), ReadBytes, ReadedBytes, 0&) = False Then ErrorFlag = True
    If ReadBytes <> ReadedBytes Then ErrorFlag = True
    If ErrorFlag Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �o�C�i���[�f�[�^�̎�M�Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10171, _
                    Description:="Binary - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")����̃o�C�i���[�f�[�^�̎�M�Ɏ��s���܂����D"
                Exit Property
        
        End Select
    End If
    Binary = bdata
End Property

'-----------------------------------
'   WAITmS �v���p�e�B
'-----------------------------------
'�w�肵�����Ԍ�ɖ߂�܂��D
'�������^(Long)�CmS�P�ʂŎw�肵�܂��D
'�ő�49.7���܂Ŏw��ł��܂��D
Public Property Let WAITmS(WaitTime As Long)
    'mS�P�ʂ̃f�B���C
    Dim STARTmS As Double   ' �J�n����(mS)
    Dim NOWmS As Double     ' ���݂̎���(mS)

    '�J�n���̎��Ԃ��擾
    STARTmS = GetTickCount
    If STARTmS < 0 Then
        STARTmS = STARTmS + 4294967296#
    End If

    '���ԑ҂�
    Do
        DoEvents
        NOWmS = GetTickCount
        If NOWmS < 0 Then
            NOWmS = NOWmS + 4294967296#
        End If
        If STARTmS > NOWmS Then
            NOWmS = NOWmS + 4294967296#
        End If
    Loop While STARTmS + WaitTime > NOWmS    ' GetTickCount

End Property

'-----------------------------------
'   DozeSeconds �v���p�e�B
'-----------------------------------
' �w�肵������(Sec)�C�������~���܂��D
' ��~���鎞�Ԃ͕b�P�ʂő�����܂��D
' Doze�Ƃ́u�������Q�v�Ƃ����Ӗ��ŁC0.1�b���Ƃ�DoEvents�����s����邱�Ƃ��疽�����܂����D
' Excel2000�ȍ~�ł͗L���ł�������ȑO�̃o�[�W�����ł͂قƂ�ǈӖ�������܂���D
Public Property Let DozeSeconds(Seconds As Integer)
    Dim WakeUp As Date                  ' �ڊo�߂̎���
    If Seconds < 1 Then Exit Property   ' 1�b�ȏ�̂ݗL��
    WakeUp = Now + TimeSerial(0, 0, Seconds)
    Do
        DoEvents
        If Now >= WakeUp Then Exit Do
        ecDef.Sleep 100
    Loop
End Property

'-----------------------------------
'   RTSCTS�v���p�e�B
'-----------------------------------
' RTS�̋��������CTS�̏�ԓǂݎ��

'�ǂݏo��(CTS�̏��)
Public Property Get RTSCTS() As Boolean
    Dim Stat As Long    ' Status

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10198, _
                    Description:="RTSCTS - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
    
    End If

    If GetCommModemStatus(ecH(Cn).Handle, Stat) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' CTS�̏�Ԃ��ǂݎ��܂���ł���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10190, _
                    Description:="RTSCTS - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")��CTS�̏�Ԃ��ǂݎ��܂���ł����D"
                RTSCTS = False
                Exit Property
        
        End Select
        
    End If

    If (Stat And MS_CTS_ON) <> 0 Then
        'CTS is ON
        RTSCTS = True
      Else
        'CTS is OFF
        RTSCTS = False
    End If
End Property

'��������
Public Property Let RTSCTS(Status As Boolean)
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10199, _
                    Description:="RTSCTS - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
        
    End If
    
    If Status = True Then
        Stat = SETRTS
      Else
        Stat = CLRRTS
    End If

    If EscapeCommFunction(ecH(Cn).Handle, Stat) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' RTS�̐ݒ�Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10191, _
                    Description:="RTSCTS - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")��RTS�̐ݒ�Ɏ��s���܂����D"
                Exit Property
        
        End Select
    End If
End Property

'-----------------------------------
'   DTRDSR�v���p�e�B
'-----------------------------------
' DTR�̋��������DSR�̏�ԓǂݎ��

'�ǂݏo��(DSR�̏��)
Public Property Get DTRDSR() As Boolean
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10208, _
                    Description:="DTRDSR - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
        
    End If
    
    If GetCommModemStatus(ecH(Cn).Handle, Stat) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' DSR�̏�Ԃ��ǂݎ��܂���ł���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10200, _
                    Description:="DTRDSR - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")��DSR�̏�Ԃ��ǂݎ��܂���ł����D"
                DTRDSR = False
                Exit Property
        
        End Select
    End If

    If (Stat And MS_DSR_ON) <> 0 Then
        'DSR is ON
        DTRDSR = True
      Else
        'DSR is OFF
        DTRDSR = False
    End If
End Property

'��������
Public Property Let DTRDSR(Status As Boolean)
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10209, _
                    Description:="DTRDSR - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property

        End Select
        
    End If
    
    If Status = True Then
        Stat = SETDTR
      Else
        Stat = CLRDTR
    End If

    If EscapeCommFunction(ecH(Cn).Handle, Stat) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' DTR�̐ݒ�Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10201, _
                    Description:="DTRDSR - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")��DTR�̐ݒ�Ɏ��s���܂����D"
                Exit Property
        
        End Select
        
    End If
End Property


'-----------------------------------
'   Delimiter�v���p�e�B
'-----------------------------------
'AsciiLine�v���p�e�B�Ŏg�p����f���~�^�̐ݒ�C�ǂݏo�����s���v���p�e�B�ł��D
'��������
Public Property Let Delimiter(DelimiterType As String)
    Dim DelimStr As String
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                Stop        ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10248, _
                    Description:="Delimiter - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                End             ' �v���O�������I�����܂��D
        
        End Select

    End If

    ' �����̐��K��
    DelimStr = Trim(UCase(StrConv(DelimiterType, vbNarrow)))
    Select Case DelimStr
        Case Is = "CR"
            ecH(Cn).Delimiter = "CR"
        Case Is = "LF"
            ecH(Cn).Delimiter = "LF"
        Case Is = "CRLF"
            ecH(Cn).Delimiter = "CRLF"
        Case Is = "LFCR"
            ecH(Cn).Delimiter = "LFCR"
        Case Else
            ecH(Cn).Delimiter = "CR"    ' �K��l�ȊO�͂b�q�Ƃ݂Ȃ��܂��D
    End Select

End Property

'�ǂݏo��
Public Property Get Delimiter() As String
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10249, _
                    Description:="Delimiter - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If

    ' �ݒ�l�̓ǂݎ��
    Select Case ecH(Cn).Delimiter
        Case Is = ""
            Delimiter = "CR"
        Case Is = "CR"
            Delimiter = "CR"
        Case Is = "LF"
            Delimiter = "LF"
        Case Is = "CRLF"
            Delimiter = "CRLF"
        Case Is = "LFCR"
            Delimiter = "LFCR"
        Case Else
            Delimiter = "CR"                ' �K��l�ȊO�͂b�q�Ƃ݂Ȃ��܂��D
            ecH(Cn).Delimiter = "CR"        ' CR �ɕ␳���܂�
    End Select

End Property

'-----------------------------------
'   Break�v���p�e�B
'-----------------------------------
'�u���[�N�M���𑗐M�C�܂��̓u���[�N�M���̑��M���~���鏑�����ݐ�p�̃v���p�e�B�ł��D
Public Property Let Break(BreakOn As Boolean)
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10259, _
                    Description:="Break - Write" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property

        End Select
        
    End If
    
    If BreakOn = True Then
        ' Break�̑��M
        If SetCommBreak(ecH(Cn).Handle) = False Then
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' Break�̑��M�Ɏ��s���܂���
                End             ' �v���O�������I�����܂��D

                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10251, _
                        Description:="Break - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̃u���[�N���M�Ɏ��s���܂����D"
                    Exit Property
            
            End Select
        End If
      Else
        ' Break���M�̒�~
        If ClearCommBreak(ecH(Cn).Handle) = False Then
            '�G���[����
            Select Case Xerror
                Case Is = 0 '-----�W���G���[
                    ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                    Stop            ' Break���M�̒�~�Ɏ��s���܂���
                    End             ' �v���O�������I�����܂��D

                Case Is = 1 '-----�g���b�v�\�G���[
                    err.Raise _
                        Number:=10252, _
                        Description:="Break - Write" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̃u���[�N���M�̒�~�Ɏ��s���܂����D"
                    Exit Property
            
            End Select
        End If
          
    End If

End Property

'-----------------------------------
'   RI�v���p�e�B(�ǎ��p)
'-----------------------------------
' RI�̏�ԓǂݎ��
Public Property Get RI() As Boolean
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10268, _
                    Description:="RI - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
        
    End If
    
    If GetCommModemStatus(ecH(Cn).Handle, Stat) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' RI�̏�Ԃ��ǂݎ��܂���ł���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10261, _
                    Description:="RI - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")��RI�̏�Ԃ��ǂݎ��܂���ł����D"
                DTRDSR = False
                Exit Property
        
        End Select
    End If

    If (Stat And MS_RING_ON) <> 0 Then
        'RI is ON
        RI = True
      Else
        'RI is OFF
        RI = False
    End If
End Property

'-----------------------------------
'   DCD�v���p�e�B(�ǎ��p)
'-----------------------------------
' DCD�̏�ԓǂݎ��
Public Property Get DCD() As Boolean
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10278, _
                    Description:="DCD - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select
        
    End If
    
    If GetCommModemStatus(ecH(Cn).Handle, Stat) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' DCD�̏�Ԃ��ǂݎ��܂���ł���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10271, _
                    Description:="DCD - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")��DCD�̏�Ԃ��ǂݎ��܂���ł����D"
                DTRDSR = False
                Exit Property
        
        End Select
    End If

    If (Stat And MS_RLSD_ON) <> 0 Then
        'DCD is ON
        DCD = True
      Else
        'DCD is OFF
        DCD = False
    End If
End Property


'-----------------------------------
'   Spec�v���p�e�B
'-----------------------------------
'�|�[�g�n���h�� ecH(Cn).Handle �̏��𕶎���ŕԂ��܂��D
Public Property Get Spec() As String
    Dim Er As Long
    Dim Mes As String
    Dim CrLf As String
    'CRLF = Chr$(&HD) & Chr$(&HA)
    CrLf = Chr$(&HA)

    If Cn = 0 Then
        '�|�[�g�ԍ����w��
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=10238, _
                    Description:="InBuffer - Read" & Chr$(&HA) & "�ΏۂƂȂ�|�[�g�ԍ����w�肳��Ă��܂���"
                Exit Property
        
        End Select

    End If

    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        '�G���[����
        Select Case Xerror
            Case Is = 0 '-----�W���G���[
                ec.CloseAll     ' ���ׂẴ|�[�g����܂��D
                Stop            ' �|�[�g�̏����擾�ł��܂���ł���
                End             ' �v���O�������I�����܂��D

            Case Is = 1 '-----�g���b�v�\�G���[
                err.Raise _
                    Number:=12300, _
                    Description:="InCount - Read" & Chr$(&HA) & "�|�[�g(" & Cn & ")�̏����擾�ł��܂���ł���"
                Exit Property
        
        End Select

    End If

    '�����o�̊m�F
    With fCOMMPROP
        Mes = "�\���̂̃o�C�g�T�C�Y" & CrLf
        Mes = Mes & "�@" & .wPacketLength & CrLf
        Mes = Mes & "�o�[�W����" & CrLf
        Mes = Mes & "�@" & .wPacketVersion & CrLf
        Mes = Mes & "���M�o�b�t�@�̍ő�o�C�g��" & CrLf
        If .dwMaxTxQueue = 0 Then
            Mes = Mes & "�@��������" & CrLf
          Else
            Mes = Mes & "�@" & .dwMaxTxQueue & CrLf
        End If
        Mes = Mes & "��M�o�b�t�@�̍ő�o�C�g��" & CrLf
        If .dwMaxRxQueue = 0 Then
            Mes = Mes & "�@��������" & CrLf
          Else
            Mes = Mes & "�@" & .dwMaxRxQueue & CrLf
        End If
        Mes = Mes & "�ő�f�[�^�]�����x" & CrLf & "�@"
        Select Case .dwMaxBaud
            Case Is = BAUD_075
                Mes = Mes & "75 bps"
            Case Is = BAUD_110
                Mes = Mes & "110 bps"
            Case Is = BAUD_134_5
                Mes = Mes & "134.5 bps"
            Case Is = BAUD_150
                Mes = Mes & "150 bps"
            Case Is = BAUD_300
                Mes = Mes & "300 bps"
            Case Is = BAUD_600
                Mes = Mes & "600 bps"
            Case Is = BAUD_1200
                Mes = Mes & "1200 bps"
            Case Is = BAUD_1800
                Mes = Mes & "1800 bps"
            Case Is = BAUD_2400
                Mes = Mes & "2400 bps"
            Case Is = BAUD_4800
                Mes = Mes & "4800 bps"
            Case Is = BAUD_7200
                Mes = Mes & "7200 bps"
            Case Is = BAUD_9600
                Mes = Mes & "9600 bps"
            Case Is = BAUD_14400
                Mes = Mes & "14400 bps"
            Case Is = BAUD_19200
                Mes = Mes & "19200 bps"
            Case Is = BAUD_38400
                Mes = Mes & "38400 bps"
            Case Is = BAUD_56K
                Mes = Mes & "56 K bps"
            Case Is = BAUD_57600
                Mes = Mes & "57600 bps"
            Case Is = BAUD_115200
                Mes = Mes & "115200 bps"
            Case Is = BAUD_128K
                Mes = Mes & "128 K bps"
            Case Is = BAUD_USER
                Mes = Mes & "�v���O���}�u��"
        End Select
        Mes = Mes & CrLf

        Mes = Mes & "�T�|�[�g����Ă���@�\" & CrLf
        
        If .dwProvCapabilities & PCF_DTRDSR <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�FDTR/DSR" & CrLf
        
        If .dwProvCapabilities & PCF_RTSCTS <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�FRTS/CTS" & CrLf
        
        If .dwProvCapabilities & PCF_RLSD <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�FCD(RLSD)" & CrLf
        
        If .dwProvCapabilities & PCF_PARITY_CHECK <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�p���e�B�`�F�b�N" & CrLf
        
        If .dwProvCapabilities & PCF_XONXOFF <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�FXON/XOFF�ɂ��t���[����" & CrLf
        
        If .dwProvCapabilities & PCF_SETXCHAR <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�FXON/XOFF�̕����w��" & CrLf
        
        If .dwProvCapabilities & PCF_TOTALTIMEOUTS <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�g�[�^���^�C���A�E�g�̐ݒ�" & CrLf
        
        If .dwProvCapabilities & PCF_INTTIMEOUTS <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�C���^�[�o���^�C���A�E�g�̐ݒ�" & CrLf
        
        If .dwProvCapabilities & PCF_SPECIALCHARS <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F���ꕶ���̎g�p" & CrLf
       
        If .dwProvCapabilities & PCF_16BITMODE <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�����16�r�b�g���[�h" & CrLf

        Mes = Mes & "�e��@�\�̐ݒ�̉�" & CrLf
        
        If .dwSettableParams & SP_PARITY <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�p���e�B�̃��[�h" & CrLf
        
        If .dwSettableParams & SP_BAUD <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�{�[���[�g" & CrLf

        If .dwSettableParams & SP_DATABITS <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�f�[�^�r�b�g��" & CrLf

        If .dwSettableParams & SP_STOPBITS <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�X�g�b�v�r�b�g��" & CrLf

        If .dwSettableParams & SP_HANDSHAKING <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�t���[����i�n���h�V�F�[�N�j" & CrLf

        If .dwSettableParams & SP_PARITY_CHECK <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�p���e�B�`�F�b�N��ON/OFF" & CrLf

        If .dwSettableParams & SP_RLSD <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�FCD(RLSD)" & CrLf

        If .dwSettableParams & SP_PARITY <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�p���e�B���[�h" & CrLf

        Mes = Mes & "�ݒ�\�ȃ{�[���[�g" & CrLf
        
        If .dwSettableBaud And BAUD_075 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F75 bps" & CrLf
        
        If .dwSettableBaud And BAUD_110 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F110 bps" & CrLf
        
        If .dwSettableBaud And BAUD_134_5 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F134.5 bps" & CrLf
        
        If .dwSettableBaud And BAUD_150 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F150 bps" & CrLf
            
        If .dwSettableBaud And BAUD_300 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F300 bps" & CrLf
            
        If .dwSettableBaud And BAUD_600 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F600 bps" & CrLf
            
        If .dwSettableBaud And BAUD_1200 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F1200 bps" & CrLf
            
        If .dwSettableBaud And BAUD_1800 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F1800 bps" & CrLf
            
        If .dwSettableBaud And BAUD_2400 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F2400 bps" & CrLf
            
        If .dwSettableBaud And BAUD_4800 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F4800 bps" & CrLf
            
        If .dwSettableBaud And BAUD_7200 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F7200 bps" & CrLf
            
        If .dwSettableBaud And BAUD_9600 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F9600 bps" & CrLf
            
        If .dwSettableBaud And BAUD_14400 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F14400 bps" & CrLf
            
        If .dwSettableBaud And BAUD_19200 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F19200 bps" & CrLf
            
        If .dwSettableBaud And BAUD_38400 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F38400 bps" & CrLf
            
        If .dwSettableBaud And BAUD_56K <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F56 K bps" & CrLf
            
        If .dwSettableBaud And BAUD_57600 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F57600 bps" & CrLf
            
        If .dwSettableBaud And BAUD_115200 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F115200 bps" & CrLf
            
        If .dwSettableBaud And BAUD_128K <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F128 K bps" & CrLf
            
        If .dwSettableBaud And BAUD_USER <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�v���O����" & CrLf

        Mes = Mes & "�ݒ�\�ȃf�[�^�r�b�g��" & CrLf
        
        If .wSettableData And DATABITS_5 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F5 �r�b�g" & CrLf

        If .wSettableData And DATABITS_6 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F6 �r�b�g" & CrLf
        
        If .wSettableData And DATABITS_7 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F7 �r�b�g" & CrLf
        
        If .wSettableData And DATABITS_8 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F8 �r�b�g" & CrLf
        
        If .wSettableData And DATABITS_16 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F16 �r�b�g" & CrLf
        
        If .wSettableData And DATABITS_16X <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F����ȃ��C�h�p�X" & CrLf

        Mes = Mes & "�ݒ�\�ȃX�g�b�v�r�b�g��" & CrLf
        
        If .wSettableStopParity And STOPBITS_10 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F1 �r�b�g" & CrLf
        
        If .wSettableStopParity And STOPBITS_15 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F1.5 �r�b�g" & CrLf
        
        If .wSettableStopParity And STOPBITS_20 <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F2 �r�b�g" & CrLf

        Mes = Mes & "�ݒ�\�ȃp���e�B�`�F�b�N" & CrLf
        
        If .wSettableStopParity And PARITY_NONE <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�p���e�B�Ȃ�" & CrLf

        If .wSettableStopParity And PARITY_ODD <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F��p���e�B" & CrLf

        If .wSettableStopParity And PARITY_EVEN <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�����p���e�B" & CrLf

        If .wSettableStopParity And PARITY_MARK <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�}�[�N�p���e�B" & CrLf

        If .wSettableStopParity And PARITY_SPACE <> 0 Then
            Mes = Mes & "�@��"
          Else
            Mes = Mes & "�@�~"
        End If
        Mes = Mes & "�F�X�y�[�X�p���e�B" & CrLf

    End With
    Spec = Mes
    
End Property
'-----------------------------------
'   CloseAll���\�b�h
'-----------------------------------
Private Sub CloseAll()
'���ׂẴ|�[�g�����C���������p�̃��\�b�h�ł��D
'ecPorts.tmp�ɕۑ�����Ă���n���h�����N���[�Y���C�t�@�C�����폜���܂��D
'�G���[��Ԃ��܂���
    Dim i As Long
    Dim rv As Long
    Dim Handle As Long
    Dim FileNumber As Integer
    Dim Fpath As String * 260
    FileNumber = FreeFile()
    
    GetTempPath 260, Fpath
    Open Left(Fpath, InStr(Fpath, vbNullChar) - 1) & "ecPort.tmp" For Random Access Read Write As #FileNumber Len = Len(Handle)

    For i = 1 To ecMaxPort
        ecH(i).Delimiter = "CR"         ' �f���~�^�����Z�b�g���܂�
        ecH(i).LineInTimeOut = 0        ' �^�C���A�E�g�̃��Z�b�g
        Get #FileNumber, i, Handle      ' �L�^����Ă���n���h�����擾���܂�
        If Handle > 0 Then
            CloseHandle Handle          ' �N���[�Y
            rv = 0&
            Put #FileNumber, i, rv
        End If
        ecH(i).Handle = 0
    Next i
    Cn = 0      ' �����Ώۂ̃|�[�g�ԍ������Z�b�g
    Close #FileNumber
End Sub

'-----------------------------------
'�B���R�}���h

'-----------------------------------
'   OutBufferSize�v���p�e�B
'-----------------------------------
'�ǂݏo����p
'���ݏ����̑ΏۂƂȂ��Ă���|�[�g�̑��M�o�b�t�@�̃T�C�Y��Ԃ��܂�
'�G���[����-1��Ԃ��܂�
Private Property Get OutBufferSize() As Long
    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        '��Ԃ̎擾�Ɏ��s
        OutBufferSize = -1
        Exit Property
    End If
    OutBufferSize = ecDef.fCOMMPROP.dwCurrentTxQueue
End Property

'-----------------------------------
'   InBufferSize�v���p�e�B
'-----------------------------------
'�ǂݏo����p
'���ݏ����̑ΏۂƂȂ��Ă���|�[�g�̎�M�o�b�t�@�̃T�C�Y��Ԃ��܂�
'�G���[����-1��Ԃ��܂�
Private Property Get InBufferSize() As Long
    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        '��Ԃ̎擾�Ɏ��s
        InBufferSize = -1
        Exit Property
    End If
    InBufferSize = ecDef.fCOMMPROP.dwCurrentRxQueue
End Property

'-----------------------------------
'   DELIMs�v���p�e�B
'-----------------------------------
'DELIMs
'�f���~�^�p�ݒ蕶����擾�v���p�e�B(�ǂݏo����p)
'   Delimiter�v���p�e�B�Ŏg�p����f���~�^�̐ݒ�p��������擾���邽�߂̃v���p�e�B�D
'   ���̗�́C������Ŏw�肷����̂�HANDSHAKINGs�v���p�e�B���g�������̂Ƃ̔�r�ł��D
'��
' ���f���~�^��Cr�ɐݒ肵�܂��D
'   ec.Delimiter = ec.DELIMs.Cr
'   ec.Delimiter = "Cr"
' ���f���~�^��Cr+Lf�ɐݒ肵�܂��D
'   ec.Delimiter = ec.DELIMs.CrLf
'   ec.Delimiter = "CRLF"

Public Property Get DELIMs() As DelimType
    ' �f���~�^�w�蕶����萔
    DELIMs.Cr = "CR"            ' CR
    DELIMs.Lf = "LF"            ' LF
    DELIMs.CrLf = "CRLF"        ' CR + LF
    DELIMs.LfCr = "LFCR"        ' LF + CR
End Property

'-----------------------------------
'   HANDSHAKEs�v���p�e�B
'-----------------------------------
'�n���h�V�F�[�L���O�ݒ蕶����擾�v���p�e�B(�ǂݏo����p)
'   HandHsaking�v���p�e�B�Ŏg�p����ݒ�p��������擾���邽�߂̃v���p�e�B�D
'   ���̗�́C������Ŏw�肷����̂�HANDSHAKINGs�v���p�e�B���g�������̂Ƃ̔�r�ł��D
'��
' ���n���h�V�F�[�N���Ȃ��ɐݒ肵�܂��D
'   ec.HandShaking = ec.HANDSHAKEs.No
'   ec.HandShaking = "N"
' ���n���h�V�F�[�N��RTS/CTS�ɐݒ肵�܂��D
'   ec.HandShaking = ec.HANDSHAKEs.RTSCTS
'   ec.HandShaking = "R"
'
Public Property Get HANDSHAKEs() As HandShakingType
    ' �n���h�V�F�[�N�w�蕶����萔
    HANDSHAKEs.No = "N"         ' �Ȃ�
    HANDSHAKEs.XonXoff = "X"    ' Xon/Off
    HANDSHAKEs.RTSCTS = "R"     ' RTS/CTS
    HANDSHAKEs.DTRDSR = "D"     ' DTR/DSR
End Property


