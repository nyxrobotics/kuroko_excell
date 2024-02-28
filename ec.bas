Attribute VB_Name = "ec"
Option Explicit
'   EasyComm
'   Module ec.bas

'－－－－－－EasyComm 利用規定－－－－－－
'   下記条件に同意された方のみご利用ください．
'   （ＥａｓｙＣｏｍｍをｅｃと略記します）
'
'（１）ｅｃを単体で販売することは出来ません．
'（２）ｅｃを動作させた結果，いかなる損害が発生しても，木下隆は一切責任を負いません．
'（３）ｅｃは個人，会社を問わず，木下隆の許可無く自由に配布し，ダウンロードできます．
'（４）ｅｃを独自のアプリケーションに組み込んだ場合，木下隆の許可無く自由に配布，販売
'       することが出来ます．
'（５）ｅｃは，アプリケーションの一部として利用する場合に限って，木下隆の許可無く自由
'       に改造し，または一部を引用して利用，配布，販売することが出来ます．
'
'   Version 1.00    Aug 25,2000  Created by T.Kinoshita
'   Version 1.10    Nov 21,2000
'   Version 1.20    Jan 14,2001
'       DCD,RI,Breakプロパティ追加
'   Version 1.30    Mar 31,2001
'       送信バッファの初期設定値を変更
'       デフォルトの送受信バッファの設定方法を変更
'       Xerror=1で，STOPの後にすべてのポートを閉じるようにコードを追加
'   Version 1.31    May 17,2001
'       WAITmSのバグ修正（町田様のご指摘による）
'   Version 1.40    Jul 25,2001
'       Binaryプロパティの書き込みじのデータ型対応を拡張
'       モードゼロのエラー発生時は，すべてのポートを閉じるように仕様変更
'       隠しメソッドのCLoseAllをCloseAllに変更
'       ログファイルの廃止
'       デフォルトのハンドシェーキングを"N"に変更
'       デリミタ指定用文字列を取得するための読み取り専用プロパティDELIMSを追加
'       ハンドシェーキング指定用文字列を取得するための読み取り専用プロパティHANDSHAKEsを追加
'       デフォルトの設定を変更
'
'   Version 1.50    Oct 23,2001
'       Windows2000に対応（ReadFile,WriteFileの定義部修正
'
'   Version 1.51    Jan 17,2002
'       tomotさんのご意見を参考に，AsciiLineの読み出しタイムアウトを追加しました．
'   Version 1.51a   Jan 17,2002
'       COMnプロパティにゼロ以下の値を代入したとき，ポートセッティング変数をリセットするように変更
'       COMnCloseプロパティで閉じたポートに対しても，ポートセッティング変数をリセットする用に変更
'   Version 1.60    Jul 9,2002
'       Versionプロパティ追加(PrivateをPublicに変更)
'       DoseSecondsプロパティの追加
'       最大同時オープンポート数を20から50に拡大
'       ファイルハンドルを記憶するテンポラリファイルecPort.tmpの使用
'       AsciiLineプロパティの受信で，設定によっては文字列の末尾にchr(0)が入るバグを修正
'   Version 1.61    Jul 23,2002
'       shinさんのご指摘により，テンポラリファイルのフォルダをThisWorkbook.Pathにしていたため，
'       Excel以外のアプリで使用不能になっていたバグを修正．
'       Windowsのテンポラリファイルに作成するようにした．
'   Version 1.62    Jul 31,2002
'       くりさんのご指摘によりExcel97での不具合を修正
'   Version 1.62a   Aug 10,2002
'       利用規定を変更
'   Version 1.70    Mar 6,2003
'       Binaryプロパティ読み出し時のバイト数を指定するBinaryBytesプロパティを追加．
'       Asciiプロパティ読み出し時のバイト数を指定するAsciiBytesプロパティを追加．
'       いずれもデフォルトでゼロ(旧バージョン互換)
'       InBufferClearメソッドを追加(メソッドは初めて)
'   Version 1.71    Mar 7,2003
'       ご迷惑陳謝バージョン
'       BinaryBytes,AsciiBytesプロパティのバグ修正
'       指定バイト数のデータがバッファに無い時は，受信されるまで停止状態になるバグの修正
'
Public Const Version As String = "1.71"
'
'   Copyright(c) 2000 T.Kinoshita
'   Copyright(c) 2001 T.Kinoshita
'   Copyright(c) 2002 T.Kinoshita
'   Copyright(c) 2003 T.Kinoshita


'===================================
' 変数の定義
'===================================

'-----------------------------------
'   AsciiBytesプロパティ(Version1.70から追加)
'-----------------------------------
'Asciiプロパティ読み出し時に受信バッファから取り出すバイト数を指定します
'デフォルトはゼロで，ゼロ以下の時は受信バッファのデータをすべて取得します．
Public AsciiBytes As Long

'-----------------------------------
'   BinaryBytesプロパティ(Version1.70から追加)
'-----------------------------------
'Asciiプロパティ読み出し時に受信バッファから取り出すバイト数を指定します
'デフォルトはゼロで，ゼロ以下の時は受信バッファのデータをすべて取得します．
Public BinaryBytes As Long

'-----------------------------------
'   Xerrorプロパティ
'-----------------------------------
'エラーモードを指定するプロパティ
'ゼロ(規定）の時はEasyComm標準のエラーモードで，エラーが発生すると停止します．
'Version1.4からは，このモードでエラーが発生するとすべてのポートを閉じるように仕様変更．
'１の時はトラップ可能なエラーを発生．
'ただし，１に設定してエラーが発生したとき，トラップルーティンがないとプログラムを終し，て変数を
'リセットするので，すでに開かれたポートを閉じることができなくなることがあります．
'２の時は，エラーを無視します．
'Public宣言することによって，読み書き可能なプロパティにしています．
Public Xerror As Long

'デリミタ指定文字列
Type DelimType
    Cr      As String   ' CR
    Lf      As String   ' LF
    CrLf    As String   ' CR + LF
    LfCr    As String   ' LF + CR
End Type

'ハンドシェーキング指定文字列
Type HandShakingType
    No      As String   ' なし
    XonXoff As String   ' Xon/Off
    RTSCTS  As String   ' RTS/CTS
    DTRDSR  As String   ' DTR/DSR
End Type

'===================================
'   プロパティ，メソッドの定義
'===================================

'-----------------------------------
'   InBufferClearメソッド
'-----------------------------------
'COMnプロパティで指定されているポートの受信バッファのデータをクリアします.
'Asciiプロパティの読み出しで代用していましたが，AsciiBytesプロパティの追加に伴い，InBufferClearメソッドを追加しました．
Public Sub InBufferClear()
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10280, _
                    Description:="InBufferClear " & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Sub
        End Select
    End If
    If ecDef.PurgeComm(ecH(Cn).Handle, ecDef.PURGE_RXCLEAR) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 受信バッファのクリアに失敗しました．
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10281, _
                    Description:="InBufferClear " & Chr$(&HA) & "受信バッファのクリアに失敗しました"
                Exit Sub
        End Select
    End If

End Sub

'-----------------------------------
'   COMnプロパティ(読み出し，書き込み)
'-----------------------------------
'操作の対象となうポート番号を指定，または取得します．
'COMnプロパティにゼロ以下の整数を代入すると，すべてのポートを閉じます．
'ec.COMn = 0 と ec.COMnClose = 0 は，同じ動作をします．
'プログラムの終了時に，いずれかのコードですべてのポートを閉じてください．

'読み出しプロシージャ
Public Property Get COMn() As Long
    COMn = Cn
End Property

'書き込みプロシージャ
Public Property Let COMn(PortNumber As Long)
    Dim rv As Long
    Dim PortName As String
    Dim SettingFlag As Boolean          ' 標準設定用フラグ
    Dim i As Integer
    Dim Fnumb As Integer                ' 記録ファイルのファイル番号
    Dim RxBuffer As Long                ' 受信バッファサイズの設定用変数
    Dim TxBuffer As Long                ' 送信バッファサイズの設定用変数
    Dim FileNumber As Integer           ' オープン可能なファイル番号
    Dim Handle As Long
    Dim Fpath As String * 260
    FileNumber = FreeFile()

    If PortNumber <= 0 Then
        'すべてのポートを閉じます
        ec.CloseAll

      ElseIf PortNumber > ecMaxPort Then

        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 指定されたポート番号は許容範囲にありません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10010, _
                    Description:="COMn - Write" & Chr$(&HA) & "指定されたポート番号(" & PortNumber & ")は許容範囲にありません"
                End

        End Select
        
      Else
        
        '既に開かれているかどうかチェック
        If ecH(PortNumber).Handle > 0 Then
            '既に開かれている
            Cn = PortNumber     '処理対象のポートを引数に設定
            Exit Property       '完了
        End If

        'ポートオープン処理
        
        PortName = "\\.\COM" & Trim(Str(PortNumber))
        rv = CreateFile(PortName, GENERIC_READ Or GENERIC_WRITE, _
                        0&, 0&, OPEN_EXISTING, 0&, 0&)

        GetTempPath 260, Fpath
        If rv <> INVALID_HANDLE_VALUE Then
            ' 記録
            Open Left(Fpath, InStr(Fpath, vbNullChar) - 1) & "ecPort.tmp" For Random Access Read Write As #FileNumber Len = Len(Handle)
            Put #FileNumber, PortNumber, rv
            Close #FileNumber
            
            ecH(PortNumber).Handle = rv     ' ハンドルリストの更新
        End If
        Cn = PortNumber                     ' アクティブポート番号の設定
        ' 失敗したときは記録されているハンドルを使ってポートを閉じてから再挑戦
        If rv = INVALID_HANDLE_VALUE Then
            Open Left(Fpath, InStr(Fpath, vbNullChar) - 1) & "ecPort.tmp" For Random Access Read Write As #FileNumber Len = Len(Handle)
            Get #FileNumber, PortNumber, Handle
            If Handle > 0 Then
                ecDef.CloseHandle Handle
                rv = CreateFile(PortName, GENERIC_READ Or GENERIC_WRITE, _
                                0&, 0&, OPEN_EXISTING, 0&, 0&)
                If rv <> INVALID_HANDLE_VALUE Then
                    ' 記録
                    Put #FileNumber, PortNumber, rv
                    ecH(PortNumber).Handle = rv     ' ハンドルリストの更新
                  Else
                    rv = 0&
                    Put #FileNumber, PortNumber, rv
                    ecH(PortNumber).Handle = 0      ' ハンドルリストの更新
                End If
            End If
            Close #FileNumber
        End If

        If rv = INVALID_HANDLE_VALUE Then
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' 指定されたポート番号を開けませんでした
                    End             ' プログラムを終了します．

                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10011, _
                        Description:="COMn - Write" & Chr$(&HA) & "指定されたポート番号(" & PortNumber & ")を開けませんでした"
                    End
            
            End Select
        End If

        '--------------------------
        '開いたポートの標準設定
        SettingFlag = True
        
        'ポートの機能を取得
        If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' ポートの情報を取得できませんでした
                    End             ' プログラムを終了します．
    
                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10013, _
                        Description:="InCount - Read" & Chr$(&HA) & "ポート(" & Cn & ")の情報を取得できませんでした"
                    Exit Property
            
            End Select
        End If

        ' 受信バッファサイズの設定
        RxBuffer = ecDef.fCOMMPROP.dwMaxRxQueue
        If RxBuffer = 0 Then    ' 無制限の時
            RxBuffer = ecDef.ecInBufferSize     ' 規定の最大値に制限する
        End If

        ' 送信バッファサイズの設定
        TxBuffer = ecDef.fCOMMPROP.dwMaxTxQueue
        If TxBuffer = 0 Then    ' 無制限の時
            TxBuffer = ecDef.ecOutBufferSize    ' 規定の最大値に制限する
        End If

        ' 実行
        If SetupComm(ecH(PortNumber).Handle, RxBuffer, TxBuffer) = False Then
            SettingFlag = False
        End If

        ' DCBのセット
        ' 標準設定値は一般的なものなので，サポートの有無を確認しない．
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
        
        'バッファのクリア
        If PurgeComm(ecH(PortNumber).Handle, PURGE_TXCLEAR) = False Then
            SettingFlag = False
        End If
        If PurgeComm(ecH(PortNumber).Handle, PURGE_RXCLEAR) = False Then
            SettingFlag = False
        End If
    
        'タイムアウトの設定
        If GetCommTimeouts(ecH(PortNumber).Handle, fCOMMTIMEOUTS) <> False Then
            'タイムアウト値の標準設定
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
            
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' ポートの標準設定が出来ませんでした
                    End             ' プログラムを終了します．
                
                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10012, _
                        Description:="COMn - Write" & Chr$(&HA) & "指定されたポート番号(" & PortNumber & ")の標準設定が出来ませんでした"
                        Exit Property
            
            End Select
            
        End If
    End If
    
End Property

'-----------------------------------
'   COMnCloseプロパティ
'-----------------------------------
'指定した番号のポートを閉じます．
'COMnCloseプロパティにゼロ以下の整数を代入すると，すべてのポートを閉じます．
'ec.COMn = 0 と ec.COMnClose = 0 は，同じ動作をします．
'プログラムの終了時に，いずれかのコードですべてのポートを閉じてください．
Public Property Let COMnClose(PortNumber As Long)
    Dim rv As Long
    Dim SheetObj As Object
    Dim i As Integer
    Dim Fnumb As Integer
    Dim FileNumber As Integer           ' オープン可能なファイル番号
    Dim Handle As Long
    Dim Fpath As String * 260
    FileNumber = FreeFile()

    If PortNumber <= 0 Then
        'すべてのポートを閉じます
        'エラーを返しません
        ec.CloseAll
      Else
        '指定されたポートを閉じます
        'エラーを返します．
        rv = CloseHandle(ecH(PortNumber).Handle)

       '最終処理
        If rv = False Then
            '失敗
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' 指定されたポート番号のクローズに失敗しました
                    End             ' プログラムを終了します．
                
                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10020, _
                        Description:="COMnClose - Write" & Chr$(&HA) & "指定されたポート番号(" & PortNumber & ")のクローズに失敗しました"
            
            End Select
          
          Else
            '成功
            ecH(PortNumber).Handle = 0          'ハンドルリストの更新
            ecH(PortNumber).Delimiter = ""      '標準デリミタ
            ecH(PortNumber).LineInTimeOut = 0   'AsciiLineTimeOutのリセット
            GetTempPath 260, Fpath
            Open Left(Fpath, InStr(Fpath, vbNullChar) - 1) & "ecPort.tmp" For Random Access Read Write As #FileNumber Len = Len(Handle)
            rv = 0&
            Put #FileNumber, PortNumber, rv
            Close #FileNumber
        End If
    End If
End Property

'-----------------------------------
'   Settingプロパティ(書き込み読み出し)
'-----------------------------------
'操作の対象となっているポートの通信条件の設定，または読み出しを行ないます．

'読み出し
Public Property Get Setting() As String

    Dim ModeStr As String

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10038, _
                    Description:="Setting - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
    
    End If

    If GetCommState(ecH(Cn).Handle, fDCB) = False Then     'DCBの読み出し
        
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 現在設定値(DCB)の読み込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10030, _
                    Description:="Setting - Read" & Chr$(&HA) & "ポート(" & Cn & ")の現在設定値(DCB)の読み込みに失敗しました"
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

'書き込み
Public Property Let Setting(Mode As String)
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10039, _
                    Description:="Setting - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
        
    End If

    ' DCBの読み出し
    If GetCommState(ecH(Cn).Handle, fDCB) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 現在設定値(DCB)の読み込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10031, _
                    Description:="Setting - Write" & Chr$(&HA) & "ポート(" & Cn & ")の現在設定値(DCB)の読み込みに失敗しました"
                Exit Property
        
        End Select
        
    End If

    '文字列の変換
    If BuildCommDCB(Mode, fDCB) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 設定文字列の解析に失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10032, _
                    Description:="Setting - Write" & Chr$(&HA) & "ポート(" & Cn & ")の設定文字列(" & Mode & ")の解析に失敗しました"
                Exit Property
        
        End Select

    End If
    
    'DCBの書き込み
    If SetCommState(ecH(Cn).Handle, fDCB) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 現在設定値(DCB)の書き込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10033, _
                    Description:="Setting - Write" & Chr$(&HA) & "ポート(" & Cn & ")の現在設定値(DCB)の書き込みに失敗しました"
                Exit Property
        
        End Select
        
    End If
    
    'バッファのクリア
    'ここまでパスしていながらバッファのクリアでエラーが発生する可能性はないと思われるので，
    'エラーチェックは行いません．
    PurgeComm ecH(Cn).Handle, PURGE_TXCLEAR + PURGE_RXCLEAR

End Property


'-----------------------------------
'   HandShakingプロパティ(読み出し書き込み)
'-----------------------------------
'COMnプロパティで指定されているポートのフロー制御方式を設定，または取得します

'読み出し
Public Property Get HandShaking() As String
    Dim Flow As String
    
    If Cn = 0 Then
        'ポート番号未指定
        
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10188, _
                    Description:="HandShaking - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If
    
    ' DCBの読み出し
    If GetCommState(ecH(Cn).Handle, fDCB) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 現在設定値(DCB)の読み込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10180, _
                    Description:="HandShaking - Read" & Chr$(&HA) & "ポート(" & Cn & ")の状態を取得できませんでした．"
                Exit Property
        
        End Select

    End If
    
    Flow = ""

    'RTSハンドシェークのチェック
    If (fDCB.fBitFields And &H2000) <> 0 Then
        Flow = ec.HANDSHAKEs.RTSCTS
    End If

    'XON/OFFハンドシェークのチェック
    If (fDCB.fBitFields And &H300) <> 0 Then
        Flow = Flow & ec.HANDSHAKEs.XonXoff
    End If
    
    'DTR/DSRハンドシェークのチェック
    If (fDCB.fBitFields And &H20) <> 0 Then
        Flow = Flow & ec.HANDSHAKEs.DTRDSR
    End If
    
    If Flow = "" Then
        Flow = ec.HANDSHAKEs.No
    End If

    HandShaking = Flow

End Property

'書き込み
Public Property Let HandShaking(Flow As String)
    Dim f As String

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10189, _
                    Description:="HandShaking - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
        
    End If
    
    f = Left(UCase(Flow), 1)        '先頭１文字のみ有効
    
    ' DCBの読み出し
    If GetCommState(ecH(Cn).Handle, fDCB) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 現在設定値(DCB)の読み込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10181, _
                    Description:="HandShaking - Write" & Chr$(&HA) & "ポート(" & Cn & ")の状態を取得できませんでした．"
                Exit Property
        
        End Select
        
    End If

    With fDCB
        .fBitFields = .fBitFields And &HCC83            '該当ビットのリセット
        Select Case f
            Case Is = "N"
                'フローなし
                .fBitFields = .fBitFields Or &H1010     '該当ビットのセット
            Case Is = "X"
                'Xon/Offによるハンドシェイク
                .fBitFields = .fBitFields Or &H1310     '該当ビットのセット
            Case Is = "R"
                'RTSによるハンドシェーク
                .fBitFields = .fBitFields Or &H2014     '該当ビットのセット
            Case Is = "D"
                'DTRによるハンドシェーク
                .fBitFields = .fBitFields Or &H1068     '該当ビットのセット
        End Select
    End With

    ' DCBの書き込み
    If SetCommState(ecH(Cn).Handle, fDCB) = False Then

        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' ハンドシェークを書き込めませんでした
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10182, _
                    Description:="HandShaking - Write" & Chr$(&HA) & "ポート(" & Cn & ")のハンドシェークを書き込めませんでした．"
                Exit Property
        
        End Select
    End If
End Property

'-----------------------------------
'   InBufferプロパティ(読み出し書き込み)
'-----------------------------------
'COMnプロパティで指定されているポートの受信バッファの状態を設定，または取得します．

'読み出し
'受信バッファにたまった受信データのバイト数を取得します．
Public Property Get InBuffer() As Long
    Dim Er As Long
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10068, _
                    Description:="InBuffer - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If
    
    If ClearCommError(ecH(Cn).Handle, Er, fCOMSTAT) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 受信バッファのデータバイト数の読み取りに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10060, _
                    Description:="InBuffer - Read" & Chr$(&HA) & "ポート(" & Cn & ")の受信バッファのデータバイト数の読み取りに失敗しました"
                Exit Property
        
        End Select
        
    End If
    InBuffer = fCOMSTAT.cbInQue
End Property

'書き込み処理
'受信バッファサイズを指定します．
'最小設定値はecMinimumInBufferで指定されます．
'ecMinimumInBufferは，ecDef内で定義されています．
Public Property Let InBuffer(BufferSize As Long)
    Dim InBuff As Long      '現在の受信バッファサイズ
    Dim OutBuff As Long     '現在の受信バッファサイズ

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10069, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If

    '指定サイズ下限のチェック
    If BufferSize < ecMinimumInBuffer Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' バッファの値が少なすぎます．
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10065, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "バッファは " & ecMinimumInBuffer & "以上に設定してください"
                Exit Property
        
        End Select
        
    End If
    
    'ポートの設定値の取得
    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 受信バッファサイズの読み取りに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10061, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "ポート(" & Cn & ")の受信バッファサイズの読み取りに失敗しました"
                Exit Property
        
        End Select
        
    End If

    '上限値のチェック
    If fCOMMPROP.dwMaxRxQueue <> 0 Then     'ゼロの時は上限なし
        '上限値あり
        If BufferSize > fCOMMPROP.dwMaxRxQueue Then
            '上限を超えた設定
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' ポートの最大バッファサイズを越えた設定をしようとしました．
                    End             ' プログラムを終了します．
    
                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10066, _
                        Description:="InBuffer - Write" & Chr$(&HA) & "ポート(" & Cn & ")の受信バッファの上限値 (" & fCOMMPROP.dwMaxRxQueue & "を越える設定は出来ません．"
                    Exit Property
            
            End Select
        End If
    End If
    
    InBuff = fCOMMPROP.dwCurrentRxQueue     '受信バッファの設定を読み出す
    OutBuff = fCOMMPROP.dwCurrentTxQueue    '送信バッファの設定を読み出す
    
    '新しいサイズの設定
    If SetupComm(ecH(Cn).Handle, BufferSize, OutBuff) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 受信バッファサイズの書き込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10062, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "ポート(" & Cn & ")の受信バッファサイズの書き込みに失敗しました"
                Exit Property
        
        End Select
        
    End If
        
    ' DCBの読み出し
    If GetCommState(ecH(Cn).Handle, fDCB) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 現在設定値(DCB)の読み込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10063, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "ポート(" & Cn & ")の現在設定値(DCB)の読み込みに失敗しました"
                Exit Property
        
        End Select
        
    End If
    
    If SetCommState(ecH(Cn).Handle, fDCB) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 現在設定値(DCB)の書き込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10064, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "ポート(" & Cn & ")の設定値(DCB)の書き込みに失敗しました"
                Exit Property
        
        End Select
        
    End If

End Property

'-----------------------------------
'   OutBufferプロパティ(書き込み専用)
'-----------------------------------
'COMnプロパティで指定されているポートの送信バッファの状態を設定，または取得します．
'最小設定値はecMinimumOutBufferで指定されます．
'ecMinimumOutBufferは，ecDef内で定義されています．
Public Property Let OutBuffer(BufferSize As Long)
    
    Dim InBuff As Long      '現在の受信バッファサイズ
    Dim OutBuff As Long     '現在の送信バッファサイズ

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10079, _
                    Description:="OutBuffer - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If
    
    '指定サイズのチェック
    If BufferSize < ecMinimumOutBuffer Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 送信バッファの設定値が小さすぎます
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10073, _
                    Description:="InBuffer - Write" & Chr$(&HA) & "バッファは " & ecMinimumOutBuffer & "以上に設定してください"
                Exit Property
        
        End Select
        
    End If
    
    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 送信バッファサイズの読み取りに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10071, _
                    Description:="OutBuffer - Write" & Chr$(&HA) & "ポート(" & Cn & ")の送信バッファサイズの読み取りに失敗しました"
                Exit Property
        
        End Select
    End If
    
    '上限値のチェック
    If fCOMMPROP.dwMaxTxQueue <> 0 Then     'ゼロの時は上限なし
        '上限値あり
        If BufferSize > fCOMMPROP.dwMaxTxQueue Then
            '上限を超えた設定
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' ポートの最大バッファサイズを越えた設定をしようとしました．
                    End             ' プログラムを終了します．
    
                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10076, _
                        Description:="InBuffer - Write" & Chr$(&HA) & "ポート(" & Cn & ")の受信バッファの上限値 (" & fCOMMPROP.dwMaxRxQueue & "を越える設定は出来ません．"
                    Exit Property
            
            End Select
        End If
    End If
    
    InBuff = fCOMMPROP.dwCurrentRxQueue     '受信バッファの設定を読み出す
    OutBuff = fCOMMPROP.dwCurrentTxQueue    '送信バッファの設定を読み出す
    
    '新しいサイズの設定
    If SetupComm(ecH(Cn).Handle, InBuff, BufferSize) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 送信バッファサイズの書き込みに失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10072, _
                    Description:="OutBuffer - Write" & Chr$(&HA) & "ポート(" & Cn & ")の送信バッファサイズの書き込みに失敗しました"
                Exit Property
        
        End Select
        
    End If

End Property

'-----------------------------------
'   Asciiプロパティ
'-----------------------------------
'読み出し
Public Property Get Ascii() As String
    Dim ReadBytes As Long       '受信バイト数
    Dim ReadedBytes As Long     '読み込めたバイト数
    Dim bdata() As Byte         'バイナリ配列
    Dim Er As Long              ' error value
    Dim ErrorFlag As Boolean    'エラーフラグ

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10108, _
                    Description:="Ascii - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
    End If

    If ClearCommError(ecH(Cn).Handle, Er, fCOMSTAT) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 受信バッファの状態を取得できませんでした
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10100, _
                    Description:="Ascii - Read" & Chr$(&HA) & "ポート(" & Cn & ")の受信バッファの状態を取得できませんでした．"
                Exit Property
        
        End Select
    
    End If
    
    '--------- Version1.70で変更
'    ReadBytes = fCOMSTAT.cbInQue       '受信バイト数の取得
    If ec.AsciiBytes <= 0 Then
        ' 旧バージョン互換
        ReadBytes = fCOMSTAT.cbInQue    '受信バイト数の取得
      Else
        '--------- Version1.71で修正
'        ReadBytes = ec.AsciiBytes
        If ec.AsciiBytes <= fCOMSTAT.cbInQue Then
            ReadBytes = ec.AsciiBytes
          Else
            ReadBytes = fCOMSTAT.cbInQue    '受信バイト数の取得
        End If
        '--------- ここまで
    End If
    '--------- ここまで


    If ReadBytes = 0 Then       'データバッファが空のとき
        Ascii = ""
        Exit Property
    End If

    ReDim bdata(ReadBytes - 1)
    
    ErrorFlag = False
    
    If ReadFile(ecH(Cn).Handle, bdata(0), ReadBytes, ReadedBytes, 0&) = False Then ErrorFlag = True
    If ReadBytes <> ReadedBytes Then ErrorFlag = True
    If ErrorFlag Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 文字列受信に失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10101, _
                    Description:="Ascii - Read" & Chr$(&HA) & "ポート(" & Cn & ")からの文字列受信に失敗しました．"
                Exit Property
        
        End Select
    End If
    Ascii = StrConv(bdata, vbUnicode)

End Property

'書き込み
Public Property Let Ascii(TxD As String)
    Dim WriteBytes As Long      '送信バイト数
    Dim WrittenBytes As Long    'バッファに書き込めたバイト数
    Dim bdata() As Byte         'バイナリ配列
    Dim ErrorFlag As Boolean    'エラーフラグ
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10109, _
                    Description:="Ascii - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If
    
    bdata() = StrConv(TxD, vbFromUnicode)   'ANSIに変換
    WriteBytes = UBound(bdata()) + 1
    
    ErrorFlag = False

    If WriteFile(ecH(Cn).Handle, bdata(0), WriteBytes, WrittenBytes, 0&) = False Then
        ErrorFlag = True
    End If

    If WriteBytes <> WrittenBytes Then ErrorFlag = True
    If ErrorFlag Then
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 文字列送信に失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10102, _
                    Description:="Ascii - Write" & Chr$(&HA) & "ポート(" & Cn & ")からの文字列送信に失敗しました．"
                Exit Property
        
        End Select
    End If
End Property

'-----------------------------------
'   AsciiLineTimeOutプロパティ
'-----------------------------------
'Version1.51で追加
'AsciiLineプロパティの読み出し時のタイムアウトを設定，または設定値を読み出します．
'AsciiLineプロパティを読み出した時点から計測し，AsciiLineTimeOut(mS)を越えてもデリ
'ミタが受信できなかったときは処理を中止し，それまでに受信した文字をそのまま返します．
'エラーは発生しません．
'AsciiLineTimeOutプロパティはポートごとに設定します．
'初期値はゼロですが，ゼロ以下の値が設定されているとタイムアウトは発生しません．
'書き込み
Public Property Let AsciiLineTimeOut(TimeOut As Long)
    If Cn = 0 Then Exit Property
    ecH(Cn).LineInTimeOut = TimeOut
End Property
'読み出し
Public Property Get AsciiLineTimeOut() As Long
    If Cn = 0 Then
        AsciiLineTimeOut = 0
        Exit Property
      Else
        AsciiLineTimeOut = ecH(Cn).LineInTimeOut
    End If
End Property

'-----------------------------------
'   AsciiLineプロパティ
'-----------------------------------
'デリミタまでの文字列を受信
'デリミタはDelimiterプロパティで，ポートごとに設定・読み出しが可能です．
'Version1.51より，読み出しタイムアウトをサポートしました．
'タイムアウトはポートごとに設定する必要があります．
'AsciiLineプロパティを読み出したときから，デリミタを受信するまでの時間が設定値(mS)を
'越えるとタイムアウトが発生し，処理を終了します．
'ただしタイムアウトが発生してもエラーにはならず，そこまでに読み込んだ文字列をそのまま
'返します．
'タイムアウト値はAsciiLineの読み出し時のみ有効です．
'
'読み出し
Public Property Get AsciiLine() As String
    Dim ReadBytes As Long       '受信バイト数
    Dim ReadedBytes As Long     '読み込めたバイト数
    Dim bdata As Byte           'バイト変数
    Dim Bstr() As Byte          'バイト配列
    Dim Er As Long              ' error value
    Dim n As Long               '文字数カウント
    Dim DelimStr As String      'デリミタ
    Dim ErrorFlag As Boolean    'エラーフラグ
    
    '----Version 1.51
    Dim TimeOutFlag As Boolean  ' タイムアウトフラグ
    Dim STARTmS As Double       ' 開始時間(mS)
    Dim NOWmS As Double         ' 現在の時間(mS)

    STARTmS = GetTickCount      ' 開始時の時間を取得
    If STARTmS < 0 Then
        STARTmS = STARTmS + 4294967296#
    End If
    TimeOutFlag = False
    '----ここまで

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10218, _
                    Description:="AsciiLine - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If

    ' 設定値の正規化
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
            ecH(Cn).Delimiter = "CR"    ' 規定値以外はＣＲとみなします．
            DelimStr = "CR"
    End Select


    n = 0   ' 文字カウンタのリセット

    Do
        'データ受信まで待つ
        Do
            If ClearCommError(ecH(Cn).Handle, Er, fCOMSTAT) = False Then
                'エラー処理
                Select Case Xerror
                    Case Is = 0 '-----標準エラー
                        ec.CloseAll     ' すべてのポートを閉じます．
                        Stop            ' 受信バッファの状態を取得できませんでした
                        End             ' プログラムを終了します．

                    Case Is = 1 '-----トラップ可能エラー
                        err.Raise _
                            Number:=10210, _
                            Description:="AsciiLine - Read" & Chr$(&HA) & "ポート(" & Cn & ")の受信バッファの状態を取得できませんでした．"
                        Exit Property
                
                End Select
                
            End If
        
            ReadBytes = fCOMSTAT.cbInQue
            DoEvents

            If ecH(Cn).LineInTimeOut > 0 Then   ' タイムアウトがゼロ以下ならばスキップ
                ' 現在の時間を取得
                NOWmS = GetTickCount
                If NOWmS < 0 Then
                    NOWmS = NOWmS + 4294967296#
                End If
                'タイムアウトのチェック
                If STARTmS + ecH(Cn).LineInTimeOut <= NOWmS Then
                    ' タイムアウトした
                    TimeOutFlag = True
                    Exit Do
                End If
            End If
        Loop While ReadBytes = 0

        If TimeOutFlag Then Exit Do

        '１文字受信
        ErrorFlag = False
        
        If ReadFile(ecH(Cn).Handle, bdata, 1&, ReadedBytes, 0&) = False Then ErrorFlag = True
        If ReadedBytes <> 1 Then ErrorFlag = True
        If ErrorFlag Then
            '失敗
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' 文字列受信に失敗しました
                    End             ' プログラムを終了します．

                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10211, _
                        Description:="AsciiLine - Read" & Chr$(&HA) & "ポート(" & Cn & ")からの文字列受信に失敗しました．"
                    Exit Property
            
            End Select
        End If

        '受信文字の解析
        Select Case bdata
            Case Is = &HD   'Crを受信
                Select Case DelimStr
                    Case Is = "CR"
                        Exit Do                     ' デリミタを受信
                    Case Is = "LF", "CRLF"
                        ReDim Preserve Bstr(n)      ' 以前のデータを残したまま再定義
                        Bstr(n) = bdata             ' 受信文字に加える
                        n = n + 1
                    Case Is = "LFCR"
                        If n >= 1 Then
                            If Bstr(n - 1) = &HA Then   ' 一つ前の文字はLf?
                                n = n - 1               ' デリミタ文字を無効に
                                Exit Do                 ' デリミタを受信
                            End If
                          Else
                            ReDim Preserve Bstr(n)  ' 以前のデータを残したまま再定義
                            Bstr(n) = bdata         ' 受信文字に加える
                            n = n + 1
                        End If
                End Select
            
            Case Is = &HA   'Lfを受信
                
                Select Case DelimStr
                    Case Is = "CR", "LFCR"
                        ReDim Preserve Bstr(n)      ' 以前のデータを残したまま再定義
                        Bstr(n) = bdata             ' 受信文字に加える
                        n = n + 1
                    Case Is = "LF"                  ' デリミタを受信
                        Exit Do
                    Case Is = "CRLF"
                        If n >= 1 Then
                            If Bstr(n - 1) = &HD Then   ' 一つ前の文字はCr?
                                n = n - 1               ' デリミタ文字を無効に
                                Exit Do                 ' デリミタを受信
                            End If
                          Else
                            ReDim Preserve Bstr(n)  ' 以前のデータを残したまま再定義
                            Bstr(n) = bdata         ' 受信文字に加える
                            n = n + 1
                        End If
                End Select
                
            Case Else   '通常の文字
                ReDim Preserve Bstr(n)          ' 以前のデータを残したまま再定義
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

'指定された文字列にデリミタを付加して送信
'書き込み
Public Property Let AsciiLine(TxD As String)
    Dim WriteBytes As Long      '送信バイト数
    Dim WrittenBytes As Long    'バッファに書き込めたバイト数
    Dim bdata() As Byte         'バイナリ配列
    Dim DelimStr As String      'デリミタ
    Dim Td As String            '送信文字列
    Dim ErrorFlag As Boolean    'エラーフラグ
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10219, _
                    Description:="AsciiLine - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If

    ' 設定値の正規化と文字列にデリミタを付加
    
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
            ecH(Cn).Delimiter = "CR"        ' 規定値以外はＣＲとみなします．
            Td = TxD & Chr(&HD)
    End Select
    
    bdata() = StrConv(Td, vbFromUnicode)   'ANSIに変換
    
    WriteBytes = UBound(bdata()) + 1
    
    ErrorFlag = False
    
    If WriteFile(ecH(Cn).Handle, bdata(0), WriteBytes, WrittenBytes, 0&) = False Then ErrorFlag = True
    If WriteBytes <> WrittenBytes Then ErrorFlag = True
    If ErrorFlag Then
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 文字列送信に失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10212, _
                    Description:="AsciiLine - Write" & Chr$(&HA) & "ポート(" & Cn & ")からの文字列送信に失敗しました．"
                Exit Property
        
        End Select
    End If
End Property


'-----------------------------------
'   Binaryプロパティ
'-----------------------------------
'書き込み
Public Property Let Binary(ByteData As Variant)
    Dim WriteBytes As Long          '送信バイト数
    Dim WrittenBytes As Long        '読み込めたバイト数
    Dim bdata() As Byte             'バイナリ配列
    Dim Er As Long                  ' error value
    Dim ErrorFlag As Boolean        'エラーフラグ
    Dim i As Long
    Dim j As Long
    Dim C As Long

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' ポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10169, _
                    Description:="Binary - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
        
    End If

    '引数の型による処理の分岐
    Select Case TypeName(ByteData)
        Case "Byte"
            ReDim bdata(0)
            bdata(0) = ByteData
        Case "Integer", "Long", "Single", "Double"
            If ByteData < 0 Or ByteData > 255 Then
                'オーバーフローエラー
                Select Case Xerror
                    Case Is = 0 '-----標準エラー
                        ec.CloseAll     ' すべてのポートを閉じます．
                        Stop            ' 引数の値が0～255の範囲にありません．
                        End             ' プログラムを終了します．
        
                    Case Is = 1 '-----トラップ可能エラー
                        err.Raise _
                            Number:=10162, _
                            Description:="Binary - Write" & Chr$(&HA) & "引数の値が0～255の範囲にありません"
                        Exit Property

                End Select
            End If
            If ByteData <> Int(ByteData) Then
                ' 非整数エラー
                Select Case Xerror
                    Case Is = 0 '-----標準エラー
                        ec.CloseAll     ' すべてのポートを閉じます．
                        Stop            ' 引数の値が整数ではありません．
                        End             ' プログラムを終了します．
        
                    Case Is = 1 '-----トラップ可能エラー
                        err.Raise _
                            Number:=10163, _
                            Description:="Binary - Write" & Chr$(&HA) & "引数が整数ではありません"
                        Exit Property
                
                End Select
            End If
            ReDim bdata(0)
            bdata(0) = CByte(ByteData)
    
        Case "String"
            '文字列はユニコードのまま送信
            ReDim bdata(LenB(ByteData) - 1)
            For i = 0 To LenB(ByteData) - 1
                bdata(i) = AscB(MidB(ByteData, i + 1, 1))
            Next i
        
        Case "Byte()"
            bdata = ByteData
    
        Case "Integer()", "Long()", "Single()", "Double()", "Variant()"
            ReDim bdata(UBound(ByteData))   ' データ配列の再宣言
            ' バイト配列への代入とオーバーフロー，整数のチェック
            For i = 0 To UBound(ByteData)
                If ByteData(i) < 0 Or ByteData(i) > 255 Then
                'オーバーフローエラー
                    Select Case Xerror
                        Case Is = 0 '-----標準エラー
                            ec.CloseAll     ' すべてのポートを閉じます．
                            Stop            ' 引数の値が0～255の範囲にありません．
                            End             ' プログラムを終了します．
            
                        Case Is = 1 '-----トラップ可能エラー
                            err.Raise _
                                Number:=10162, _
                                Description:="Binary - Write" & Chr$(&HA) & "引数の値が0～255の範囲にありません"
                            Exit Property
    
                    End Select
                End If
                If ByteData(i) <> Int(ByteData(i)) Then
                    ' 非整数エラー
                    Select Case Xerror
                        Case Is = 0 '-----標準エラー
                            ec.CloseAll     ' すべてのポートを閉じます．
                            Stop            ' 引数の値が整数ではありません．
                            End             ' プログラムを終了します．
            
                        Case Is = 1 '-----トラップ可能エラー
                            err.Raise _
                                Number:=10163, _
                                Description:="Binary - Write" & Chr$(&HA) & "引数が整数ではありません"
                            Exit Property
                    
                    End Select
                End If
                ' データのセット
                bdata(i) = ByteData(i)
            Next i
        
        Case "String()"
            '文字配列はユニコードのまま送信
            C = 0   ' 現在のトータル文字数
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
            ' 非対応型
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' 指定された型には対応していません．
                    End             ' プログラムを終了します．
    
                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10164, _
                        Description:="Binary - Write" & Chr$(&HA) & "指定された型には対応していません"
                    Exit Property
            
            End Select
    
    End Select
    
    WriteBytes = UBound(bdata) + 1
    
    ErrorFlag = False
    
    If WriteFile(ecH(Cn).Handle, bdata(0), WriteBytes, WrittenBytes, 0&) = False Then ErrorFlag = True
    If WriteBytes <> WrittenBytes Then ErrorFlag = True
    If ErrorFlag Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' バイナリーデータの送信に失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10160, _
                    Description:="Binary - Write" & Chr$(&HA) & "ポート(" & Cn & ")からのバイナリーデータの送信に失敗しました．"
                Exit Property
        
        End Select
    End If
End Property

'読み出し
'Version1.70より，BinaryBytesプロパティで受信バッファから読み出すバイト数が指定できるようになりました．
'ただし，Bytesプロパティがゼロ以下の時は，すべてのデータを取得します(旧バージョン互換)
Public Property Get Binary() As Variant
    Dim ReadBytes As Long       '受信バイト数
    Dim ReadedBytes As Long     '読み込めたバイト数
    Dim bdata() As Byte         'バイナリ配列
    Dim Er As Long              ' error value
    Dim ErrorFlag As Boolean    'エラーフラグ
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10179, _
                    Description:="Binary - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
        
    End If
    
    ' ポートの状態を取得
    If ClearCommError(ecH(Cn).Handle, Er, fCOMSTAT) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 受信バッファの状態を取得できませんでした
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10172, _
                    Description:="Binary - Read" & Chr$(&HA) & "ポート(" & Cn & ")の受信バッファの状態を取得できませんでした．"
                Exit Property
        
        End Select
        
    End If

    '--------- Version1.70で変更
'    ReadBytes = fCOMSTAT.cbInQue       '受信バイト数の取得
    If ec.BinaryBytes <= 0 Then
        ' 旧バージョン互換
        ReadBytes = fCOMSTAT.cbInQue    '受信バイト数の取得
      Else
        '--------- Version1.71で修正
'        ReadBytes = ec.BinaryBytes
        If ec.AsciiBytes <= fCOMSTAT.cbInQue Then
            ReadBytes = ec.BinaryBytes
          Else
            ReadBytes = fCOMSTAT.cbInQue    '受信バイト数の取得
        End If
        '--------- ここまで
    End If
    '--------- ここまで
    
    If ReadBytes = 0 Then           'データバッファが空のとき
        Binary = 0                  ' 0 を返す
        Exit Property
    End If
    
    ReDim bdata(ReadBytes - 1)

    ErrorFlag = False

    If ReadFile(ecH(Cn).Handle, bdata(0), ReadBytes, ReadedBytes, 0&) = False Then ErrorFlag = True
    If ReadBytes <> ReadedBytes Then ErrorFlag = True
    If ErrorFlag Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' バイナリーデータの受信に失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10171, _
                    Description:="Binary - Read" & Chr$(&HA) & "ポート(" & Cn & ")からのバイナリーデータの受信に失敗しました．"
                Exit Property
        
        End Select
    End If
    Binary = bdata
End Property

'-----------------------------------
'   WAITmS プロパティ
'-----------------------------------
'指定した時間後に戻ります．
'長整数型(Long)，mS単位で指定します．
'最大49.7日まで指定できます．
Public Property Let WAITmS(WaitTime As Long)
    'mS単位のディレイ
    Dim STARTmS As Double   ' 開始時間(mS)
    Dim NOWmS As Double     ' 現在の時間(mS)

    '開始時の時間を取得
    STARTmS = GetTickCount
    If STARTmS < 0 Then
        STARTmS = STARTmS + 4294967296#
    End If

    '時間待ち
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
'   DozeSeconds プロパティ
'-----------------------------------
' 指定した時間(Sec)，処理を停止します．
' 停止する時間は秒単位で代入します．
' Dozeとは「うたた寝」という意味で，0.1秒ごとにDoEventsが実行されることから命名しました．
' Excel2000以降では有効ですがそれ以前のバージョンではほとんど意味がありません．
Public Property Let DozeSeconds(Seconds As Integer)
    Dim WakeUp As Date                  ' 目覚めの時刻
    If Seconds < 1 Then Exit Property   ' 1秒以上のみ有効
    WakeUp = Now + TimeSerial(0, 0, Seconds)
    Do
        DoEvents
        If Now >= WakeUp Then Exit Do
        ecDef.Sleep 100
    Loop
End Property

'-----------------------------------
'   RTSCTSプロパティ
'-----------------------------------
' RTSの強制制御とCTSの状態読み取り

'読み出し(CTSの状態)
Public Property Get RTSCTS() As Boolean
    Dim Stat As Long    ' Status

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10198, _
                    Description:="RTSCTS - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
    
    End If

    If GetCommModemStatus(ecH(Cn).Handle, Stat) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' CTSの状態が読み取れませんでした
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10190, _
                    Description:="RTSCTS - Read" & Chr$(&HA) & "ポート(" & Cn & ")のCTSの状態が読み取れませんでした．"
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

'書き込み
Public Property Let RTSCTS(Status As Boolean)
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10199, _
                    Description:="RTSCTS - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
        
    End If
    
    If Status = True Then
        Stat = SETRTS
      Else
        Stat = CLRRTS
    End If

    If EscapeCommFunction(ecH(Cn).Handle, Stat) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' RTSの設定に失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10191, _
                    Description:="RTSCTS - Write" & Chr$(&HA) & "ポート(" & Cn & ")のRTSの設定に失敗しました．"
                Exit Property
        
        End Select
    End If
End Property

'-----------------------------------
'   DTRDSRプロパティ
'-----------------------------------
' DTRの強制制御とDSRの状態読み取り

'読み出し(DSRの状態)
Public Property Get DTRDSR() As Boolean
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10208, _
                    Description:="DTRDSR - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
        
    End If
    
    If GetCommModemStatus(ecH(Cn).Handle, Stat) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' DSRの状態が読み取れませんでした
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10200, _
                    Description:="DTRDSR - Read" & Chr$(&HA) & "ポート(" & Cn & ")のDSRの状態が読み取れませんでした．"
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

'書き込み
Public Property Let DTRDSR(Status As Boolean)
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10209, _
                    Description:="DTRDSR - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property

        End Select
        
    End If
    
    If Status = True Then
        Stat = SETDTR
      Else
        Stat = CLRDTR
    End If

    If EscapeCommFunction(ecH(Cn).Handle, Stat) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' DTRの設定に失敗しました
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10201, _
                    Description:="DTRDSR - Write" & Chr$(&HA) & "ポート(" & Cn & ")のDTRの設定に失敗しました．"
                Exit Property
        
        End Select
        
    End If
End Property


'-----------------------------------
'   Delimiterプロパティ
'-----------------------------------
'AsciiLineプロパティで使用するデリミタの設定，読み出しを行うプロパティです．
'書き込み
Public Property Let Delimiter(DelimiterType As String)
    Dim DelimStr As String
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                Stop        ' 対象となるポート番号が指定されていません

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10248, _
                    Description:="Delimiter - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                End             ' プログラムを終了します．
        
        End Select

    End If

    ' 引数の正規化
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
            ecH(Cn).Delimiter = "CR"    ' 規定値以外はＣＲとみなします．
    End Select

End Property

'読み出し
Public Property Get Delimiter() As String
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10249, _
                    Description:="Delimiter - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If

    ' 設定値の読み取り
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
            Delimiter = "CR"                ' 規定値以外はＣＲとみなします．
            ecH(Cn).Delimiter = "CR"        ' CR に補正します
    End Select

End Property

'-----------------------------------
'   Breakプロパティ
'-----------------------------------
'ブレーク信号を送信，またはブレーク信号の送信を停止する書き込み専用のプロパティです．
Public Property Let Break(BreakOn As Boolean)
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10259, _
                    Description:="Break - Write" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property

        End Select
        
    End If
    
    If BreakOn = True Then
        ' Breakの送信
        If SetCommBreak(ecH(Cn).Handle) = False Then
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' Breakの送信に失敗しました
                End             ' プログラムを終了します．

                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10251, _
                        Description:="Break - Write" & Chr$(&HA) & "ポート(" & Cn & ")のブレーク送信に失敗しました．"
                    Exit Property
            
            End Select
        End If
      Else
        ' Break送信の停止
        If ClearCommBreak(ecH(Cn).Handle) = False Then
            'エラー処理
            Select Case Xerror
                Case Is = 0 '-----標準エラー
                    ec.CloseAll     ' すべてのポートを閉じます．
                    Stop            ' Break送信の停止に失敗しました
                    End             ' プログラムを終了します．

                Case Is = 1 '-----トラップ可能エラー
                    err.Raise _
                        Number:=10252, _
                        Description:="Break - Write" & Chr$(&HA) & "ポート(" & Cn & ")のブレーク送信の停止に失敗しました．"
                    Exit Property
            
            End Select
        End If
          
    End If

End Property

'-----------------------------------
'   RIプロパティ(読取専用)
'-----------------------------------
' RIの状態読み取り
Public Property Get RI() As Boolean
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10268, _
                    Description:="RI - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
        
    End If
    
    If GetCommModemStatus(ecH(Cn).Handle, Stat) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' RIの状態が読み取れませんでした
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10261, _
                    Description:="RI - Read" & Chr$(&HA) & "ポート(" & Cn & ")のRIの状態が読み取れませんでした．"
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
'   DCDプロパティ(読取専用)
'-----------------------------------
' DCDの状態読み取り
Public Property Get DCD() As Boolean
    Dim Stat As Long    ' Status
    
    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10278, _
                    Description:="DCD - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select
        
    End If
    
    If GetCommModemStatus(ecH(Cn).Handle, Stat) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' DCDの状態が読み取れませんでした
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10271, _
                    Description:="DCD - Read" & Chr$(&HA) & "ポート(" & Cn & ")のDCDの状態が読み取れませんでした．"
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
'   Specプロパティ
'-----------------------------------
'ポートハンドル ecH(Cn).Handle の情報を文字列で返します．
Public Property Get Spec() As String
    Dim Er As Long
    Dim Mes As String
    Dim CrLf As String
    'CRLF = Chr$(&HD) & Chr$(&HA)
    CrLf = Chr$(&HA)

    If Cn = 0 Then
        'ポート番号未指定
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' 対象となるポート番号が指定されていません
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=10238, _
                    Description:="InBuffer - Read" & Chr$(&HA) & "対象となるポート番号が指定されていません"
                Exit Property
        
        End Select

    End If

    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        'エラー処理
        Select Case Xerror
            Case Is = 0 '-----標準エラー
                ec.CloseAll     ' すべてのポートを閉じます．
                Stop            ' ポートの情報を取得できませんでした
                End             ' プログラムを終了します．

            Case Is = 1 '-----トラップ可能エラー
                err.Raise _
                    Number:=12300, _
                    Description:="InCount - Read" & Chr$(&HA) & "ポート(" & Cn & ")の情報を取得できませんでした"
                Exit Property
        
        End Select

    End If

    'メンバの確認
    With fCOMMPROP
        Mes = "構造体のバイトサイズ" & CrLf
        Mes = Mes & "　" & .wPacketLength & CrLf
        Mes = Mes & "バージョン" & CrLf
        Mes = Mes & "　" & .wPacketVersion & CrLf
        Mes = Mes & "送信バッファの最大バイト数" & CrLf
        If .dwMaxTxQueue = 0 Then
            Mes = Mes & "　制限無し" & CrLf
          Else
            Mes = Mes & "　" & .dwMaxTxQueue & CrLf
        End If
        Mes = Mes & "受信バッファの最大バイト数" & CrLf
        If .dwMaxRxQueue = 0 Then
            Mes = Mes & "　制限無し" & CrLf
          Else
            Mes = Mes & "　" & .dwMaxRxQueue & CrLf
        End If
        Mes = Mes & "最大データ転送速度" & CrLf & "　"
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
                Mes = Mes & "プログラマブル"
        End Select
        Mes = Mes & CrLf

        Mes = Mes & "サポートされている機能" & CrLf
        
        If .dwProvCapabilities & PCF_DTRDSR <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：DTR/DSR" & CrLf
        
        If .dwProvCapabilities & PCF_RTSCTS <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：RTS/CTS" & CrLf
        
        If .dwProvCapabilities & PCF_RLSD <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：CD(RLSD)" & CrLf
        
        If .dwProvCapabilities & PCF_PARITY_CHECK <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：パリティチェック" & CrLf
        
        If .dwProvCapabilities & PCF_XONXOFF <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：XON/XOFFによるフロー制御" & CrLf
        
        If .dwProvCapabilities & PCF_SETXCHAR <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：XON/XOFFの文字指定" & CrLf
        
        If .dwProvCapabilities & PCF_TOTALTIMEOUTS <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：トータルタイムアウトの設定" & CrLf
        
        If .dwProvCapabilities & PCF_INTTIMEOUTS <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：インターバルタイムアウトの設定" & CrLf
        
        If .dwProvCapabilities & PCF_SPECIALCHARS <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：特殊文字の使用" & CrLf
       
        If .dwProvCapabilities & PCF_16BITMODE <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：特殊な16ビットモード" & CrLf

        Mes = Mes & "各種機能の設定の可否" & CrLf
        
        If .dwSettableParams & SP_PARITY <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：パリティのモード" & CrLf
        
        If .dwSettableParams & SP_BAUD <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：ボーレート" & CrLf

        If .dwSettableParams & SP_DATABITS <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：データビット数" & CrLf

        If .dwSettableParams & SP_STOPBITS <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：ストップビット数" & CrLf

        If .dwSettableParams & SP_HANDSHAKING <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：フロー制御（ハンドシェーク）" & CrLf

        If .dwSettableParams & SP_PARITY_CHECK <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：パリティチェックのON/OFF" & CrLf

        If .dwSettableParams & SP_RLSD <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：CD(RLSD)" & CrLf

        If .dwSettableParams & SP_PARITY <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：パリティモード" & CrLf

        Mes = Mes & "設定可能なボーレート" & CrLf
        
        If .dwSettableBaud And BAUD_075 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：75 bps" & CrLf
        
        If .dwSettableBaud And BAUD_110 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：110 bps" & CrLf
        
        If .dwSettableBaud And BAUD_134_5 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：134.5 bps" & CrLf
        
        If .dwSettableBaud And BAUD_150 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：150 bps" & CrLf
            
        If .dwSettableBaud And BAUD_300 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：300 bps" & CrLf
            
        If .dwSettableBaud And BAUD_600 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：600 bps" & CrLf
            
        If .dwSettableBaud And BAUD_1200 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：1200 bps" & CrLf
            
        If .dwSettableBaud And BAUD_1800 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：1800 bps" & CrLf
            
        If .dwSettableBaud And BAUD_2400 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：2400 bps" & CrLf
            
        If .dwSettableBaud And BAUD_4800 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：4800 bps" & CrLf
            
        If .dwSettableBaud And BAUD_7200 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：7200 bps" & CrLf
            
        If .dwSettableBaud And BAUD_9600 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：9600 bps" & CrLf
            
        If .dwSettableBaud And BAUD_14400 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：14400 bps" & CrLf
            
        If .dwSettableBaud And BAUD_19200 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：19200 bps" & CrLf
            
        If .dwSettableBaud And BAUD_38400 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：38400 bps" & CrLf
            
        If .dwSettableBaud And BAUD_56K <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：56 K bps" & CrLf
            
        If .dwSettableBaud And BAUD_57600 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：57600 bps" & CrLf
            
        If .dwSettableBaud And BAUD_115200 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：115200 bps" & CrLf
            
        If .dwSettableBaud And BAUD_128K <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：128 K bps" & CrLf
            
        If .dwSettableBaud And BAUD_USER <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：プログラム" & CrLf

        Mes = Mes & "設定可能なデータビット数" & CrLf
        
        If .wSettableData And DATABITS_5 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：5 ビット" & CrLf

        If .wSettableData And DATABITS_6 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：6 ビット" & CrLf
        
        If .wSettableData And DATABITS_7 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：7 ビット" & CrLf
        
        If .wSettableData And DATABITS_8 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：8 ビット" & CrLf
        
        If .wSettableData And DATABITS_16 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：16 ビット" & CrLf
        
        If .wSettableData And DATABITS_16X <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：特殊なワイドパス" & CrLf

        Mes = Mes & "設定可能なストップビット数" & CrLf
        
        If .wSettableStopParity And STOPBITS_10 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：1 ビット" & CrLf
        
        If .wSettableStopParity And STOPBITS_15 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：1.5 ビット" & CrLf
        
        If .wSettableStopParity And STOPBITS_20 <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：2 ビット" & CrLf

        Mes = Mes & "設定可能なパリティチェック" & CrLf
        
        If .wSettableStopParity And PARITY_NONE <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：パリティなし" & CrLf

        If .wSettableStopParity And PARITY_ODD <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：奇数パリティ" & CrLf

        If .wSettableStopParity And PARITY_EVEN <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：偶数パリティ" & CrLf

        If .wSettableStopParity And PARITY_MARK <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：マークパリティ" & CrLf

        If .wSettableStopParity And PARITY_SPACE <> 0 Then
            Mes = Mes & "　○"
          Else
            Mes = Mes & "　×"
        End If
        Mes = Mes & "：スペースパリティ" & CrLf

    End With
    Spec = Mes
    
End Property
'-----------------------------------
'   CloseAllメソッド
'-----------------------------------
Private Sub CloseAll()
'すべてのポートを閉じる，内部処理用のメソッドです．
'ecPorts.tmpに保存されているハンドルもクローズし，ファイルを削除します．
'エラーを返しません
    Dim i As Long
    Dim rv As Long
    Dim Handle As Long
    Dim FileNumber As Integer
    Dim Fpath As String * 260
    FileNumber = FreeFile()
    
    GetTempPath 260, Fpath
    Open Left(Fpath, InStr(Fpath, vbNullChar) - 1) & "ecPort.tmp" For Random Access Read Write As #FileNumber Len = Len(Handle)

    For i = 1 To ecMaxPort
        ecH(i).Delimiter = "CR"         ' デリミタをリセットします
        ecH(i).LineInTimeOut = 0        ' タイムアウトのリセット
        Get #FileNumber, i, Handle      ' 記録されているハンドルを取得します
        If Handle > 0 Then
            CloseHandle Handle          ' クローズ
            rv = 0&
            Put #FileNumber, i, rv
        End If
        ecH(i).Handle = 0
    Next i
    Cn = 0      ' 処理対象のポート番号をリセット
    Close #FileNumber
End Sub

'-----------------------------------
'隠しコマンド

'-----------------------------------
'   OutBufferSizeプロパティ
'-----------------------------------
'読み出し専用
'現在処理の対象となっているポートの送信バッファのサイズを返します
'エラー時は-1を返します
Private Property Get OutBufferSize() As Long
    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        '状態の取得に失敗
        OutBufferSize = -1
        Exit Property
    End If
    OutBufferSize = ecDef.fCOMMPROP.dwCurrentTxQueue
End Property

'-----------------------------------
'   InBufferSizeプロパティ
'-----------------------------------
'読み出し専用
'現在処理の対象となっているポートの受信バッファのサイズを返します
'エラー時は-1を返します
Private Property Get InBufferSize() As Long
    If GetCommProperties(ecH(Cn).Handle, fCOMMPROP) = False Then
        '状態の取得に失敗
        InBufferSize = -1
        Exit Property
    End If
    InBufferSize = ecDef.fCOMMPROP.dwCurrentRxQueue
End Property

'-----------------------------------
'   DELIMsプロパティ
'-----------------------------------
'DELIMs
'デリミタ用設定文字列取得プロパティ(読み出し専用)
'   Delimiterプロパティで使用するデリミタの設定用文字列を取得するためのプロパティ．
'   次の例は，文字列で指定するものとHANDSHAKINGsプロパティを使ったものとの比較です．
'例
' ●デリミタをCrに設定します．
'   ec.Delimiter = ec.DELIMs.Cr
'   ec.Delimiter = "Cr"
' ●デリミタをCr+Lfに設定します．
'   ec.Delimiter = ec.DELIMs.CrLf
'   ec.Delimiter = "CRLF"

Public Property Get DELIMs() As DelimType
    ' デリミタ指定文字列定数
    DELIMs.Cr = "CR"            ' CR
    DELIMs.Lf = "LF"            ' LF
    DELIMs.CrLf = "CRLF"        ' CR + LF
    DELIMs.LfCr = "LFCR"        ' LF + CR
End Property

'-----------------------------------
'   HANDSHAKEsプロパティ
'-----------------------------------
'ハンドシェーキング設定文字列取得プロパティ(読み出し専用)
'   HandHsakingプロパティで使用する設定用文字列を取得するためのプロパティ．
'   次の例は，文字列で指定するものとHANDSHAKINGsプロパティを使ったものとの比較です．
'例
' ●ハンドシェークをなしに設定します．
'   ec.HandShaking = ec.HANDSHAKEs.No
'   ec.HandShaking = "N"
' ●ハンドシェークをRTS/CTSに設定します．
'   ec.HandShaking = ec.HANDSHAKEs.RTSCTS
'   ec.HandShaking = "R"
'
Public Property Get HANDSHAKEs() As HandShakingType
    ' ハンドシェーク指定文字列定数
    HANDSHAKEs.No = "N"         ' なし
    HANDSHAKEs.XonXoff = "X"    ' Xon/Off
    HANDSHAKEs.RTSCTS = "R"     ' RTS/CTS
    HANDSHAKEs.DTRDSR = "D"     ' DTR/DSR
End Property


