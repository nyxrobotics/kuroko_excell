Attribute VB_Name = "export_modules_for_git"
Sub ExportAll()
    Dim module                  As VBComponent      '// モジュール
    Dim moduleList              As VBComponents     '// VBAプロジェクトの全モジュール
    Dim extension                                   '// モジュールの拡張子
    Dim sPath                                       '// 処理対象ブックのパス
    Dim sFilePath                                   '// エクスポートファイルパス
    Dim TargetBook                                  '// 処理対象ブックオブジェクト
    
    '// ブックが開かれていない場合は個人用マクロブック（personal.xlsb）を対象とする
    If (Workbooks.count = 1) Then
        Set TargetBook = ThisWorkbook
    '// ブックが開かれている場合は表示しているブックを対象とする
    Else
        Set TargetBook = ActiveWorkbook
    End If
    
    sPath = TargetBook.Path
    
    '// 処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents
    
    '// VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        '// クラス
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        '// フォーム
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// .frxも一緒にエクスポートされる
            extension = "frm"
        '// 標準モジュール
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        '// その他
        Else
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If
        
        '// エクスポート実施
        sFilePath = sPath & "\" & module.Name & "." & extension
        Call module.Export(sFilePath)
        convertCharCode_SJIS_to_utf8 (sFilePath)
        
        '// 出力先確認用ログ出力
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub

' ファイルの文字コードをSJISからUTF8(BOM無し)に変換する
Private Sub convertCharCode_SJIS_to_utf8(file As String)
    Dim destWithBOM As Object: Set destWithBOM = CreateObject("ADODB.Stream")
    With destWithBOM
        .Type = 2
        .Charset = "utf-8"
        .Open
        
        ' ファイルをSJIS で開いて、dest へ 出力
        With CreateObject("ADODB.Stream")
            .Type = 2
            .Charset = "shift-jis"
            .Open
            .LoadFromFile file
            .Position = 0
            .copyTo destWithBOM
            .Close
        End With
        
        ' BOM消去
        ' 3バイト無視してからバイナリとして出力
        .Position = 0
        .Type = 1 ' adTypeBinary
        .Position = 3
        
        Dim dest: Set dest = CreateObject("ADODB.Stream")
        With dest
            .Type = 1 ' adTypeBinary
            .Open
            destWithBOM.copyTo dest
            .savetofile file, 2
            .Close
        End With
        
        .Close
    End With
End Sub

Public Property Get isExportSelf() As Boolean
    isExportSelf = exportSelf
End Property

Public Property Let isExportSelf(ByVal vNewValue As Boolean)
    exportSelf = vNewValue
End Property
