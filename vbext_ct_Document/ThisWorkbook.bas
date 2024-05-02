VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'ブック保存時の処理
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Call ExportAll
End Sub

'ブックを閉じる時の処理
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ec.COMn = 0
End Sub
