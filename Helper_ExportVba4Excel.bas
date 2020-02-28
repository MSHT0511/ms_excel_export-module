Attribute VB_Name = "Helper_ExportVba4Excel"
Public Function ExportVbaProgramCode()

  Dim vbcmp As Object
  Dim strFileName As String
  Dim strExt As String
'  Set dbs = CurrentDb
'
'Debug.Print ThisWorkbook.Path
'Debug.Print ThisWorkbook.Name
'Exit Function
'  savepath = CurrentProject.Path & "\_VBA_" & Mid(dbs.Name, InStrRev(dbs.Name, "\") + 1) & "\" 'on ACCESS VBA
  savepath = ThisWorkbook.Path & "\_VBA_" & Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "\") + 1) & "\" 'on EXCEL VBA
  
    If Dir(savepath, vbDirectory) = "" Then
        MkDir savepath
    End If
  
  For Each vbcmp In ThisWorkbook.VBProject.VBComponents
    With vbcmp
      
      '// 出力先ファイルパス
      strFileName = savepath & .Name
'      Debug.Print .Type & " " & .Name
'      Debug.Print strFileName
      '拡張子を設定
      Select Case .Type
        Case 1    '標準モジュールの場合
          strExt = ".bas"
        Case 2    'クラスモジュールの場合
          strExt = ".cls"
        Case 3 'ユーザーフォーム
          strExt = ".frm"
        Case 100  'フォーム/レポートのモジュールの場合
          strExt = ".cls"
      End Select
      'モジュールをエクスポート
      .Export strFileName & strExt
    End With
  Next vbcmp
  
'  Set dbs = Nothing
MsgBox "プログラムコード出力完了"
End Function
