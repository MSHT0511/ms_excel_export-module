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
      
      '// �o�͐�t�@�C���p�X
      strFileName = savepath & .Name
'      Debug.Print .Type & " " & .Name
'      Debug.Print strFileName
      '�g���q��ݒ�
      Select Case .Type
        Case 1    '�W�����W���[���̏ꍇ
          strExt = ".bas"
        Case 2    '�N���X���W���[���̏ꍇ
          strExt = ".cls"
        Case 3 '���[�U�[�t�H�[��
          strExt = ".frm"
        Case 100  '�t�H�[��/���|�[�g�̃��W���[���̏ꍇ
          strExt = ".cls"
      End Select
      '���W���[�����G�N�X�|�[�g
      .Export strFileName & strExt
    End With
  Next vbcmp
  
'  Set dbs = Nothing
MsgBox "�v���O�����R�[�h�o�͊���"
End Function
