Attribute VB_Name = "mdlMain"
Option Explicit


Public Const RES_BMP_NOW   As Long = 101
Public Const RES_BMP_OK As Long = 102
Public Const RES_BMP_NG As Long = 103
Public Const RES_BMP_END As Long = 104
Public Const RES_BMP_START As Long = 105


Sub Main()

    frmSplash.Show vbModeless
    DoEvents
    
    Load frmMain
    
    Unload frmSplash
    
    frmMain.Show vbModeless

End Sub

'-----------------------------------------------------------
' SUB: AddDirSep
' �߽�̖������ިڸ�؋�؂�L���̉~�L�� (\) ���Ȃ��ꍇ�A
' �~�L����ǉ����܂��B
'
' ����/�o�͈���: [strPathName] - �~�L����ǉ������߽
'-----------------------------------------------------------
'
Public Function AddDirSep(strPathName As String) As String
    If Right(Trim(strPathName), 1) <> "\" Then
        strPathName = RTrim$(strPathName) & "\"
    End If

    AddDirSep = strPathName

End Function

