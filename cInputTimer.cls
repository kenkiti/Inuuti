VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInputTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///   ���̓^�C�}�[�N���X

Private mTanngoYomi     As String '�P��ǂ�
Private mTanngoKanji    As String '�P�ꊿ��
Private mTime           As String '��������
Private mInputString    As String '���͒��̕�����
Private mNokoriString   As String '�c��̓��͂��镶����
Private mInputMiss      As Long   '���̓~�X��

Private WithEvents cTime    As Timer
Attribute cTime.VB_VarHelpID = -1
Private lblTanngoKanji      As Label
Private lblTanngoYomi       As Label


'�C�x���g
Public Event InputEnd(ByVal bOK As Boolean)  '���͏I��
Public Event InputMiss() '���̓~�X
Public Event InputUpdate(ByVal strTanngo As String)

Const INP_OK As Boolean = True
Const INP_NG As Boolean = False


Public Sub SetTimerControl(objTimer As Timer)
    Set cTime = objTimer
    
    Call InitInput

End Sub

Private Sub Class_Terminate()
    Set cTime = Nothing
End Sub


'�������Ԍo�������̏���
Private Sub cTime_Timer()
    
    '�^�C�}�[�X�g�b�v
    Call InitInput
    
    '�C�x���g����
    If mNokoriString = "" Then
        RaiseEvent InputEnd(INP_OK)
    Else
        RaiseEvent InputEnd(INP_NG)
    End If


End Sub


'///    ��������֐�

'���͊J�n
Public Sub GekiutiStart()
    '���͕�����̃Z�b�g
    mNokoriString = mTanngoYomi
    
    '�^�C�}�[�J�n
    cTime.Enabled = True

End Sub

'�ꕶ������
Public Sub InputChar(ByRef sChar As String)
    If cTime.Enabled = False Then Exit Sub
    
    If UCase(Left(mNokoriString, 1)) = UCase(sChar) Then
        '���͏o������1�������炷
        mNokoriString = Mid(mNokoriString, 2, Len(mNokoriString) - 1)
        
        '���͍ςݕ������ǉ�
        mInputString = mInputString & sChar
        RaiseEvent InputUpdate(mInputString)
    
        '���͏I���̃`�F�b�N
        If mNokoriString = "" Then
            
            Call InitInput
            RaiseEvent InputEnd(INP_OK)
        
        End If
    
    Else
        '���̓~�X�J�E���^���C���N�������g
        mInputMiss = mInputMiss + 1
        
        RaiseEvent InputMiss
    
    End If

End Sub


Private Sub InitInput()
    cTime.Enabled = False
    cTime.Interval = 0
    mInputString = ""

End Sub


'///    ��������v���p�e�B

'���͕�����i�ǂ݁j�v���p�e�B
Public Property Get TanngoYomi() As String
    TanngoYomi = mTanngoYomi
End Property
Public Property Let TanngoYomi(ByVal sNewValue As String)
    mTanngoYomi = sNewValue
End Property

'���͕�����i�����j�v���p�e�B
Public Property Get TanngoKanji() As String
    TanngoKanji = mTanngoKanji
End Property
Public Property Let TanngoKanji(ByVal sNewValue As String)
    mTanngoKanji = sNewValue
End Property

'�������ԃv���p�e�B
Public Property Get TimeInterval() As Long
    TimeInterval = cTime.Interval
End Property
Public Property Let TimeInterval(ByVal lNewValue As Long)
    cTime.Interval = lNewValue
End Property

''���ł��������ʂ���t���O
'Public Property Get NowGekiuti() As Boolean
'    NowGekiuti = cTime.Enabled
'End Property
'Public Property Let NowGekiuti(ByVal bNewValue As Boolean)
'    '�擾��p
'End Property