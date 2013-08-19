VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///    WAVE�t�@�C���Đ��N���X


'Wavȩ�ق̍Đ��E��~���s�Ȃ�API
Private Declare Function sndPlaySound Lib "winmm.dll" _
        Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long

'sndPlaySound�֐���uFlags�Ŏg�p����萔
Const SND_SYNC = &H0         '�Đ��I����A�����߂�
Const SND_ASYNC = &H1        '�֐����s��A�����ɐ����߂�
Const SND_LOOP = &H8         '��~�𖽗߂���܂ōĐ����J��Ԃ�
                             '(SND_ASYNC���ꏏ�Ɏw�肷��K�v������܂�)
Const SND_NODEFAULT = &H2    '�w�肵��Wavȩ�ق�������Ȃ������ꍇ�A��̫�Ă�Wavȩ�ق��Đ����Ȃ�
Const SND_MEMORY = &H4       '��؏��Wavȩ�ق��Đ�����
Const SND_NOSTOP = &H10      '����Wavȩ�ق��Đ����̏ꍇ�A�Đ��𒆎~����

Private mWavFile As String


Public Sub PlayWave(ByVal sWaveFile As String, Optional bLoop As Boolean = False)
    If sWaveFile = "" Then Exit Sub

    Select Case bLoop
        Case True
            '�w�肳�ꂽWavȩ�ق��J��Ԃ��Đ�����
            Call sndPlaySound(sWaveFile, SND_ASYNC Or SND_LOOP)
        Case False
            '�w�肳�ꂽWavȩ�ق��Đ�����
            Call sndPlaySound(sWaveFile, SND_ASYNC)
    
    End Select
    
End Sub

Public Sub StopWave()
    'Wavȩ�ق̍Đ����~����
    Call sndPlaySound("", 0)

End Sub


Public Property Get WavFile() As String
    WavFile = mWavFile
End Property
Public Property Let WavFile(ByVal sNewValue As String)
    mWavFile = sNewValue
End Property