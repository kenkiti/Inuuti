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

'///    WAVEファイル再生クラス


'Waveﾌｧｲﾙの再生・停止を行なうAPI
Private Declare Function sndPlaySound Lib "winmm.dll" _
        Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long

'sndPlaySound関数のuFlagsで使用する定数
Const SND_SYNC = &H0         '再生終了後、制御を戻す
Const SND_ASYNC = &H1        '関数実行後、直ちに制御を戻す
Const SND_LOOP = &H8         '停止を命令するまで再生を繰り返す
                             '(SND_ASYNCも一緒に指定する必要があります)
Const SND_NODEFAULT = &H2    '指定したWaveﾌｧｲﾙが見つからなかった場合、ﾃﾞﾌｫﾙﾄのWaveﾌｧｲﾙを再生しない
Const SND_MEMORY = &H4       'ﾒﾓﾘ上のWaveﾌｧｲﾙを再生する
Const SND_NOSTOP = &H10      '現在Waveﾌｧｲﾙが再生中の場合、再生を中止する

Private mWavFile As String


Public Sub PlayWave(ByVal sWaveFile As String, Optional bLoop As Boolean = False)
    If sWaveFile = "" Then Exit Sub

    Select Case bLoop
        Case True
            '指定されたWaveﾌｧｲﾙを繰り返し再生する
            Call sndPlaySound(sWaveFile, SND_ASYNC Or SND_LOOP)
        Case False
            '指定されたWaveﾌｧｲﾙを再生する
            Call sndPlaySound(sWaveFile, SND_ASYNC)
    
    End Select
    
End Sub

Public Sub StopWave()
    'Waveﾌｧｲﾙの再生を停止する
    Call sndPlaySound("", 0)

End Sub


Public Property Get WavFile() As String
    WavFile = mWavFile
End Property
Public Property Let WavFile(ByVal sNewValue As String)
    mWavFile = sNewValue
End Property
