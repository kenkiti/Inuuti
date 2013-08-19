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

'///   入力タイマークラス

Private mTanngoYomi     As String '単語読み
Private mTanngoKanji    As String '単語漢字
Private mTime           As String '制限時間
Private mInputString    As String '入力中の文字列
Private mNokoriString   As String '残りの入力する文字列
Private mInputMiss      As Long   '入力ミス回数

Private WithEvents cTime    As Timer
Attribute cTime.VB_VarHelpID = -1
Private lblTanngoKanji      As Label
Private lblTanngoYomi       As Label


'イベント
Public Event InputEnd(ByVal bOK As Boolean)  '入力終了
Public Event InputMiss() '入力ミス
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


'制限時間経った時の処理
Private Sub cTime_Timer()
    
    'タイマーストップ
    Call InitInput
    
    'イベント発生
    If mNokoriString = "" Then
        RaiseEvent InputEnd(INP_OK)
    Else
        RaiseEvent InputEnd(INP_NG)
    End If


End Sub


'///    ここから関数

'入力開始
Public Sub GekiutiStart()
    '入力文字列のセット
    mNokoriString = mTanngoYomi
    
    'タイマー開始
    cTime.Enabled = True

End Sub

'一文字入力
Public Sub InputChar(ByRef sChar As String)
    If cTime.Enabled = False Then Exit Sub
    
    If UCase(Left(mNokoriString, 1)) = UCase(sChar) Then
        '入力出来たら1文字減らす
        mNokoriString = Mid(mNokoriString, 2, Len(mNokoriString) - 1)
        
        '入力済み文字列を追加
        mInputString = mInputString & sChar
        RaiseEvent InputUpdate(mInputString)
    
        '入力終了のチェック
        If mNokoriString = "" Then
            
            Call InitInput
            RaiseEvent InputEnd(INP_OK)
        
        End If
    
    Else
        '入力ミスカウンタをインクリメント
        mInputMiss = mInputMiss + 1
        
        RaiseEvent InputMiss
    
    End If

End Sub


Private Sub InitInput()
    cTime.Enabled = False
    cTime.Interval = 0
    mInputString = ""

End Sub


'///    ここからプロパティ

'入力文字列（読み）プロパティ
Public Property Get TanngoYomi() As String
    TanngoYomi = mTanngoYomi
End Property
Public Property Let TanngoYomi(ByVal sNewValue As String)
    mTanngoYomi = sNewValue
End Property

'入力文字列（漢字）プロパティ
Public Property Get TanngoKanji() As String
    TanngoKanji = mTanngoKanji
End Property
Public Property Let TanngoKanji(ByVal sNewValue As String)
    mTanngoKanji = sNewValue
End Property

'制限時間プロパティ
Public Property Get TimeInterval() As Long
    TimeInterval = cTime.Interval
End Property
Public Property Let TimeInterval(ByVal lNewValue As Long)
    cTime.Interval = lNewValue
End Property

''激打ち中か判別するフラグ
'Public Property Get NowGekiuti() As Boolean
'    NowGekiuti = cTime.Enabled
'End Property
'Public Property Let NowGekiuti(ByVal bNewValue As Boolean)
'    '取得専用
'End Property
