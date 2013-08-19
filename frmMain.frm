VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7545
   StartUpPosition =   2  '画面の中央
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   2640
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "開始(&S)"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.PictureBox picKey 
      Height          =   1935
      Left            =   1130
      ScaleHeight     =   1875
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   4560
      Width           =   5535
   End
   Begin VB.PictureBox picMain 
      Align           =   1  '上揃え
      BackColor       =   &H00000000&
      Height          =   4560
      Left            =   0
      ScaleHeight     =   4500
      ScaleWidth      =   7485
      TabIndex        =   3
      Top             =   0
      Width           =   7545
      Begin VB.Image imgMain 
         Height          =   3255
         Left            =   2160
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblTanngoYomi 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   3840
         Width           =   5295
      End
      Begin VB.Label lblTanngoKanji 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＤＦ特太ゴシック体"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1080
         TabIndex        =   5
         Top             =   3480
         Width           =   5535
      End
      Begin VB.Label lblInput 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   4200
         Width           =   5535
      End
   End
   Begin VB.Label lblKeyPush 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'StretchBlt API
Private Declare Function StretchBlt Lib "gdi32" _
        (ByVal hdc As Long, _
         ByVal x As Long, ByVal y As Long, _
         ByVal nWidth As Long, ByVal nHeight As Long, _
         ByVal hSrcDC As Long, _
         ByVal xSrc As Long, ByVal ySrc As Long, _
         ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
         ByVal dwRop As Long) As Long

Private mclsKeyBoard            As cKeyBoard    'キーボード描画クラス
Private mclsTanngo              As cTanngo      '単語取得クラス
Private mclsPlayWave            As cWave        'WAVE再生クラス
Private WithEvents mclsTimer    As cInputTimer  '入力チェッククラス
Attribute mclsTimer.VB_VarHelpID = -1


Private mNowGekiuti             As Boolean '激打ち中か判別するフラグ
Private mTanngoCounter          As Long    '単語カウンタ

Const INP_OK As Long = 0
Const INP_NG As Long = 1
Const INP_START As Long = 2
Const INP_END As Long = 3


'単語入力開始
Private Sub SetTanngo(lTime As Long, strKanji As String, strYomi As String)
    
    'ラベルに単語を表示
    lblTanngoKanji.Caption = strKanji
    lblTanngoYomi.Caption = strYomi
    lblInput.Caption = ""
    
    Set imgMain.Picture = LoadResPicture(RES_BMP_NOW, vbResBitmap)
    
    mclsTimer.TanngoYomi = strYomi
    mclsTimer.TimeInterval = lTime
    mclsTimer.GekiutiStart
    
    '激打ち開始！
    mNowGekiuti = True

End Sub

Private Sub cmdStart_Click()
    mTanngoCounter = 0
    cmdStart.Enabled = False
        
    Call DispMessage(INP_START)
        
    Call NextTanngo

End Sub

'次の単語表示
Private Sub NextTanngo()
    If mNowGekiuti Then Exit Sub
    
    Dim lTime As Long
    Dim sKanji As String
    Dim sYomi As String
    Dim bRet As Boolean
    
    bRet = mclsTanngo.GetTango(mTanngoCounter, lTime, sKanji, sYomi)
    If bRet Then
        Call SetTanngo(lTime, sKanji, sYomi)
    
        mTanngoCounter = mTanngoCounter + 1

    Else
        'ゲーム終了

        Call DispMessage(INP_END)
        
        cmdStart.Enabled = True
    
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If mNowGekiuti = False Then Exit Sub
    
    Dim sChar As String
    lblKeyPush.Caption = Chr(KeyCode) & ":Cd=" & CStr(KeyCode)

    sChar = mclsKeyBoard.PushKeyBoardDown(KeyCode)
    Call mclsTimer.InputChar(sChar)

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If mNowGekiuti = False Then Exit Sub
    
    Call mclsKeyBoard.PushKeyBoardUp(KeyCode)

End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Me.Caption = "～犬打ち～"
    
    'キーボード描画クラスのインスタンス作成
    Set mclsKeyBoard = New cKeyBoard
    mclsKeyBoard.SetDevice picKey
    
    '入力チェッククラスのインスタンス作成
    Set mclsTimer = New cInputTimer
    mclsTimer.SetTimerControl Me.Timer1
    
    '単語取得クラスのインスタンス作成
    Set mclsTanngo = New cTanngo
    mclsTanngo.ReadTangoText
    
    'WAVE再生クラスのインスタンス作成
    Set mclsPlayWave = New cWave

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'キーボード描画クラスのインスタンス破棄
    Set mclsKeyBoard = Nothing
    
    '入力チェッククラスのインスタンス破棄
    Set mclsTimer = Nothing
    
    '単語取得クラスのインスタンス破棄
    Set mclsTanngo = Nothing

    'WAVE再生クラスのインスタンス作成
    Set mclsPlayWave = Nothing

End Sub

Private Sub DispMessage(ByVal idx As Long)
    Select Case idx
        Case INP_OK
            '入力成功！
            lblTanngoKanji.Caption = "やった！"
            lblTanngoYomi.Caption = ""
            lblInput.Caption = ""
            mclsPlayWave.PlayWave AddDirSep(App.Path) & "External.wav"
            Set imgMain.Picture = LoadResPicture(RES_BMP_OK, vbResBitmap)
        
        Case INP_NG
            '入力失敗！
            lblTanngoKanji.Caption = "失敗！"
            lblTanngoYomi.Caption = ""
            lblInput.Caption = ""
            mclsPlayWave.PlayWave AddDirSep(App.Path) & "ChatAction.wav"
            Set imgMain.Picture = LoadResPicture(RES_BMP_NG, vbResBitmap)
        
        Case INP_START
            '練習開始
            lblTanngoKanji.Caption = "舞々開始！"
            lblTanngoYomi.Caption = ""
            lblInput.Caption = ""
            mclsPlayWave.PlayWave AddDirSep(App.Path) & "Startup.wav"
            Set imgMain.Picture = LoadResPicture(RES_BMP_START, vbResBitmap)
            
            'ウエイトをかける
            Dim Start As Long
            Dim Wait As Long
            Wait = 3
            Start = Timer
            Do While Timer < Start + Wait
                DoEvents
            Loop
        
        Case INP_END
            '練習開始
            lblTanngoKanji.Caption = "しゅーりょうー"
            lblTanngoYomi.Caption = ""
            lblInput.Caption = ""
            mclsPlayWave.PlayWave AddDirSep(App.Path) & "Homepage.wav"
            Set imgMain.Picture = LoadResPicture(RES_BMP_END, vbResBitmap)
    
    End Select

End Sub


Private Sub mclsTimer_InputEnd(ByVal bOK As Boolean)
    
    If bOK Then
        Call DispMessage(INP_OK)
    Else
        Call DispMessage(INP_NG)
    End If
    
    mNowGekiuti = False
    Call mclsKeyBoard.ShowKeyboard
    
    'ウエイトをかける
    Dim Start As Long
    Dim Wait As Long
    Wait = 2
    Start = Timer
    Do While Timer < Start + Wait
        DoEvents
    Loop
    
    '次の単語表示
    Call NextTanngo

End Sub

'入力ミス
Private Sub mclsTimer_InputMiss()
    Set imgMain.Picture = LoadResPicture(RES_BMP_NG, vbResBitmap)
End Sub

'正常入力
Private Sub mclsTimer_InputUpdate(ByVal strTanngo As String)
    Set imgMain.Picture = LoadResPicture(RES_BMP_NOW, vbResBitmap)
    lblInput.Caption = strTanngo
End Sub
