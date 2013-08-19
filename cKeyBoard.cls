VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///    キーボード描画クラス

Private mpicKeyBoard As PictureBox
Private mKeyHeight   As Long 'キーの高さ
Private mKeyWide     As Long 'キーの幅

Private mKeyMargin(1 To 5)   As Long '任意の行の左端のキーの位置
Private mKeyArray(1 To 5) As String 'キートップの文字列格納用配列

Const KEY_DOWN_COLOR As Long = vbRed 'RGB(255, 0, 0)
Const KEY_UP_COLOR As Long = vbBlack 'RGB(0, 0, 0)



Private Sub Class_Initialize()
    mKeyArray(1) = " 1234567890-^\ "
    mKeyArray(2) = "QWERTYUIOP@["
    mKeyArray(3) = "ASDFGHJKL;:]"
    mKeyArray(4) = "ZXCVBNM,./_"
    
End Sub

'///    出力デバイスのセット
Public Sub SetDevice(ByRef picDest As PictureBox)

    Set mpicKeyBoard = picDest
    
    With mpicKeyBoard
    mKeyHeight = .ScaleHeight / 5
    mKeyWide = .ScaleWidth / 15
    .AutoRedraw = True
    End With

    mKeyMargin(1) = 0
    mKeyMargin(2) = mKeyWide * 1.5
    mKeyMargin(3) = mKeyWide * 1.8
    mKeyMargin(4) = mKeyWide * 2.3
    mKeyMargin(5) = mKeyWide * 5.5

    Call ShowKeyboard

End Sub

'///    キーボード表示
Public Sub ShowKeyboard()
    
    Dim i As Long
    
    mpicKeyBoard.ForeColor = KEY_UP_COLOR
    
    'キー１段目表示
    For i = 1 To 15
        Call ShowKey(i, 1)
    Next i
    
    'キー２段目表示
    For i = 1 To 12
        Call ShowKey(i, 2)
    Next i
    
    'キー３段目表示
    For i = 1 To 12
        Call ShowKey(i, 3)
    Next i
    
    'キー４段目表示
    For i = 1 To 11
        Call ShowKey(i, 4)
    Next i
    
    'キー５段目表示（スペースキーのみ）
    Call ShowKeySpace(1, 5)


End Sub

'///    キーの表示
Private Sub ShowKey(x As Long, y As Long, Optional FaceColor As Long)
    
    mpicKeyBoard.Line (mKeyMargin(y) + (x - 1) * mKeyWide, (y - 1) * mKeyHeight)-(mKeyMargin(y) + x * mKeyWide, y * mKeyHeight), RGB(0, 0, 0), B
    With mpicKeyBoard
    .CurrentX = mKeyMargin(y) + (2 * x - 1) * mKeyWide / 2 - .TextWidth(Mid$(mKeyArray(y), x, 1)) / 2
    .CurrentY = (2 * y - 1) * mKeyHeight / 2 - .TextHeight(Mid$(mKeyArray(y), x, 1)) / 2
    End With
    mpicKeyBoard.Print Mid$(mKeyArray(y), x, 1)

End Sub

'///    スペースキーの表示
Private Sub ShowKeySpace(x As Long, y As Long, Optional FaceColor As Long)
    
    mpicKeyBoard.Line (mKeyMargin(y) + (x - 1) * mKeyWide, (y - 1) * mKeyHeight)-(mKeyMargin(y) + x * mKeyWide * 4.5, y * mKeyHeight), RGB(0, 0, 0), B

End Sub


Public Function PushKeyBoardDown(iKey As Integer) As String
    Dim i       As Long
    Dim strPos  As Long
    Dim sKey    As String
    
    Select Case iKey
        Case 189:   sKey = "-"
        Case 222:   sKey = "^"
        Case 220:   sKey = "\"
        Case 9:     sKey = "TAB"
        Case 192:   sKey = "@"
        Case 219:   sKey = "["
        Case 187:   sKey = ";"
        Case 186:   sKey = ":"
        Case 221:   sKey = "]"
        Case 188:   sKey = ","
        Case 190:   sKey = "."
        Case 191:   sKey = "/"
        Case 32:    sKey = "SPC"
        Case Else: sKey = Chr$(iKey)
    End Select
    
    For i = LBound(mKeyArray) To UBound(mKeyArray)
        strPos = InStr(1, mKeyArray(i), sKey)
        If strPos <> 0 Then Exit For
    Next i

    '入力文字の取得
    If UBound(mKeyArray) >= i Then
        mpicKeyBoard.ForeColor = KEY_DOWN_COLOR
        Call ShowKey(strPos, i)
        
        PushKeyBoardDown = sKey
    
    Else
    
        PushKeyBoardDown = ""

    End If

End Function

Public Sub PushKeyBoardUp(iKey As Integer)
    Dim i As Long
    Dim strPos As Long
    Dim sKey    As String
    
    Select Case iKey
        Case 189:   sKey = "-"
        Case 222:   sKey = "^"
        Case 220:   sKey = "\"
        Case 9:     sKey = "TAB"
        Case 192:   sKey = "@"
        Case 219:   sKey = "["
        Case 187:   sKey = ";"
        Case 186:   sKey = ":"
        Case 221:   sKey = "]"
        Case 188:   sKey = ","
        Case 190:   sKey = "."
        Case 191:   sKey = "/"
        Case 32:    sKey = "SPC"
        Case Else: sKey = Chr$(iKey)
    End Select
    
    For i = LBound(mKeyArray) To UBound(mKeyArray)
        strPos = InStr(1, mKeyArray(i), sKey)
        If strPos <> 0 Then Exit For
    Next i

    If UBound(mKeyArray) >= i Then
        mpicKeyBoard.ForeColor = KEY_UP_COLOR
        Call ShowKey(strPos, i)

    End If

End Sub


