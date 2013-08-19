VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTanngo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///    �P��f�[�^�擾�N���X

Private Type tTango
    lTime As Long
    sKanji As String
    sYomi As String
End Type
Private mTangoString() As tTango


Const TANNGO_FILE_NAME As String = "Tanngo.dat"

Public Sub ReadTangoText()
    
    '�P��i�[�t�@�C���p�X�̎擾
    Dim DatPath As String
    DatPath = AddDirSep(App.Path) & TANNGO_FILE_NAME
    
    Dim Fn      As Integer
    Dim lCount  As Long '�J�E���^
    
    Fn = FreeFile
    Open DatPath For Input As #Fn
    
    lCount = 0
    Do Until EOF(Fn)
        
        '�z��̍Ċm��
        ReDim Preserve mTangoString(lCount)
        
        '�P��擾
        Input #Fn, mTangoString(lCount).lTime, mTangoString(lCount).sKanji, mTangoString(lCount).sYomi
        Debug.Print mTangoString(lCount).lTime & ":" & mTangoString(lCount).sKanji & ":" & mTangoString(lCount).sYomi
        lCount = lCount + 1

    Loop
    
    Close #Fn


End Sub

Public Function GetTango(ByVal idx As Long, ByRef lTime As Long, _
                         ByRef sKanji As String, ByRef sYomi As String) As Boolean
    If idx < LBound(mTangoString()) Or UBound(mTangoString()) < idx Then
        GetTango = False
        lTime = 0
        sKanji = ""
        sYomi = ""
    
    Else
    
        GetTango = True
        lTime = mTangoString(idx).lTime
        sKanji = mTangoString(idx).sKanji
        sYomi = mTangoString(idx).sYomi


    End If

End Function
