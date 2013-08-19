VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   ClientHeight    =   4635
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4635
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
   Begin VB.Label lblVersion 
      BackStyle       =   0  '“§–¾
      Caption         =   "lblVersion"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblProductName 
      BackStyle       =   0  '“§–¾
      Caption         =   "lblProductName"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "ÊÞ°¼Þ®Ý " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
