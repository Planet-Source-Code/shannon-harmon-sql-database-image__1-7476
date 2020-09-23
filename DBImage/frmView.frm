VERSION 5.00
Begin VB.Form frmView 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "View Pic"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image imgPic 
      Height          =   825
      Left            =   1695
      Top             =   1215
      Width           =   990
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'*  By Shannon Harmon                        *
'*  Copyright 2000 - All rights reserved...  *
'*********************************************

Option Explicit

Private Sub Form_Load()
  imgPic.Visible = False
  imgPic.Picture = frmMain.imgGet.Picture
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
  imgPic.Left = Me.ScaleWidth / 2 - imgPic.Width / 2
  imgPic.Top = Me.ScaleHeight / 2 - imgPic.Height / 2
  imgPic.Refresh
  imgPic.Visible = True
End Sub

Private Sub imgPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Unload Me
End Sub
