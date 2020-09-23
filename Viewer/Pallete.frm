VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form4"
   ScaleHeight     =   3465
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BWidth, BHeight, CIndex, bxStart, byStart, clr As Long
Dim ih, iv, BSpace As Integer
Dim Color As Long
Sub Soso()
BSpace = 2
BWidth = Int((Picture1.ScaleWidth - 32) / 16)
BHeight = Int((Picture1.ScaleHeight - 32) / 16)

For ih = 0 To 15
    For iv = 0 To 15
     CIndex = ih * 16 + iv
     bxStart = ih * (BWidth + BSpace)
     byStart = iv * (BHeight + BSpace)
     clr = Color + CIndex
     Picture1.Line (bxStart, byStart)-Step(BWidth, BHeight), clr, BF
    Next
Next
End Sub
Private Sub Form_Click()
CommonDialog1.Action = 3
Color = CommonDialog1.Color
Call Soso
End Sub

