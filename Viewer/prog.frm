VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   390
   ClientLeft      =   2340
   ClientTop       =   3315
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Painter.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Painter.Enabled = True
End Sub

