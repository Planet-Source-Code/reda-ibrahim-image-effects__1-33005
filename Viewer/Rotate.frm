VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rotate"
   ClientHeight    =   1425
   ClientLeft      =   2085
   ClientTop       =   3030
   ClientWidth     =   5235
   Icon            =   "Rotate.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " Degrees"
      Height          =   1215
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton Option5 
         Caption         =   "&270"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "&180"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&90"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame D 
      Caption         =   " Direction"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton Option2 
         Caption         =   "&Left"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Right"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Painter.Enabled = True
If Option3.Value = True Then
    On Error Resume Next
    Painter.ActiveForm.Picture1.PaintPicture Painter.ActiveForm.Picture1.Picture, 0, 0, Painter.ActiveForm.Picture1.ScaleHeight, Painter.ActiveForm.Picture1.ScaleWidth, 0, Painter.ActiveForm.Picture1.ScaleHeight, Painter.ActiveForm.Picture1.ScaleHeight, -Painter.ActiveForm.Picture1.ScaleWidth, vbSrcCopy
    Painter.ActiveForm.Picture1.Refresh
    Painter.ActiveForm.Picture1.Picture = Painter.ActiveForm.Picture1.Image
End If
Unload Me
End Sub

Private Sub Command2_Click()
Painter.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Painter.Enabled = False
End Sub
