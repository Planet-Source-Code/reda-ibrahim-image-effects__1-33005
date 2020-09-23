VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Image"
   ClientHeight    =   2940
   ClientLeft      =   2955
   ClientTop       =   1485
   ClientWidth     =   3780
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " New Image Properities"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "new.frx":0000
         Left            =   1680
         List            =   "new.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Text            =   " "
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Text            =   " "
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Background color"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "x"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "&Height"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "&Width"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Painter.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()
Painter.Enabled = True
Set Form = New Form3
On Error Resume Next
Form.Height = Val(Text1.Text) * 10
Form.Width = Val(Text2.Text) * 10
Form.Picture1.Height = Val(Text1.Text) * 10
Form.Picture1.Width = Val(Text2.Text) * 10
Form.Picture1.BackColor = vbWhite
Form.Picture1.ForeColor = vbBlack
Form.Show
Set xx = Painter.TreeView1.Nodes.Add(, , , "Layer", 7)
Unload Me
End Sub

Private Sub Form_Load()
Painter.Enabled = False
Text1.Text = "500"
Text2.Text = "500"
Combo1.Text = "Background Color"
Me.ZOrder 0
End Sub


