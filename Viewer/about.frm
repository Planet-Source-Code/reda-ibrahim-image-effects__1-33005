VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   0  'None
   ClientHeight    =   4215
   ClientLeft      =   2535
   ClientTop       =   1155
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PowerWare"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   0
      Picture         =   "about.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Painter.Enabled = True
Painter.SetFocus
Unload Me
End Sub

Private Sub Form_Load()
Painter.Enabled = False
Me.ZOrder
End Sub

