VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custom Filter"
   ClientHeight    =   3180
   ClientLeft      =   2670
   ClientTop       =   1785
   ClientWidth     =   4305
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Now"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Text            =   "1"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Text            =   "0"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Text            =   "0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Text            =   "0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Text            =   "0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   2640
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      Begin VB.OptionButton Option2 
         Caption         =   "5 X 5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 X 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Filter Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Divide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Bias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   2280
      Width           =   375
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CustomFilter(3, 3) As Long
Dim FilterNorm, FilterBias As Integer
Private Sub Command1_Click()
Dim RedSum, GreenSum, BlueSum As Integer
Dim red, green, blue As Integer
Dim fi, fj As Integer
Dim i, j As Integer
Dim Offset As Integer
On Error Resume Next
CustomFilter(1, 1) = Val(Text1.Text)
CustomFilter(2, 1) = Val(Text2.Text)
CustomFilter(3, 1) = Val(Text3.Text)
CustomFilter(1, 2) = Val(Text4.Text)
CustomFilter(2, 2) = Val(Text5.Text)
CustomFilter(3, 2) = Val(Text6.Text)
CustomFilter(1, 3) = Val(Text7.Text)
CustomFilter(2, 3) = Val(Text8.Text)
CustomFilter(3, 3) = Val(Text9.Text)
Painter.Enabled = True
Unload Me
Form6.Show
    Form6.Caption = "Smooth the Image..."
    hBMP = CreateCompatibleBitmap(Painter.ActiveForm.Picture1.hdc, Painter.ActiveForm.Picture1.ScaleWidth, Painter.ActiveForm.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Painter.ActiveForm.Picture1.hdc)
    SelectObject hDestDC, hBMP
X = Painter.ActiveForm.Picture1.ScaleWidth
Y = Painter.ActiveForm.Picture1.ScaleHeight
If Me.Option1.Value = True Then
    Offset = 1
Else
    Offset = 2
End If
For i = Offset To Y - Offset - 1
    For j = Offset To X - Offset - 1
        RedSum = 0: GreenSum = 0: BlueSum = 0
            For fi = -Offset To Offset
                For fj = -Offset To Offset
                    RedSum = RedSum + ImagePixels(0, i + fi, j + fj) * CustomFilter(fi + 2, fj + 2)
                    GreenSum = GreenSum + ImagePixels(1, i + fi, j + fj) * CustomFilter(fi + 2, fj + 2)
                    BlueSum = BlueSum + ImagePixels(2, i + fi, j + fj) * CustomFilter(fi + 2, fj + 2)
                Next
            Next
            red = Abs(RedSum / FilterNorm + FilterBias)
            green = Abs(GreenSum / FilterNorm + FilterBias)
            blue = Abs(BlueSum / FilterNorm + FilterBias)
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
    Next
    DoEvents
    Form6.ProgressBar1.Value = i * 100 / (Y - 1)
    'Painter.ActiveForm.Picture1.Refresh
Next
Unload Form6
    BitBlt Painter.ActiveForm.Picture1.hdc, 1, 1, Painter.ActiveForm.Picture1.ScaleWidth - 2, Painter.ActiveForm.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Painter.ActiveForm.Picture1.Refresh
    Painter.ActiveForm.Picture1.Picture = Painter.ActiveForm.Picture1.Image
End Sub

Private Sub Command2_Click()
Painter.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
'Painter.Enabled = False
End Sub

Private Sub Text1_Change()
CustomFilter(1, 1) = Val(Text1.Text)
End Sub

Private Sub Text10_Change()
FilterBias = Val(Text10.Text)
End Sub

Private Sub Text11_Change()
FilterNorm = Val(Text11.Text)
End Sub

Private Sub Text2_Change()
CustomFilter(2, 1) = Val(Text2.Text)
End Sub

Private Sub Text3_Change()
CustomFilter(3, 1) = Val(Text3.Text)
End Sub

Private Sub Text4_Change()
CustomFilter(1, 2) = Val(Text4.Text)
End Sub

Private Sub Text5_Change()
CustomFilter(2, 2) = Val(Text5.Text)
End Sub

Private Sub Text6_Change()
CustomFilter(3, 2) = Val(Text6.Text)
End Sub

Private Sub Text7_Change()
CustomFilter(1, 3) = Val(Text7.Text)
End Sub

Private Sub Text8_Change()
CustomFilter(2, 3) = Val(Text8.Text)
End Sub

Private Sub Text9_Change()
CustomFilter(3, 3) = Val(Text9.Text)
End Sub
