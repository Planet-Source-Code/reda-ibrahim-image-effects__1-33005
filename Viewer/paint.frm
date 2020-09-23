VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Image"
   ClientHeight    =   3255
   ClientLeft      =   2685
   ClientTop       =   2445
   ClientWidth     =   4695
   DrawStyle       =   4  'Dash-Dot-Dot
   FillColor       =   &H00400000&
   FillStyle       =   6  'Cross
   Icon            =   "paint.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox Picture2 
         Height          =   735
         Left            =   1080
         Picture         =   "paint.frx":0442
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   1
         Top             =   3240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Shape Shape1 
         Height          =   1335
         Left            =   1200
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   16
         X2              =   120
         Y1              =   152
         Y2              =   192
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XStart, YStart As Long
Dim XDash, YDash As Long
Dim Xold, Yold As Long
Dim XS, YS As Long
Dim flg As Boolean
Dim flg2 As Boolean
Dim SaveFlg As Boolean
Dim Str As Integer

Private Sub Form_Load()
'PictWidth = Me.ScaleX(Me.Picture1.Width, vbHimetric, vbTwips)
'PictHeight = Me.ScaleY(Me.Picture1.Height, vbHimetric, vbTwips)
'Me.Move 0, 0, PictWidth, PictHeight
Painter.Toolbar1.Buttons(2).Enabled = True
Painter.Toolbar1.Buttons(3).Enabled = True
Painter.Toolbar1.Buttons(4).Enabled = True
Painter.Toolbar1.Buttons(5).Enabled = True
Painter.Toolbar1.Buttons(6).Enabled = True
Painter.Toolbar1.Buttons(7).Enabled = True
Painter.Toolbar1.Buttons(8).Enabled = True
Picture1.AutoRedraw = True
Picture1.ScaleMode = 3
flg2 = True
Xold = Yold = 0
Str = 1
TabK = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Msg
If SaveFlg Then
Msg = MsgBox(Me.Caption & " has been Changed Are You want Save Changes?", vbYesNoCancel + vbExclamation, "Save")
If Msg = vbYes Then
CommonDialog1.Filter = "Bitmap File |*.bmp": CommonDialog1.ShowSave
SavePicture Picture1.Image, CommonDialog1.FileName
ElseIf Msg = vbCancel Then Cancel = 1
End If
End If
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbTab) And TabK Then
Painter.Toolbar1.Visible = False
Painter.Picture1.Visible = False
Else
Painter.Toolbar1.Visible = True
Painter.Picture1.Visible = True
End If
TabK = Not TabK
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SaveFlg = True
XStart = X: YStart = Y
Select Case Flag:
    Case "Line":
        flg = True
    Case "Rect":
        flg = True
    Case "Brush":
        Xold = XStart: Yold = YStart
        flg = True
    Case "Circle":
        flg = True
    Case "Free":
       If flg2 Then XS = X: YS = Y: flg2 = False
        flg = True
        If Str <> 1 Then Picture1.Line (X, Y)-(Xold, Yold)
        Xold = X: Yold = Y
        Str = 2
        If Button = 2 Then Picture1.Line (XS, YS)-(X, Y): flg2 = True
    Case "Text":
        If coun > 0 Then Load Label1(coun)
        Label1(coun).Left = X
        Label1(coun).Top = Y
        Set Lab = Label1(coun)
        coun = coun + 1
        Form9.Show
End Select
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Picture1.Refresh
Select Case Flag:
    Case "Line":
        Line1.BorderColor = Painter.FCol.BackColor
        If flg Then Line1.Visible = False: Line1.Visible = True: Line1.X1 = XStart: Line1.Y1 = YStart: Line1.X2 = X: Line1.Y2 = Y
        
    Case "Brush":
        If flg Then Picture1.Line (Xold, Yold)-(X, Y): Xold = X: Yold = Y
    Case "Rect":
            If flg Then Shape1.Visible = True: Shape1.Left = XStart: Shape1.Top = YStart: Shape1.Width = Abs(X - XStart): Shape1.Height = Abs(Y - YStart)
    Case "Circle":
        If flg Then Picture1.Circle (XStart, YStart), Sqr(Abs((XStart - X) * (XStart - X)) + Abs((YStart - Y) * (YStart - Y)))
        Picture1.Refresh
    Case "Free":
       flg2 = False
        If flg Then Line1.Visible = True: Line1.X1 = XDash: Line1.X2 = X: Line1.Y1 = YDash: Line1.Y2 = Y ': Xold = X: Yold = Y: 'If Button = 1 Then Picture1.Line (XDash, YDash)-(X, Y)
End Select
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line1.Visible = False
Shape1.Visible = False
Picture1.ForeColor = Painter.FCol.BackColor
Select Case Flag:
    Case "Line":
        Picture1.DrawStyle = 0
        Picture1.DrawMode = 13
        Picture1.Line (XStart, YStart)-(X, Y)
        flg = False
    Case "Rect":
        Picture1.Line (XStart, YStart)-(X, Y), , B
        flg = False
    Case "Brush":
        flg = False
    Case "Circle":
        Picture1.Circle (XStart, YStart), Sqr(Abs((XStart - X) * (XStart - X)) + Abs((YStart - Y) * (YStart - Y)))
        flg = False
    Case "Free":
        'flg2 = False
        XDash = X: YDash = Y
        'If Button = 2 Then flg = False: flg2 = True: XDash = YDash = 0
        'Picture1.Line (Xold, Yold)-(X, Y)
        'If Not flg2 Then Picture1.Line (XDash, YDash)-(X, Y)
End Select
If Button = 2 Then Me.PopupMenu Painter.mnuEf, 2
End Sub
