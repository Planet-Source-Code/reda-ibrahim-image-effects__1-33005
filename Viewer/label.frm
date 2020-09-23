VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert label"
   ClientHeight    =   4680
   ClientLeft      =   2025
   ClientTop       =   1200
   ClientWidth     =   5460
   Icon            =   "label.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5460
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":044A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":055E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":0672
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":0786
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":095E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":0A4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":0B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":0C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "label.frx":0D0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "color"
            Object.ToolTipText     =   "Color"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "black"
                  Text            =   "Black"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "white"
                  Text            =   "White"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "blue"
                  Text            =   "Blue"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "yellow"
                  Text            =   "Yellow"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "green"
                  Text            =   "Green"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "gray"
                  Text            =   "Gray"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "red"
                  Text            =   "Red"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "label.frx":0DFA
      Left            =   2640
      List            =   "label.frx":0E2E
      TabIndex        =   2
      Text            =   "28"
      Top             =   480
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1230
      Width           =   4815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Combo1_Click()
Text1.FontName = Combo1.Text
Lab.FontName = Combo1.Text
End Sub

Private Sub Combo2_Change()
On Error Resume Next
Text1.FontSize = Val(Combo2.Text)
Lab.FontSize = Val(Combo2.Text)
End Sub

Private Sub Combo2_Click()
Text1.FontSize = Val(Combo2.Text)
Lab.FontSize = Val(Combo2.Text)
End Sub

Private Sub Command1_Click()
Painter.Enabled = True
Lab.Visible = True
Unload Me
End Sub

Private Sub Command2_Click()
Painter.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Text1.FontSize = Val(Combo2.Text)
Lab.FontSize = Val(Combo2.Text)
Painter.Enabled = False
Lab.Visible = True
For i = 0 To Screen.FontCount - 1
Combo1.AddItem Screen.Fonts(i)
Next i
Combo1.Text = "Times New Roman"
Text1.FontName = Combo1.Text
Lab.FontName = Combo1.Text
End Sub

Private Sub Text1_Change()
Lab.Caption = Text1.Text
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "bold":
    Text1.FontBold = Not Text1.FontBold
    Lab.FontBold = Not Lab.FontBold
Case "italic":
    Text1.FontItalic = Not Text1.FontItalic
    Lab.FontItalic = Not Lab.FontItalic
Case "underline":
    Text1.FontUnderline = Not Text1.FontUnderline
    Lab.FontUnderline = Not Lab.FontUnderline
Case "color":
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "yellow":
            Toolbar1.Buttons(5).Image = 10
            Text1.ForeColor = vbYellow
            Lab.ForeColor = vbYellow
        Case "blue":
            Toolbar1.Buttons(5).Image = 5
            Text1.ForeColor = vbBlue
            Lab.ForeColor = vbBlue
        Case "gray":
            Toolbar1.Buttons(5).Image = 6
            Text1.ForeColor = vbGrayText
            Lab.ForeColor = vbGrayText
        Case "red":
            Toolbar1.Buttons(5).Image = 4
            Text1.ForeColor = vbRed
            Lab.ForeColor = vbRed
        Case "white":
            Toolbar1.Buttons(5).Image = 9
            Text1.ForeColor = vbWhite
            Lab.ForeColor = vbWhite
        Case "black":
            Toolbar1.Buttons(5).Image = 8
            Text1.ForeColor = vbBlack
            Lab.ForeColor = vbBlack
        Case "green":
            Toolbar1.Buttons(5).Image = 7
            Text1.ForeColor = vbGreen
            Lab.ForeColor = vbGreen
    End Select

End Sub
