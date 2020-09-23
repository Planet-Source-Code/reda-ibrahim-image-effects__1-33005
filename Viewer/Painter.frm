VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm Painter 
   BackColor       =   &H8000000C&
   Caption         =   "Painter"
   ClientHeight    =   6345
   ClientLeft      =   930
   ClientTop       =   630
   ClientWidth     =   7530
   Icon            =   "Painter.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList Layeri 
      Left            =   1890
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":0894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   5505
      Left            =   4680
      ScaleHeight     =   5475
      ScaleWidth      =   2820
      TabIndex        =   3
      Top             =   465
      Width           =   2850
      Begin TabDlg.SSTab SSTab1 
         Height          =   7455
         Left            =   30
         TabIndex        =   4
         Top             =   -60
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   13150
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Layers"
         TabPicture(0)   =   "Painter.frx":0CE6
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Toolbar3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "TreeView1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Picture2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Colors"
         TabPicture(1)   =   "Painter.frx":0D02
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture4"
         Tab(1).Control(1)=   "Picture3"
         Tab(1).Control(2)=   "Command1"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Others"
         TabPicture(2)   =   "Painter.frx":0D1E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.CommandButton Command1 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -72690
            TabIndex        =   14
            Top             =   4770
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            Height          =   1485
            Left            =   -74940
            Picture         =   "Painter.frx":0D3A
            ScaleHeight     =   1425
            ScaleWidth      =   2565
            TabIndex        =   13
            Top             =   3210
            Width           =   2625
         End
         Begin VB.PictureBox Picture4 
            Height          =   1995
            Left            =   -74940
            MousePointer    =   2  'Cross
            Picture         =   "Painter.frx":377D
            ScaleHeight     =   1935
            ScaleWidth      =   2595
            TabIndex        =   12
            Top             =   450
            Width           =   2655
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   465
            Left            =   600
            ScaleHeight     =   465
            ScaleWidth      =   2085
            TabIndex        =   7
            Top             =   390
            Width           =   2085
            Begin VB.TextBox Text1 
               Height          =   315
               Left            =   1380
               TabIndex        =   9
               Text            =   "100"
               Top             =   30
               Width           =   405
            End
            Begin ComctlLib.Slider Slider1 
               Height          =   375
               Left            =   1800
               TabIndex        =   8
               Top             =   30
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   661
               _Version        =   327682
               Orientation     =   1
               LargeChange     =   1
               Max             =   100
               SelStart        =   100
               TickStyle       =   3
               Value           =   100
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   3375
            Left            =   90
            TabIndex        =   6
            Top             =   900
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   5953
            _Version        =   393217
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   1
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   420
            Left            =   90
            TabIndex        =   5
            Top             =   4350
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   741
            ButtonWidth     =   609
            ButtonHeight    =   582
            Appearance      =   1
            ImageList       =   "Layeri"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   14
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "New Layer"
                  Style           =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Add"
                  Object.ToolTipText     =   "New Layer"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Delete Layer"
                  ImageIndex      =   2
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":4C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":530E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":5762
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":5BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":600A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":6686
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":6D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":737E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":79FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":8076
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":86F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":8D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":93EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":9A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":A0E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":A75E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":ABB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":B006
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":B462
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":B8B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":BD0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":C15E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":C5B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":CA06
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":CE5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5970
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5080
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5080
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1560
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":D2AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":D702
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":DB56
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":DFB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":E406
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":E51E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":E636
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":E74E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":E866
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":E97A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":EA8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":EBA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Painter.frx":ECB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   820
      ButtonWidth     =   820
      ButtonHeight    =   767
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New File"
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Image"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Save To File"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Prev"
            Object.ToolTipText     =   "Preview"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Print"
            Object.ToolTipText     =   "Print File"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   5505
      Left            =   0
      TabIndex        =   2
      Top             =   465
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   9710
      ButtonWidth     =   820
      ButtonHeight    =   767
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pointer"
            Object.ToolTipText     =   "Pointer"
            ImageIndex      =   4
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Line"
            Object.ToolTipText     =   "Line"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Ellipse"
            Object.ToolTipText     =   "Ellipse"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Arc"
            Object.ToolTipText     =   "Arc"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Free"
            Object.ToolTipText     =   "Free Line"
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Rect"
            Object.ToolTipText     =   "Rectangle"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Brush"
            Object.ToolTipText     =   "Brush"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Text"
            Object.ToolTipText     =   "Text"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Fill"
            Object.ToolTipText     =   "Fill Region"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Style"
            Object.ToolTipText     =   "Line style"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox FCol 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   30
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   11
         Top             =   4590
         Width           =   285
      End
      Begin VB.PictureBox BCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         ScaleHeight     =   285
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   4530
         Width           =   285
      End
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuN 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuc 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSa 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSv 
         Caption         =   "Save &As"
      End
      Begin VB.Menu dsas 
         Caption         =   "-"
      End
      Begin VB.Menu mnuP 
         Caption         =   "&Print Setup"
      End
      Begin VB.Menu mnuPr 
         Caption         =   "P&rint "
         Shortcut        =   ^P
      End
      Begin VB.Menu mnudd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEx 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEd 
      Caption         =   "&Edit"
      Begin VB.Menu mnuU 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu gge 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCop 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPs 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuD 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuEf 
      Caption         =   "E&ffect"
      Begin VB.Menu mnuFli 
         Caption         =   "&Flip"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnumir 
         Caption         =   "&mirror"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuRot 
         Caption         =   "&Rotate"
         Shortcut        =   ^R
      End
      Begin VB.Menu spcr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInv 
         Caption         =   "&Inverse"
      End
      Begin VB.Menu mnuGlass 
         Caption         =   "&Glass"
      End
      Begin VB.Menu gg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSm 
         Caption         =   "&Smooth"
      End
      Begin VB.Menu mnuSh 
         Caption         =   "S&harpen.."
      End
      Begin VB.Menu mnuEm 
         Caption         =   "&Emboss"
      End
      Begin VB.Menu mnuDif 
         Caption         =   "&Diffuse"
      End
      Begin VB.Menu mnuSol 
         Caption         =   "S&olarize"
      End
      Begin VB.Menu spcc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCu 
         Caption         =   "&Custom Filter.."
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Window"
      Begin VB.Menu mnuTH 
         Caption         =   "&Tile Horizontally"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuVer 
         Caption         =   "Tile &Vertically"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuCas 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuArr 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuCon 
         Caption         =   "&Contents.."
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "&Index"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Painter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub FileOpen()
    Dim i, j As Integer
    Dim red, green, blue As Integer
    Dim pixel As Long
    Me.Enabled = False
    X = Me.ActiveForm.Picture1.ScaleWidth
    Y = Me.ActiveForm.Picture1.ScaleHeight
    If X > 800 Or Y > 800 Then
    MsgBox "Image too large to process."
    X = 0: Y = 0
    Exit Sub
    End If
    hBMP = CreateCompatibleBitmap(Me.ActiveForm.Picture1.hdc, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Me.ActiveForm.Picture1.hdc)
    SelectObject hDestDC, hBMP
    Form6.Show
    Form6.Refresh
    Form6.Caption = "Read File Pixels..."
    For i = 0 To Y - 1
        For j = 0 To X - 1
            pixel = GetPixel(Me.ActiveForm.Picture1.hdc, j, i)
            red = pixel& Mod 256
            green = ((pixel And &HFF00) / 256&) Mod 256&
            blue = (pixel And &HFF0000) / 65536
            ImagePixels(0, i, j) = red
            ImagePixels(1, i, j) = green
            ImagePixels(2, i, j) = blue
        Next
        Form6.ProgressBar1.Value = i * 100 / (Y - 1)
        DoEvents
    Next
    Me.Enabled = True
    Unload Form6
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
End Sub

Private Sub BCol_Click()
Dim C As Long
C = FCol.BackColor
FCol.Appearance = 0
FCol.BackColor = C
C = BCol.BackColor
BCol.Appearance = 1
BCol.BackColor = C
CColor = "Back"
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowColor
If CColor = "Back" Then
BCol.BackColor = CommonDialog1.Color
Else
FCol.BackColor = CommonDialog1.Color
End If
End Sub

Private Sub FCol_Click()
Dim C As Long
C = BCol.BackColor
BCol.Appearance = 0
BCol.BackColor = C
C = FCol.BackColor
FCol.Appearance = 1
FCol.BackColor = C
CColor = "Fore"
End Sub

Private Sub MDIForm_Deactivate()
'Form5.ZOrder
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuAbout_Click()
About.Show
End Sub

Private Sub mnuArr_Click()
Painter.Arrange vbArrangeIcons
End Sub

Private Sub mnuCas_Click()
Painter.Arrange vbCascade
End Sub

Private Sub mnuCu_Click()
Form8.Show
End Sub

Private Sub mnuDif_Click()
Dim i, j As Integer
Dim Rx, Ry As Integer
Dim red, blue, green As Integer
    Form6.Show
    Form6.Refresh
    Form6.Caption = "Diffuse the Image..."
    hBMP = CreateCompatibleBitmap(Me.ActiveForm.Picture1.hdc, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Me.ActiveForm.Picture1.hdc)
    SelectObject hDestDC, hBMP

    For i = 2 To Y - 3
        For j = 2 To X - 3
            Rx = Rnd() * 4 - 2
            Ry = Rnd() * 4 - 2
            red = ImagePixels(0, i + Rx, j + Ry)
            green = ImagePixels(1, i + Rx, j + Ry)
            blue = ImagePixels(2, i + Rx, j + Ry)
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        DoEvents
        Form6.ProgressBar1.Value = i * 100 / (Y - 1)
    Next
    BitBlt Me.ActiveForm.Picture1.hdc, 1, 1, Me.ActiveForm.Picture1.ScaleWidth - 2, Me.ActiveForm.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Me.ActiveForm.Picture1.Refresh
    Me.ActiveForm.Picture1.Picture = Me.ActiveForm.Picture1.Image
    Unload Form6
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
End Sub

Private Sub mnuEm_Click()
Dim i, j As Integer
Const Dx As Integer = 1
Const Dy As Integer = 1
Dim red, blue, green As Integer
    Form6.Show
    Form6.Refresh
    Form6.Caption = "Emboss the Image..."
    hBMP = CreateCompatibleBitmap(Me.ActiveForm.Picture1.hdc, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Me.ActiveForm.Picture1.hdc)
    SelectObject hDestDC, hBMP

    For i = 1 To Y - 2
        For j = 1 To X - 2
            red = Abs(ImagePixels(0, i, j) - ImagePixels(0, i + Dx, j + Dy) + 128)
            green = Abs(ImagePixels(1, i, j) - ImagePixels(1, i + Dx, j + Dy) + 128)
            blue = Abs(ImagePixels(2, i, j) - ImagePixels(2, i + Dx, j + Dy) + 128)
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        Form6.ProgressBar1.Value = i * 100 / (Y - 1)
    Next
    BitBlt Me.ActiveForm.Picture1.hdc, 1, 1, Me.ActiveForm.Picture1.ScaleWidth - 2, Me.ActiveForm.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Me.ActiveForm.Picture1.Refresh
    Me.ActiveForm.Picture1.Picture = Me.ActiveForm.Picture1.Image
    Unload Form6
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
End Sub

Private Sub mnuEx_Click()
End
End Sub

Private Sub mnuFli_Click()
On Error Resume Next
Me.ActiveForm.Picture1.PaintPicture Me.ActiveForm.Picture1.Picture, 0, 0, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight, -Me.ActiveForm.Picture1.ScaleWidth, -Me.ActiveForm.Picture1.ScaleHeight, vbSrcCopy
Me.ActiveForm.Picture1.Refresh
Set Pic = Me.ActiveForm.Picture1
Me.ActiveForm.Picture1.Picture = Me.ActiveForm.Picture1.Image
End Sub

Private Sub mnuInv_Click()
Me.ActiveForm.Picture1.PaintPicture Me.ActiveForm.Picture1.Picture, 0, 0, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight, 0, 0, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight, vbNotSrcCopy
Me.ActiveForm.Picture1.Refresh
Me.ActiveForm.Picture1.Picture = Me.ActiveForm.Picture1.Image
End Sub

Private Sub mnumir_Click()
On Error Resume Next
Me.ActiveForm.Picture1.PaintPicture Me.ActiveForm.Picture1.Picture, 0, 0, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight, Me.ActiveForm.Picture1.ScaleWidth, 0, -Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight, vbSrcCopy
Me.ActiveForm.Picture1.Refresh
Me.ActiveForm.Picture1.Picture = Me.ActiveForm.Picture1.Image
End Sub

Private Sub mnuN_Click()
Dim xx
Form5.Show
End Sub


Private Sub mnuOpen_Click()
CommonDialog1.Filter = "Image Files |*.bmp;*.jpg;*.gif;*.tif;*.dib;*.wmf;*.gif|All Files|*.*"
CommonDialog1.Action = 1
Set Form = New Form3
Form.Picture1.Picture = LoadPicture(CommonDialog1.FileName)
Form.Height = Form.Picture1.Height
Form.Width = Form.Picture1.Width
Form.Caption = CommonDialog1.FileTitle
Form.Show
Set Pic = Form.Picture1
Pic.Picture = Form.Picture1.Picture
Toolbar2.Buttons(3).Enabled = True
Call FileOpen
End Sub

Private Sub mnuRot_Click()
Form7.Show
End Sub

Private Sub mnuSa_Click()
Dim T1, e
Form6.Show
For T1 = Form6.ProgressBar1.Min To Form6.ProgressBar1.Max
e = Form6.ProgressBar1.Value
On Error Resume Next
Form6.ProgressBar1.Value = e + 10
DoEvents
Next T1
Unload Form6
End Sub

Private Sub mnuSh_Click()
Dim i, j As Integer
Const Dx As Integer = 1
Const Dy As Integer = 1
Dim red, blue, green As Integer
    Form6.Show
    Form6.Refresh
    Form6.Caption = "Sharpen the Image..."
    hBMP = CreateCompatibleBitmap(Me.ActiveForm.Picture1.hdc, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Me.ActiveForm.Picture1.hdc)
    SelectObject hDestDC, hBMP
    For i = 1 To Y - 2
        For j = 1 To X - 2
            red = ImagePixels(0, i, j) + 0.5 * (ImagePixels(0, i, j) - ImagePixels(0, i - Dx, j - Dy))
            red = ImagePixels(1, i, j) + 0.5 * (ImagePixels(1, i, j) - ImagePixels(1, i - Dx, j - Dy))
            red = ImagePixels(2, i, j) + 0.5 * (ImagePixels(2, i, j) - ImagePixels(2, i - Dx, j - Dy))
            If red > 255 Then red = 255
            If red < 0 Then red = 0
            If green > 255 Then green = 255
            If green < 0 Then green = 0
            If blue > 255 Then blue = 255
            If blue < 0 Then blue = 0
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        Form6.ProgressBar1.Value = i * 100 / (Y - 1)
        
    Next
    BitBlt Me.ActiveForm.Picture1.hdc, 1, 1, Me.ActiveForm.Picture1.ScaleWidth - 2, Me.ActiveForm.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Me.ActiveForm.Picture1.Refresh
    Me.ActiveForm.Picture1.Picture = Me.ActiveForm.Picture1.Image
    Unload Form6
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
End Sub

Private Sub mnuSm_Click()
Dim i, j As Integer
Const Dx As Integer = 1
Const Dy As Integer = 1
Dim red, blue, green As Integer
    Form6.Show
    Form6.Refresh
    Form6.Caption = "Smooth the Image..."
    hBMP = CreateCompatibleBitmap(Me.ActiveForm.Picture1.hdc, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Me.ActiveForm.Picture1.hdc)
    SelectObject hDestDC, hBMP

    For i = 1 To Y - 2
        For j = 1 To X - 2
            red = ImagePixels(0, i - 1, j - 1) + ImagePixels(0, i - 1, j) + ImagePixels(0, i - 1, j + 1) + ImagePixels(0, i, j - 1) + ImagePixels(0, i, j) + ImagePixels(0, i, j + 1) + ImagePixels(0, i + 1, j - 1) + ImagePixels(0, i + 1, j) + ImagePixels(0, i + 1, j + 1)
            green = ImagePixels(1, i - 1, j - 1) + ImagePixels(1, i - 1, j) + ImagePixels(1, i - 1, j + 1) + ImagePixels(1, i, j - 1) + ImagePixels(1, i, j) + ImagePixels(1, i, j + 1) + ImagePixels(1, i + 1, j - 1) + ImagePixels(1, i + 1, j) + ImagePixels(1, i + 1, j + 1)
            blue = ImagePixels(2, i - 1, j - 1) + ImagePixels(2, i - 1, j) + ImagePixels(2, i - 1, j + 1) + ImagePixels(2, i, j - 1) + ImagePixels(2, i, j) + ImagePixels(2, i, j + 1) + ImagePixels(2, i + 1, j - 1) + ImagePixels(2, i + 1, j) + ImagePixels(1, i + 1, j + 1)
            SetPixelV hDestDC, j, i, RGB(red / 9, green / 9, blue / 9)
        Next
        Form6.ProgressBar1.Value = i * 100 / (Y - 1)
        DoEvents
    Next
    BitBlt Me.ActiveForm.Picture1.hdc, 1, 1, Me.ActiveForm.Picture1.ScaleWidth - 2, Me.ActiveForm.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Me.ActiveForm.Picture1.Refresh
    Me.ActiveForm.Picture1.Picture = Me.ActiveForm.Picture1.Image
    Unload Form6
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
   
End Sub

Private Sub mnuSol_Click()
Dim i, j As Integer
Dim red, blue, green As Integer
    Form6.Show
    Form6.Refresh
    Form6.Caption = "Solarize the Image..."
    hBMP = CreateCompatibleBitmap(Me.ActiveForm.Picture1.hdc, Me.ActiveForm.Picture1.ScaleWidth, Me.ActiveForm.Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Me.ActiveForm.Picture1.hdc)
    SelectObject hDestDC, hBMP

    For i = 1 To Y - 2
        For j = 1 To X - 2
            red = ImagePixels(0, i, j)
            green = ImagePixels(1, i, j)
            blue = ImagePixels(2, i, j)
            If ((red < 128) Or (red > 255)) Then red = 255 - red
            If ((green < 128) Or (green > 255)) Then green = 255 - green
            If ((blue < 128) Or (blue > 255)) Then blue = 255 - blue
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        DoEvents
        Form6.ProgressBar1.Value = i * 100 / (Y - 1)
    Next
    BitBlt Me.ActiveForm.Picture1.hdc, 1, 1, Me.ActiveForm.Picture1.ScaleWidth - 2, Me.ActiveForm.Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Me.ActiveForm.Picture1.Refresh
    Me.ActiveForm.Picture1.Picture = Me.ActiveForm.Picture1.Image

    Unload Form6
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMP)
End Sub

Private Sub mnuTH_Click()
Painter.Arrange 1
End Sub

Private Sub mnuVer_Click()
Painter.Arrange vbTileVertical
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CColor = "Back" Then
CFlag = "0"
BColor = Picture3.Point(X, Y)
BCol.BackColor = BColor
Else
CFlag = "1"
FColor = Picture3.Point(X, Y)
FCol.BackColor = FColor
End If
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CFlag = ""
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CColor = "Back" Then
CFlag = "0"
BColor = Picture4.Point(X, Y)
BCol.BackColor = BColor
Else
CFlag = "1"
FColor = Picture4.Point(X, Y)
FCol.BackColor = FColor
End If
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If CFlag = "0" Then
BColor = Picture4.Point(X, Y)
BCol.BackColor = BColor
ElseIf CFlag = "1" Then
FColor = Picture4.Point(X, Y)
FCol.BackColor = FColor
End If
End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CFlag = ""
End Sub

Private Sub Slider1_Change()
Text1.Text = Slider1.Value

End Sub

Private Sub Slider1_Scroll()
Text1.Text = Slider1.Value

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Line"
Form.Picture1.MousePointer = 2
Flag = "Line"
Case "Rect":
Flag = "Rect"
Case "Pointer"
On Error Resume Next
Form.Picture1.MousePointer = 0
Flag = "Pointer"
Case "Brush":
Form.Picture1.MouseIcon = Form.Picture2.Picture
Flag = "Brush"
Case "Ellipse":
Flag = "Circle"
Case "Free":
Flag = "Free"
Case "Text":
Flag = "Text"
Case Default:
Flag = ""
Form.Picture1.MousePointer = 0
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "New":
Call mnuN_Click
Case "Open"
Call mnuOpen_Click
Case "Help":
    About.Show
End Select
End Sub

Private Sub Toolbar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 1:
StatusBar1.Panels(1).Text = Toolbar2.Buttons(1).ToolTipText
Case 2:
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(1).Text = Toolbar2.Buttons(2).ToolTipText
Case 3:
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(1).Text = Toolbar2.Buttons(3).ToolTipText
End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Xr
Select Case Button.Key
Case "Add":
Dim D As String
D = "L" & Str(LKey + 1)

Set Xr = TreeView1.Nodes.Add(, , D, "Layer" + Str(LKey), 7)

Case "Remove":
Call MenuF.mnuDeleteL_Click
End Select
LKey = LKey + 1
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu MenuF.mnuLayers

End Sub
