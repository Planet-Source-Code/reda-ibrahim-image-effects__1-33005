VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "prop.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "prop.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(1)=   "Image2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Preview"
      TabPicture(1)   =   "prop.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Sumuries"
      TabPicture(2)   =   "prop.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   3375
         ScaleWidth      =   4215
         TabIndex        =   5
         Top             =   480
         Width           =   4215
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   3375
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4215
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   3375
         ScaleWidth      =   4215
         TabIndex        =   3
         Top             =   480
         Width           =   4215
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Height          =   135
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Label7"
            Height          =   195
            Left            =   1080
            TabIndex        =   12
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label6 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Label5"
            Height          =   195
            Left            =   1080
            TabIndex        =   10
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label Label4 
            Caption         =   "Size:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Label3"
            Height          =   195
            Left            =   1080
            TabIndex        =   8
            Top             =   1200
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Location"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   375
            Left            =   1200
            TabIndex        =   6
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   -74880
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim St As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = Form1.CommonDialog1.FileTitle & "Properties"
Image1.Picture = LoadPicture(Form1.CommonDialog1.FileName)
St = Form1.CommonDialog1.FileTitle
St = Right(St, 3)
Label1.Caption = Form1.CommonDialog1.FileTitle
Label3.Caption = Form1.CommonDialog1.FileName
Label5.Caption = FileLen(Label3.Caption) & " Byte(s)"
Select Case St
Case "jpg":
Label7.Caption = " JPEG File"
Case "gif":
Label7.Caption = "GIFF File"
Case "tif":
Label7.Caption = "TIFF File"
Case "bmp":
Label7.Caption = "Windows Bitmap File"
Case "wmf":
Label7.Caption = "Windows Meta File"
End Select
End Sub

Private Sub TabStrip1_Click()

End Sub

