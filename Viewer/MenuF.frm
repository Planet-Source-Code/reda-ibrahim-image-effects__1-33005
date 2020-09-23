VERSION 5.00
Begin VB.Form MenuF 
   Caption         =   "Form11"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuLayers 
      Caption         =   "Layers"
      Begin VB.Menu mnuNewL 
         Caption         =   "&Layer Property"
      End
      Begin VB.Menu mnuDeleteL 
         Caption         =   "&Delete Layer"
      End
   End
End
Attribute VB_Name = "MenuF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub mnuDeleteL_Click()
If Painter.TreeView1.Nodes.Count > 0 Then
Painter.TreeView1.Nodes.Remove (Painter.TreeView1.SelectedItem.Index)
'LKey = LKey - 1
End If
End Sub

