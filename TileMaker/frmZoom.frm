VERSION 5.00
Begin VB.Form frmZoom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zoom Window"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7290
   Icon            =   "frmZoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   486
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picZoomed 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuAct 
      Caption         =   "&Action"
      Begin VB.Menu mnuAction 
         Caption         =   "&Save Background As..."
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuAction 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuAction 
         Caption         =   "&Close"
         Index           =   2
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmMain.stbar.Text = "Zoomed Background"
End Sub


Private Sub Form_Resize()
picZoomed.Move 0, 0
picZoomed.Width = Me.ScaleWidth
picZoomed.Height = Me.ScaleHeight

BmpTile picZoomed, frmMain.picWork
picZoomed.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.stbar.Text = ""
End Sub


Private Sub mnuAction_Click(Index As Integer)
Select Case Index
Case 0
picZoomed.Picture = picZoomed.Image
Save_Pic picZoomed
Case 2
Unload Me
End Select
End Sub
