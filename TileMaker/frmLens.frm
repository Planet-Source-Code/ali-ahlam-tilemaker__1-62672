VERSION 5.00
Begin VB.Form frmLens 
   Caption         =   "Zoomed Tile"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   ClipControls    =   0   'False
   Icon            =   "frmLens.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer LensTimer 
      Interval        =   30
      Left            =   2760
      Top             =   2760
   End
   Begin VB.PictureBox picLens 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   2895
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.Image Pointer 
         Height          =   480
         Left            =   1200
         Picture         =   "frmLens.frx":030A
         Top             =   1080
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmLens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LX, LY As Integer



Private Sub Form_Activate()
MakeIt3D picLens, 5, 5, True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyF1
frmMain.helpMnu_Click 0
Case vbKeyF2
frmMain.imageMnu_Click 13
Case vbKeyF3
frmMain.loadfrom_Click
Case vbKeyF4
frmMain.optionsMnu_Click 0
Case vbKeyF5
frmMain.optionsMnu_Click 2
Case vbKeyF6
frmMain.Brush_set_Click
Case vbKeyF7
frmMain.imageMnu_Click 0
Case vbKeyF8
frmMain.imageMnu_Click 1
'Case vbKeyF9
'Case vbKeyF10
'Case vbKeyF11
Case vbKeyF12
frmMain.helpMnu_Click 2
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 1 'Ctrl+A
frmMain.add_to_Gallery_Click
'Case 2 'Ctrl+B
Case 3 'Ctrl+C
frmMain.editMnu_Click 4
Case 4 'Ctrl+D
frmMain.viewMnu_Click 0
Case 5 'Ctrl+E
frmMain.imageMnu_Click 9
Case 6 'Ctrl+F
frmMain.imageMnu_Click 7
'Case 7 'Ctrl+G
'Case 8 'Ctrl+H
'Case 9 'Ctrl+I
'Case 10 'Ctrl+J
Case 11 'Ctrl+K
frmMain.imageMnu_Click 11
'Case 12 'Ctrl+L
Case 13 'Ctrl+M
frmMain.imageMnu_Click 4
Case 14 'Ctrl+N
frmMain.fileMnu_Click 0
Case 15 'Ctrl+O
frmMain.fileMnu_Click 1
Case 16 'Ctrl+P
frmMain.fileMnu_Click 6
'Case 17 'Ctrl+Q
'Case 18 'Ctrl+R
Case 19 'Ctrl+S
frmMain.fileMnu_Click 3
'Case 20 'Ctrl+T
'Case 21 'Ctrl+U
Case 22 'Ctrl+V
frmMain.editMnu_Click 5
'Case 23 'Ctrl+W
Case 24 'Ctrl+X
frmMain.editMnu_Click 3
Case 25 'Ctrl+Y
frmMain.editMnu_Click 1
Case 26 'Ctrl+Z
frmMain.editMnu_Click 0
End Select
End Sub

Private Sub Form_Load()
Float_Form Me, frmMain
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmToolBox.Show 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Cls
Me.Width = (Me.Height - 280)
If Me.Width <= 3590 Then GoTo SkipMin
If Me.Width >= 5590 Then GoTo SkipMax

picLens.Width = Me.ScaleWidth - 220
picLens.Height = Me.ScaleHeight - 220

Draw_Preview frmMain.picWork, picLens

Exit Sub

SkipMin:
Me.Height = 3590
Me.Width = (Me.Height - 280)
picLens.Width = Me.ScaleWidth - 220
picLens.Height = Me.ScaleHeight - 220
Me.Cls
Draw_Preview frmMain.picWork, picLens
Exit Sub

SkipMax:
Me.Height = 5590
Me.Width = (Me.Height - 280)
picLens.Width = Me.ScaleWidth - 220
picLens.Height = Me.ScaleHeight - 220
Me.Cls
Draw_Preview frmMain.picWork, picLens
Exit Sub
End Sub

Private Sub LensTimer_Timer()
MakeIt3D picLens, 3, 4, True

End Sub

Private Sub mnuClose_Click()
frmMain.mnuView(0).Checked = False
picLens.Picture = LoadPicture()
Unload Me
End Sub


Private Sub picLens_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LX = Int(X / (picLens.Width / frmMain.picWork.ScaleWidth))
LY = Int(Y / (picLens.Height / frmMain.picWork.ScaleHeight))
frmMain.PicWorkMouseDown Button, Shift, LX, LY
End Sub

Private Sub picLens_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 0 Then
X = 0
End If
If Y < 0 Then
Y = 0
End If
If X >= picLens.ScaleWidth Then
X = picLens.ScaleWidth + 10
End If
If Y >= picLens.ScaleHeight Then
Y = picLens.ScaleHeight + 10
End If
LX = Int(X / (picLens.Width / frmMain.picWork.ScaleWidth)) '+ 210
LY = Int(Y / (picLens.Height / frmMain.picWork.ScaleHeight)) ' + 220
frmMain.PicWorkMouseMove Button, Shift, LX, LY, False
End Sub

Private Sub picLens_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 0 Then
X = 0
End If
If Y < 0 Then
Y = 0
End If
If X > picLens.ScaleWidth Then
X = picLens.ScaleWidth + 10
End If
If Y > picLens.ScaleHeight Then
Y = picLens.ScaleHeight + 10
End If

LX = Int(X / (picLens.Width / frmMain.picWork.ScaleWidth)) '+ 210
LY = Int(Y / (picLens.Height / frmMain.picWork.ScaleHeight)) ' + 220
frmMain.PicWorkMouseUp Button, Shift, LX, LY
End Sub
