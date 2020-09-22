VERSION 5.00
Begin VB.Form frmPicSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tile Maker - Load Image..."
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7005
   ControlBox      =   0   'False
   Icon            =   "frmPicSel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picShowTiled 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   4860
      Width           =   255
   End
   Begin VB.PictureBox picTiled 
      AutoRedraw      =   -1  'True
      Height          =   2295
      Left            =   4440
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   16
      Top             =   2800
      Width           =   2535
   End
   Begin VB.PictureBox PicSelTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   120
      ScaleHeight     =   810
      ScaleWidth      =   810
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picCont 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2655
      Index           =   1
      Left            =   4440
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   9
      Top             =   120
      Width           =   2535
      Begin VB.PictureBox PicClipBox 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2325
         Left            =   75
         ScaleHeight     =   155
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   155
         TabIndex        =   10
         Top             =   75
         Width           =   2325
      End
   End
   Begin VB.HScrollBar HBar 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   200
      Left            =   30
      TabIndex        =   13
      Top             =   4860
      Width           =   4100
   End
   Begin VB.VScrollBar VBar 
      Enabled         =   0   'False
      Height          =   2055
      LargeChange     =   200
      Left            =   4120
      TabIndex        =   12
      Top             =   2800
      Width           =   255
   End
   Begin VB.PictureBox Container 
      Height          =   2055
      Left            =   30
      ScaleHeight     =   1995
      ScaleWidth      =   4035
      TabIndex        =   11
      Top             =   2800
      Width           =   4100
      Begin VB.PictureBox picWorkT 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   2775
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.Timer MainTimer 
      Interval        =   100
      Left            =   2520
      Top             =   5640
   End
   Begin VB.PictureBox picCont 
      Height          =   2655
      Index           =   0
      Left            =   30
      ScaleHeight     =   2595
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   120
      Width           =   4400
      Begin VB.ComboBox Ftypes 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
      End
      Begin VB.FileListBox ImFil 
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   2280
         TabIndex        =   3
         Top             =   300
         Width           =   1935
      End
      Begin VB.DriveListBox ImDriv 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   2055
      End
      Begin VB.DirListBox ImDir 
         ForeColor       =   &H00C00000&
         Height          =   1440
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "File Type:"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   8
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Drive:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "File Name:"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   6
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Folder:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   6960
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   0
      X2              =   6960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "L&oad"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Capture Screen"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Return to TileMaker"
         Index           =   3
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmPicSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SelCancel As Boolean
Private cX, cY As Single
Private Xpos, Ypos As Single



Property Get SelectCancel() As Boolean
SelectCancel = SelCancel
End Property

Private Sub Form_Load()
frmMain.Visible = False
frmMain.stbar.Text = "Load Image from..."
AddFTypes
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMain.stbar.Text = ""
frmMain.Visible = True
End Sub

Private Sub Ftypes_Click()
ImFil.Pattern = Left(Right(Ftypes.Text, 6), 5)
End Sub

Private Sub HBar_Change()
HBar_Scroll
End Sub

Private Sub HBar_Scroll()
picWorkT.Left = -HBar.Value
End Sub

Private Sub ImDir_Change()
ImFil = ImDir
ChDrive (ImDriv)
ChDir (ImDir.Path)
End Sub

Private Sub ImDriv_Change()
On Error GoTo handleDriv
ImDir = ImDriv
ChDrive (ImDriv)
ChDir (ImDir.Path)
Exit Sub
handleDriv:
If MsgBox("Selected drive not ready.", vbCritical + vbRetryCancel) = vbRetry Then
Resume
Else
ImDriv.Refresh
ImDir.Refresh
Exit Sub
End If

End Sub

Private Sub ImFil_Click()
ChDrive (ImDriv)
ChDir (ImDir.Path)
Load_Pict (ImFil.Filename)
End Sub

Private Sub ImFil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImFil_Click
End Sub

Private Sub ImFil_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChDrive (ImDriv)
ChDir (ImDir.Path)
End Sub



Private Sub mnuFile_Click(Index As Integer)
Dim Wstate As Integer
Select Case Index
Case 0 'open
SelCancel = False
With frmMain.picWork
.Width = PicSelTile.Width
.Height = PicSelTile.Height
.Picture = PicSelTile.Picture
End With
frmMain.Save_UndoAction
Unload Me
Case 1 'capture screen
Me.Visible = False
PauseFor 1
Draw_ScreenTo picWorkT
Me.Visible = True
HBar.Max = (picWorkT.Width - Container.ScaleWidth): HBar.Enabled = True
VBar.Max = (picWorkT.Height - Container.ScaleHeight): VBar.Enabled = True
picWorkT.Move 0, 0
picWorkT.Visible = True
Case 3 'return
SelCancel = True
Unload Me
End Select
End Sub

Sub AddFTypes()
With Ftypes
.AddItem "Bitmap files (*.bmp)"
.AddItem "JPEG files (*.jpg)"
.AddItem "GIFF files (*.gif)"
.AddItem "Icon files (*.ico)"
.AddItem "Cursor files (*.cur)"
.AddItem "Meta files (*.wmf)"
.AddItem "PCX files (*.pcx)"
.ListIndex = 0
End With
End Sub


Sub Load_Pict(ByVal Fname As String)
On Error GoTo handleLoadERR
Screen.MousePointer = 11
picWorkT.Move 0, 0
VBar.Value = 0
VBar_Change
HBar.Value = 0
HBar_Change

picWorkT.Picture = LoadPicture(Fname)
picWorkT.Picture = picWorkT.Image
Paint_ClipSize picWorkT, PicClipBox
picWorkT.Visible = True

'adjust scrol bars Horz
If picWorkT.Width > Container.ScaleWidth Then
HBar.Max = (picWorkT.Width - Container.ScaleWidth): HBar.Enabled = True ': px = 0
Else: HBar.Enabled = False
End If
'adjust scrol bars Vert
If picWorkT.Height > Container.ScaleHeight Then
VBar.Max = (picWorkT.Height - Container.ScaleHeight): VBar.Enabled = True ': py = 0
Else: VBar.Enabled = False
End If
Screen.MousePointer = 0
Exit Sub
handleLoadERR:
MsgBox Err.Description, vbCritical
Screen.MousePointer = 0
Exit Sub
End Sub


Private Sub picWorkT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cX = X
cY = Y
Draw_Box cX, cY, 54
mnuFile(0).Enabled = True
End Sub

Private Sub picWorkT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And 1 Then
cX = X
cY = Y
Draw_Box cX, cY, 54
End If

End Sub

Private Sub picWorkT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draw_Box cX, cY, 54
'draw tile part
Draw_TilePart Xpos, Ypos, 54
End Sub

Private Sub VBar_Change()
VBar_Scroll
End Sub

Private Sub VBar_Scroll()
picWorkT.Top = -VBar.Value
End Sub

Sub Draw_Box(ByVal X As Single, ByVal Y As Single, ByVal TWid As Single)
Dim Wid, Hgt
Wid = picWorkT.ScaleWidth
Hgt = picWorkT.ScaleHeight

If (X - ((TWid * 15) / 2)) <= 0 Then
X = (TWid * 15) / 2
ElseIf (X + ((TWid * 15) / 2)) >= (Wid - 40) Then
X = Wid - (TWid * 15) / 2 - 40
End If

If (Y - ((TWid * 15) / 2)) <= 0 Then
Y = (TWid * 15) / 2
ElseIf (Y + ((TWid * 15) / 2)) >= (Hgt - 40) Then
Y = Hgt - (TWid * 15) / 2 - 40
End If

'cordinates for the tile part
Xpos = X
Ypos = Y

'clean first
picWorkT.Cls
picWorkT.Refresh

'draw box
picWorkT.Line (X - ((TWid * 15) / 2), Y - ((TWid * 15) / 2))-(X + ((TWid * 15) / 2), Y + ((TWid * 15) / 2)), , B

End Sub

Private Sub Draw_TilePart(ByVal curX As Single, ByVal curY As Single, ByVal TilWid As Single)
On Error Resume Next
    PicSelTile.Picture = LoadPicture()
    PicSelTile.PaintPicture picWorkT.Picture, 0, 0, PicSelTile.ScaleWidth, PicSelTile.ScaleHeight, _
    (curX - ((TilWid * 15) / 2)), (curY - ((TilWid * 15) / 2)), _
    PicSelTile.ScaleWidth, PicSelTile.ScaleHeight, vbSrcCopy
    PicSelTile.Picture = PicSelTile.Image
    PicSelTile.ScaleMode = 3
    BmpTile picTiled, PicSelTile
    PicSelTile.ScaleMode = 1
Exit Sub
End Sub
