VERSION 5.00
Begin VB.Form frmReflect 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reflection"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer drawtimer 
      Interval        =   50
      Left            =   3960
      Top             =   2760
   End
   Begin VB.CommandButton cmds 
      Caption         =   "Restore"
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   11
      ToolTipText     =   "Close"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&Apply"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   9
      ToolTipText     =   "Apply Effect"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "OK"
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   4680
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   6
      Tag             =   "Tile #1"
      Top             =   360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox contReflect 
      BorderStyle     =   0  'None
      Height          =   2440
      Left            =   240
      ScaleHeight     =   2445
      ScaleWidth      =   2430
      TabIndex        =   1
      Top             =   480
      Width           =   2430
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   240
         ScaleHeight     =   1935
         ScaleWidth      =   1935
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton ReflectOpt 
         Height          =   1935
         Index           =   1
         Left            =   2190
         Picture         =   "frmReflect.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Mirror Position Right"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton ReflectOpt 
         Height          =   1935
         Index           =   0
         Left            =   0
         Picture         =   "frmReflect.frx":146A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Mirror Position Left"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton ReflectOpt 
         Height          =   255
         Index           =   3
         Left            =   250
         Picture         =   "frmReflect.frx":28D4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Mirror Position Bottom"
         Top             =   2190
         Width           =   1935
      End
      Begin VB.OptionButton ReflectOpt 
         Height          =   255
         Index           =   2
         Left            =   250
         Picture         =   "frmReflect.frx":3CCA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Mirror Position Top"
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   4680
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   0
      Tag             =   "Tile #1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Select the Mirror position"
      Height          =   195
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Width           =   1740
   End
End
Attribute VB_Name = "frmReflect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Reflect_Image(pic As PictureBox, pic2 As PictureBox, ByVal Direction As Integer)
Dim oldSmode(0 To 1) As Integer
Dim Wid, Hgt As Single
Dim StX1, StY1 As Single 'first half
Dim EndX1, EndY1 As Single
Dim StX2, StY2 As Single
Dim EndX2, EndY2 As Single
Dim StX3, StY3 As Single 'second half
Dim EndX3, EndY3 As Single
Dim StX4, StY4 As Single
Dim EndX4, EndY4 As Single
Dim pixels() As Long
Dim X As Single, Y As Single
Dim maxX As Single, maxY As Single

oldSmode(0) = pic.ScaleMode
oldSmode(1) = pic2.ScaleMode
pic.ScaleMode = vbPixels
pic2.ScaleMode = vbPixels

Wid = pic.ScaleWidth
Hgt = pic.ScaleHeight
pic2.Picture = LoadPicture()
pic2.Width = pic.Width
pic2.Height = pic.Height


Select Case Direction
Case 0 'vertical L
StX1 = -1
StY1 = 0
EndX1 = ((Wid / 2) + 1)
EndY1 = Hgt
StX2 = Wid
StY2 = 0
EndX2 = -((Wid / 2) + 1)
EndY2 = Hgt

StX3 = Wid / 2
StY3 = 0
EndX3 = Wid
EndY3 = Hgt
StX4 = Wid / 2
StY4 = 0
EndX4 = Wid
EndY4 = Hgt

Case 1 'vertical R
StX1 = ((Wid / 2))
StY1 = 0
EndX1 = Wid
EndY1 = Hgt
StX2 = (Wid / 2) - 1
StY2 = 0
EndX2 = -(Wid + 1)
EndY2 = Hgt

StX3 = 0
StY3 = 0
EndX3 = Wid / 2
EndY3 = Hgt
StX4 = 0
StY4 = 0
EndX4 = Wid / 2
EndY4 = Hgt
Case 2 'horizontal T
StX1 = 0
StY1 = -1
EndX1 = Wid
EndY1 = (Hgt / 2) + 1
StX2 = 0
StY2 = Hgt
EndX2 = Wid
EndY2 = -(Hgt / 2) - 1

StX3 = 0
StY3 = Hgt / 2
EndX3 = Wid
EndY3 = Hgt / 2
StX4 = 0
StY4 = Hgt / 2
EndX4 = Wid
EndY4 = Hgt / 2

Case 3 'horizontal B
StX1 = 0
StY1 = (Hgt / 2) - 1
EndX1 = Wid
EndY1 = Hgt
StX2 = 0
StY2 = Hgt / 2
EndX2 = Wid
EndY2 = -(Hgt)

StX3 = 0
StY3 = 0
EndX3 = Wid
EndY3 = Hgt / 2
StX4 = 0
StY4 = 0
EndX4 = Wid
EndY4 = Hgt / 2

End Select


pic2.PaintPicture pic.Picture, StX1, StY1, EndX1, EndY1, _
StX2, StY2, EndX2, EndY2, vbSrcCopy
pic2.PaintPicture pic.Picture, StX3, StY3, EndX3, EndY3, _
StX4, StY4, EndX4, EndY4, vbSrcCopy

If Direction = 1 Or Direction = 3 Then 'end most lines
If Direction = 1 Then
maxX = 0
maxY = Hgt
Else
maxX = Wid
maxY = 0
End If

'dimension pixels array
ReDim pixels(-1 To maxX, -1 To maxY) As Long

    'read pixels
    For X = -1 To maxX
        For Y = -1 To maxY
            pixels(X, Y) = pic.Point(X, Y)
        Next Y
    Next X
    'place them
    For X = -1 To maxX
        For Y = -1 To maxY
             If Direction = 1 Then
                pic2.PSet (Wid - X - 1, Y), pixels(X, Y)
             Else
                pic2.PSet (X, Hgt - Y - 1), pixels(X, Y)
             End If
        Next Y
    Next X
End If

pic2.Refresh
pic2.Picture = pic2.Image

pic.ScaleMode = oldSmode(0)
pic2.ScaleMode = oldSmode(1)
oldSmode(0) = 0
oldSmode(1) = 0
End Sub

Private Sub cmds_Click(Index As Integer)
Select Case Index
Case 0 'ok
frmMain.picWork.Picture = pic2.Image
frmMain.Save_UndoAction
Unload Me
Case 1 'apply
frmMain.picWork.Picture = pic2.Image
pic.Picture = frmMain.picWork.Picture
ReflectOpt(0).Value = False
ReflectOpt(1).Value = False
ReflectOpt(2).Value = False
ReflectOpt(3).Value = False
Case 2 'cancel
Unload Me
Case 3 'restore
pic.Width = frmMain.picWork.Width
pic.Height = frmMain.picWork.Height
pic.Picture = frmMain.picWork.Image

pic2.Width = frmMain.picWork.Width
pic2.Height = frmMain.picWork.Height
pic2.Picture = frmMain.picWork.Image
ReflectOpt(0).Value = False
ReflectOpt(1).Value = False
ReflectOpt(2).Value = False
ReflectOpt(3).Value = False
End Select
Draw_Preview pic2, picPreview
End Sub

Private Sub drawtimer_Timer()
pic.Width = frmMain.picWork.Width
pic.Height = frmMain.picWork.Height
pic.Picture = frmMain.picWork.Image

pic2.Width = frmMain.picWork.Width
pic2.Height = frmMain.picWork.Height
pic2.Picture = frmMain.picWork.Image

Draw_Preview frmMain.picWork, picPreview
End Sub

Private Sub Form_Load()
pic.Width = frmMain.picWork.Width
pic.Height = frmMain.picWork.Height
pic.Picture = frmMain.picWork.Image

pic2.Width = frmMain.picWork.Width
pic2.Height = frmMain.picWork.Height
pic2.Picture = frmMain.picWork.Image
Draw_Preview frmMain.picWork, picPreview

MakeIt3D contReflect, 1, 1, True
End Sub

Private Sub ReflectOpt_Click(Index As Integer)
drawtimer.Enabled = False
Reflect_Image pic, pic2, Index
Draw_Preview pic2, picPreview
End Sub


