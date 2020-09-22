VERSION 5.00
Begin VB.Form FrmOffset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Offset Image"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer drawTimer 
      Interval        =   50
      Left            =   4440
      Top             =   2760
   End
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&Restore"
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.PictureBox picPrev 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5040
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&Apply"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Offset Style"
      Height          =   2415
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   2415
      Begin VB.OptionButton offsetOPt 
         Caption         =   "Both"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton offsetOPt 
         Caption         =   "Vertical"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton offsetOPt 
         Caption         =   "Horizontal"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.PictureBox picSample 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         ScaleHeight     =   137
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmds_Click(Index As Integer)
Select Case Index
Case 0 'OK
OffCancel = False
frmMain.picWork.Picture = picPrev.Picture
Unload Me
Case 1
frmMain.picWork.Picture = picPrev.Picture
offsetOPt(0).Value = False
offsetOPt(1).Value = False
offsetOPt(2).Value = False
Case 2
picPrev.Picture = frmMain.picWork.Picture
offsetOPt(0).Value = False
offsetOPt(1).Value = False
offsetOPt(2).Value = False
Case 3 'cancel
OffCancel = True
Unload Me
End Select
Draw_Preview picPrev, picSample
End Sub


Private Sub drawtimer_Timer()
Draw_Preview picPrev, picSample
End Sub

Private Sub Form_Load()
picPrev.Width = frmMain.picWork.Width 'change main and dest same size
picPrev.Height = frmMain.picWork.Height ' "     "   "    "    "    "
picPrev.Picture = frmMain.picWork.Picture 'get image
OffCancel = False 'offset cancel flag
cmds_Click 2
End Sub

Private Sub offsetOPt_Click(Index As Integer)
drawTimer.Enabled = False
Offset_Image frmMain.picWork, picPrev, Index
Draw_Preview picPrev, picSample
End Sub

