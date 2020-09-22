VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmEffects 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Effects"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "OK"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox picEffectsCont 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   2
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   4215
      TabIndex        =   46
      Top             =   1560
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdWavePreview 
         Caption         =   "Preview"
         Height          =   375
         Left            =   1200
         TabIndex        =   56
         Top             =   2040
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Wave"
         Height          =   1695
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   3975
         Begin VB.TextBox txtWaveLen 
            Height          =   285
            Left            =   3000
            TabIndex        =   51
            Text            =   "2"
            Top             =   960
            Width           =   495
         End
         Begin VB.CheckBox chkWave 
            Caption         =   "Vertical"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   50
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkWave 
            Caption         =   "Horizontal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.TextBox txtAmplitude 
            Height          =   285
            Left            =   1080
            TabIndex        =   48
            Text            =   "1"
            Top             =   960
            Width           =   495
         End
         Begin ComCtl2.UpDown UDWlen 
            Height          =   285
            Left            =   3480
            TabIndex        =   52
            Top             =   960
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   2
            BuddyControl    =   "txtwavelen"
            BuddyDispid     =   196613
            OrigLeft        =   3480
            OrigTop         =   960
            OrigRight       =   3720
            OrigBottom      =   1245
            Min             =   2
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UDAmplitude 
            Height          =   285
            Left            =   1560
            TabIndex        =   53
            Top             =   960
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "txtamplitude"
            BuddyDispid     =   196611
            OrigLeft        =   720
            OrigTop         =   1080
            OrigRight       =   960
            OrigBottom      =   1335
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbldes 
            AutoSize        =   -1  'True
            Caption         =   "Wavelength:"
            Height          =   195
            Index           =   1
            Left            =   2040
            TabIndex        =   55
            Top             =   960
            Width           =   915
         End
         Begin VB.Label lbldes 
            AutoSize        =   -1  'True
            Caption         =   "Amplitude:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   960
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox picEffectsCont 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   1
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   4215
      TabIndex        =   30
      Top             =   1560
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdPreviewRep 
         Caption         =   "Preview"
         Height          =   375
         Left            =   1200
         TabIndex        =   44
         ToolTipText     =   "Preview"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Click on the image to select a color to be replaced"
         Height          =   1935
         Left            =   120
         TabIndex        =   31
         Top             =   0
         Width           =   3975
         Begin VB.CommandButton cmdApplyNow 
            Caption         =   "Apply Now"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   "Applies Changes"
            Top             =   1440
            Width           =   2175
         End
         Begin VB.HScrollBar ColBar1 
            Height          =   225
            Index           =   0
            Left            =   360
            Max             =   255
            TabIndex        =   37
            Top             =   600
            Width           =   975
         End
         Begin VB.HScrollBar ColBar1 
            Height          =   225
            Index           =   1
            Left            =   360
            Max             =   255
            TabIndex        =   36
            Top             =   840
            Width           =   975
         End
         Begin VB.HScrollBar ColBar1 
            Height          =   225
            Index           =   2
            Left            =   360
            Max             =   255
            TabIndex        =   35
            Top             =   1080
            Width           =   975
         End
         Begin VB.PictureBox picWithCol 
            Height          =   690
            Left            =   1440
            ScaleHeight     =   630
            ScaleWidth      =   795
            TabIndex        =   34
            Top             =   600
            Width           =   855
         End
         Begin VB.PictureBox picSelCol 
            Height          =   375
            Left            =   2640
            ScaleHeight     =   315
            ScaleWidth      =   1035
            TabIndex        =   33
            Top             =   590
            Width           =   1095
         End
         Begin VB.PictureBox repColor 
            Height          =   375
            Left            =   2640
            ScaleHeight     =   315
            ScaleWidth      =   1035
            TabIndex        =   32
            Top             =   1420
            Width           =   1095
         End
         Begin VB.Label lblRe 
            AutoSize        =   -1  'True
            Caption         =   "R:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   165
         End
         Begin VB.Label lblRe 
            AutoSize        =   -1  'True
            Caption         =   "G:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   165
         End
         Begin VB.Label lblRe 
            AutoSize        =   -1  'True
            Caption         =   "B:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   41
            Top             =   1080
            Width           =   150
         End
         Begin VB.Label lblRe 
            AutoSize        =   -1  'True
            Caption         =   "Selected Color:"
            Height          =   195
            Index           =   3
            Left            =   2640
            TabIndex        =   40
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblRe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Replace with:"
            Height          =   195
            Index           =   4
            Left            =   2640
            TabIndex        =   39
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Color Combination"
            Height          =   195
            Left            =   600
            TabIndex        =   38
            Top             =   360
            Width           =   1275
         End
      End
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   4440
      MousePointer    =   99  'Custom
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Timer EffectsTimer 
      Interval        =   30
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   1560
      MousePointer    =   99  'Custom
      ScaleHeight     =   86
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   86
      TabIndex        =   2
      Top             =   120
      Width           =   1290
   End
   Begin VB.PictureBox picEffectsCont 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   0
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   4215
      TabIndex        =   3
      ToolTipText     =   "Cloth Effect"
      Top             =   1560
      Width           =   4215
      Begin VB.CommandButton cmdPreviewCloth 
         Caption         =   "Preview"
         Height          =   375
         Left            =   1200
         TabIndex        =   29
         Top             =   2040
         Width           =   975
      End
      Begin ComCtl2.UpDown RaiseY 
         Height          =   285
         Left            =   3975
         TabIndex        =   25
         ToolTipText     =   "Raise Y"
         Top             =   1440
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRaiseY"
         BuddyDispid     =   196624
         OrigLeft        =   3240
         OrigTop         =   1440
         OrigRight       =   3480
         OrigBottom      =   1695
         Max             =   8
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtRaiseY 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0"
         Top             =   1440
         Width           =   375
      End
      Begin ComCtl2.UpDown RaiseX 
         Height          =   285
         Left            =   3975
         TabIndex        =   23
         ToolTipText     =   "Raise X"
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRaiseX"
         BuddyDispid     =   196626
         OrigLeft        =   3480
         OrigTop         =   1320
         OrigRight       =   3720
         OrigBottom      =   1575
         Max             =   8
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtRaiseX 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0"
         Top             =   1080
         Width           =   375
      End
      Begin ComCtl2.UpDown StepY 
         Height          =   285
         Left            =   3975
         TabIndex        =   20
         ToolTipText     =   "Steps Y"
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSY"
         BuddyDispid     =   196628
         OrigLeft        =   3840
         OrigTop         =   1200
         OrigRight       =   4080
         OrigBottom      =   1455
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSY 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "1"
         Top             =   600
         Width           =   375
      End
      Begin ComCtl2.UpDown StepX 
         Height          =   285
         Left            =   3975
         TabIndex        =   17
         ToolTipText     =   "Steps X"
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSX"
         BuddyDispid     =   196630
         OrigLeft        =   4080
         OrigTop         =   240
         OrigRight       =   4320
         OrigBottom      =   495
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSX 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame fram 
         Caption         =   "Channel"
         Height          =   855
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1695
         Begin VB.OptionButton Clothchannel 
            Caption         =   "Monochrome"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Channel Monochrome"
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Clothchannel 
            Caption         =   "RGB"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Channel RGB"
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.HScrollBar ColBar 
         Height          =   200
         Index           =   2
         Left            =   240
         Max             =   255
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.HScrollBar ColBar 
         Height          =   200
         Index           =   1
         Left            =   240
         Max             =   255
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.HScrollBar ColBar 
         Height          =   200
         Index           =   0
         Left            =   240
         Max             =   255
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblRA 
         AutoSize        =   -1  'True
         Caption         =   "Raise Y:"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   27
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label lblRA 
         AutoSize        =   -1  'True
         Caption         =   "Raise X:"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   26
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblST 
         AutoSize        =   -1  'True
         Caption         =   "Step Y:"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   21
         Top             =   600
         Width           =   525
      End
      Begin VB.Label lblST 
         AutoSize        =   -1  'True
         Caption         =   "Step X:"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   18
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblCa 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   12
         Top             =   1440
         Width           =   150
      End
      Begin VB.Label lblCa 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label lblCa 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   960
         Width           =   165
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "Value B"
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         ToolTipText     =   "Value G"
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         ToolTipText     =   "Value R"
         Top             =   960
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private meCancelled As Boolean
Private SelColor As Long
Private WithColor As Long


Public Property Get Didcancel() As Boolean
Didcancel = meCancelled
End Property

Private Sub cmdApplyNow_Click()
cmdPreviewRep_Click
frmMain.picWork.Picture = picImage.Image
picImage.Picture = frmMain.picWork.Picture
End Sub

Private Sub cmdPreviewCloth_Click()
    frmMain.ProgBar.Visible = True
    frmMain.ProgBar.Cls
    Cloth_Effect frmMain.picWork, picImage, txtSX.Text, txtSY.Text, _
    ColBar(0).Value, ColBar(1).Value, ColBar(2).Value, _
    txtRaiseX.Text, txtRaiseY.Text, Clothchannel(0).Value
    picImage.Picture = picImage.Image
    frmMain.ProgBar.Cls
    frmMain.ProgBar.Visible = False
    cmds(0).Enabled = True
End Sub

Private Sub cmds_Click(Index As Integer)
Select Case Index
Case 0
meCancelled = False
frmMain.picWork.Picture = picImage.Image
Unload Me
Case 1
meCancelled = True
Unload Me
End Select
End Sub

Private Sub cmdWavePreview_Click()
DrawWaves frmMain.picWork, picImage, UDAmplitude.Value, UDWlen.Value, chkWave(0).Value, chkWave(1).Value
cmds(0).Enabled = True
End Sub

Private Sub ColBar_Change(Index As Integer)
lbl(Index).Caption = ColBar(Index).Value

End Sub


Private Sub EffectsTimer_Timer()
Draw_Preview picImage, picPreview
If Clothchannel(0).Value = True Then
ColBar(0).Enabled = True
ColBar(1).Enabled = True
ColBar(2).Enabled = True
Else
ColBar(0).Enabled = False
ColBar(1).Enabled = False
ColBar(2).Enabled = False
End If
End Sub

Private Sub Form_Load()
    picImage.Picture = frmMain.picWork.Image
    picImage.Width = frmMain.picWork.Width
    picImage.Height = frmMain.picWork.Height
End Sub

'''
Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picSelCol.BackColor = picPreview.Point(X, Y)
End Sub

Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then
If X < 0 Then X = 0
If Y < 0 Then Y = 0
If X > picPreview.ScaleWidth Then X = 0
If Y > picPreview.ScaleHeight Then Y = 0
picSelCol.BackColor = picPreview.Point(X, Y)
End If

If picEffectsCont(1).Visible = True Then
picPreview.MousePointer = 2
Else
picPreview.MousePointer = 99
End If
End Sub

Private Sub picPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SelColor = picSelCol.BackColor
End Sub

Private Sub ColBar1_Change(Index As Integer)
ColBar1_Scroll Index
End Sub

Private Sub ColBar1_Scroll(Index As Integer)
picWithCol.BackColor = RGB(ColBar1(0), ColBar1(1), ColBar1(2))
repColor.BackColor = picWithCol.BackColor
WithColor = repColor.BackColor
End Sub

Private Sub cmdPreviewRep_Click()
Replace_Color frmMain.picWork, picImage, SelColor, WithColor
cmds(0).Enabled = True
End Sub

