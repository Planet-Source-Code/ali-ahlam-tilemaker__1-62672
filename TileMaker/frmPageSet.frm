VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmPageSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "frmPageSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      ToolTipText     =   "Cancel"
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox picCont 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   1
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   3255
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Frame Frame2 
         Caption         =   "Brush Sizes:"
         Height          =   2055
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   3040
         Begin ComCtl2.UpDown RSBar 
            Height          =   285
            Left            =   2650
            TabIndex        =   25
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327680
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "rSize"
            BuddyDispid     =   196613
            OrigLeft        =   2880
            OrigTop         =   1560
            OrigRight       =   3120
            OrigBottom      =   1815
            Max             =   25
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox rSize 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1560
            Width           =   615
         End
         Begin ComCtl2.UpDown BSBar 
            Height          =   285
            Left            =   1690
            TabIndex        =   23
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327680
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "bSize"
            BuddyDispid     =   196615
            OrigLeft        =   1800
            OrigTop         =   1560
            OrigRight       =   2040
            OrigBottom      =   1815
            Max             =   25
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox bSize 
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1560
            Width           =   615
         End
         Begin VB.PictureBox picRect 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   895
            Left            =   2040
            ScaleHeight     =   58
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   58
            TabIndex        =   20
            Top             =   600
            Width           =   895
         End
         Begin VB.PictureBox picBox 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   895
            Left            =   1080
            ScaleHeight     =   58
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   58
            TabIndex        =   18
            Top             =   600
            Width           =   895
         End
         Begin VB.TextBox cRadius 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1560
            Width           =   615
         End
         Begin VB.PictureBox picCirc 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   895
            Left            =   120
            ScaleHeight     =   58
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   58
            TabIndex        =   14
            Top             =   600
            Width           =   895
         End
         Begin ComCtl2.UpDown RBar 
            Height          =   285
            Left            =   720
            TabIndex        =   15
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327680
            Value           =   1
            BuddyControl    =   "cRadius"
            BuddyDispid     =   196618
            OrigLeft        =   600
            OrigTop         =   1320
            OrigRight       =   840
            OrigBottom      =   1695
            Max             =   25
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Rectangle"
            Height          =   195
            Left            =   2040
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Box"
            Height          =   195
            Left            =   1080
            TabIndex        =   19
            Top             =   360
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Circle"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   390
         End
      End
   End
   Begin VB.PictureBox picCont 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   0
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.Frame Frame1 
         Caption         =   "Page Settings"
         Height          =   1575
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3015
         Begin VB.CheckBox chkAuto 
            Caption         =   "Auto Size"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Auto Size"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox chkProp 
            Caption         =   "Proportional"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Keep proprotional"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox cHeight 
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Text            =   "0"
            ToolTipText     =   "Page Height"
            Top             =   570
            Width           =   615
         End
         Begin VB.TextBox cWidth 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Text            =   "0"
            ToolTipText     =   "Page Width"
            Top             =   570
            Width           =   615
         End
         Begin ComCtl2.UpDown Wid 
            Height          =   285
            Left            =   736
            TabIndex        =   7
            Top             =   570
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327680
            BuddyControl    =   "cWidth"
            BuddyDispid     =   196628
            OrigLeft        =   3600
            OrigTop         =   2880
            OrigRight       =   3840
            OrigBottom      =   3615
            Max             =   150
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown Hgt 
            Height          =   285
            Left            =   2040
            TabIndex        =   8
            Top             =   570
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327680
            BuddyControl    =   "cHeight"
            BuddyDispid     =   196627
            OrigLeft        =   4680
            OrigTop         =   3360
            OrigRight       =   4920
            OrigBottom      =   3615
            Max             =   150
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblW 
            AutoSize        =   -1  'True
            Caption         =   "Width:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "Height:"
            Height          =   195
            Left            =   1440
            TabIndex        =   9
            Top             =   360
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdSize 
         Caption         =   "&Apply Size"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Apply Size"
         Top             =   1680
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPageSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BSBar_Change()
picBox.Picture = LoadPicture()
picBox.Line ((picCirc.ScaleWidth / 2) - BSBar.Value, (picCirc.ScaleHeight / 2) - BSBar.Value)-((picCirc.ScaleWidth / 2) + BSBar.Value, (picCirc.ScaleHeight / 2) + BSBar.Value), vbBlack, B
BbSiz = BSBar.Value
End Sub

Private Sub cHeight_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then 'allow only numerics
If KeyAscii = 8 Then 'backspace
Else
KeyAscii = 0
End If
End If
End Sub


Private Sub cmdClose_Click()
If Val(GetIniVal("General", "SaveSettings", IniFile)) = 1 Then
SetIniVal "Brush", "BcRad", BcRad, IniName
SetIniVal "Brush", "BbSiz", BbSiz, IniName
SetIniVal "Brush", "BrSiz", BrSiz, IniName
SetIniVal "PageSettings", "Width", cWidth.Text, IniName
SetIniVal "PageSettings", "Height", cHeight.Text, IniName
SetIniVal "PageSettings", "Proportional", chkProp.Value, IniName
SetIniVal "PageSettings", "AutoSize", chkAuto.Value, IniName
End If
Me.Hide
End Sub


Private Sub cmdSize_Click()
If MsgBox("This will erase the current Image" & vbCrLf & _
"Continue anyway?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
frmMain.File_New
frmMain.picWork.Width = Val(cWidth.Text) * 15
frmMain.picWork.Height = Val(cHeight.Text) * 15
End If
End Sub

Private Sub chkAuto_Click()
Select Case chkAuto.Value
Case 0 'uncheck
frmMain.picWork.AutoSize = False
cmdSize.Enabled = True
cWidth.Enabled = True
lblW.Enabled = True
Wid.Enabled = True
Hgt.Enabled = True
lblH.Enabled = True
cHeight.Enabled = True
chkProp.Enabled = True
Case 1 'check
frmMain.picWork.AutoSize = True
cmdSize.Enabled = False
cWidth.Enabled = False
lblW.Enabled = False
Wid.Enabled = False
Hgt.Enabled = False
lblH.Enabled = False
cHeight.Enabled = False
chkProp.Enabled = False
End Select
End Sub

Private Sub chkProp_Click()
Dim FEnable As Boolean
Select Case chkProp.Value
Case 0 'uncheck
FEnable = True
Case 1 'checked
On Error GoTo HandleSizeErr
Hgt.Value = cWidth
FEnable = False
End Select
Hgt.Enabled = FEnable
lblH.Enabled = FEnable
cHeight.Enabled = FEnable
Exit Sub

HandleSizeErr:
Exit Sub
End Sub


Private Sub cHeight_LostFocus()
If Val(cHeight.Text) > 150 Then
MsgBox "Value Out of Range [0-150]", vbCritical
cHeight.Text = 150
End If
End Sub




Private Sub cWidth_Change()
If chkProp.Value = 1 Then
cHeight.Text = cWidth.Text
End If
End Sub

Private Sub cWidth_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then 'allow only numerics
If KeyAscii = 8 Then 'backspace
Else
KeyAscii = 0
End If
End If
End Sub

Private Sub cWidth_LostFocus()
If Val(cWidth.Text) > 150 Then
MsgBox "Value Out of Range [0-150]", vbCritical
cWidth.Text = 150
End If
End Sub

Private Sub Form_Activate()
cWidth = frmMain.picWork.ScaleWidth
cHeight = frmMain.picWork.ScaleHeight
frmMain.stbar.Text = "Page Settings"
End Sub


Private Sub Form_Load()

frmMain.stbar.Text = "Loading Page settings..."
chkProp.Value = Val(GetIniVal("PageSettings", "Proportional", IniName))
chkAuto.Value = Val(GetIniVal("PageSettings", "AutoSize", IniName))

RBar.Value = BcRad
BSBar.Value = BbSiz
RSBar.Value = BrSiz
RBar_Change
BSBar_Change
RSBar_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.stbar.Text = ""
End Sub

Private Sub RBar_Change()
picCirc.Picture = LoadPicture()
picCirc.Circle (picCirc.ScaleWidth / 2, picCirc.ScaleHeight / 2), cRadius.Text, vbBlack
BcRad = RBar.Value
End Sub

Private Sub RSBar_Change()
picRect.Picture = LoadPicture()
picRect.Line ((picRect.ScaleWidth / 2) - RSBar.Value - 5, (picRect.ScaleHeight / 2) - RSBar.Value - 2)-((picRect.ScaleWidth / 2) + RSBar.Value + 5, (picRect.ScaleHeight / 2) + RSBar.Value + 2), vbBlack, B
BrSiz = RSBar.Value
End Sub

Private Sub Wid_Change()
If chkProp.Value = 1 Then
cHeight.Text = cWidth.Text
End If
End Sub
