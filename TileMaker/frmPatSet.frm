VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmPatSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pattern Settings"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   ControlBox      =   0   'False
   Icon            =   "frmPatSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmds 
      Caption         =   "Co&lors"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   23
      ToolTipText     =   "Set Pattern Colors"
      Top             =   4690
      Width           =   1095
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&Clear All"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   22
      ToolTipText     =   "Clear All Styles"
      Top             =   4690
      Width           =   1140
   End
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Done"
      Top             =   4690
      Width           =   1095
   End
   Begin VB.Timer MainTimer 
      Interval        =   50
      Left            =   2520
      Top             =   5520
   End
   Begin VB.PictureBox cont1 
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   50
      ScaleHeight     =   4620
      ScaleWidth      =   8070
      TabIndex        =   0
      Top             =   50
      Width           =   8070
      Begin VB.PictureBox contCOlors 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   2055
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton cmdColorOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   0
            TabIndex        =   20
            ToolTipText     =   "OK"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Frame Frame1 
            Caption         =   "Pattern Colors"
            Height          =   2000
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   2000
            Begin VB.OptionButton optColMode 
               Caption         =   "Use Combination"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   18
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton optColMode 
               Caption         =   "Use Selection"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox ColTxt 
               Height          =   285
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   1440
               Width           =   495
            End
            Begin VB.PictureBox picSelColors 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   365
               Left            =   360
               ScaleHeight     =   330
               ScaleWidth      =   1290
               TabIndex        =   13
               Top             =   600
               Width           =   1325
               Begin VB.PictureBox ColBlock 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   160
                  Index           =   0
                  Left            =   0
                  MouseIcon       =   "frmPatSet.frx":030A
                  MousePointer    =   99  'Custom
                  ScaleHeight     =   165
                  ScaleWidth      =   165
                  TabIndex        =   14
                  Top             =   0
                  Width           =   160
               End
            End
            Begin VB.PictureBox selectCol 
               AutoRedraw      =   -1  'True
               Height          =   200
               Left            =   1480
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   12
               Top             =   390
               Width           =   200
            End
            Begin ComCtl2.UpDown ColBar 
               Height          =   285
               Left            =   1335
               TabIndex        =   17
               Top             =   1440
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   327681
               Value           =   1
               BuddyControl    =   "ColTxt"
               BuddyDispid     =   196616
               OrigLeft        =   1440
               OrigTop         =   2280
               OrigRight       =   1680
               OrigBottom      =   2655
               Max             =   16
               Min             =   1
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label lblTcols 
               AutoSize        =   -1  'True
               Caption         =   "Colors:"
               Height          =   195
               Left            =   360
               TabIndex        =   19
               Top             =   1440
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox picWorkSet 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2000
         Left            =   120
         ScaleHeight     =   133
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   1
         Top             =   120
         Width           =   2000
      End
      Begin VB.Frame Frame2 
         Height          =   2000
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   2000
         Begin VB.PictureBox picPrev 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1665
            Left            =   160
            ScaleHeight     =   111
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   111
            TabIndex        =   9
            Top             =   210
            Width           =   1665
         End
      End
      Begin VB.HScrollBar StyleScrollH 
         Height          =   255
         LargeChange     =   1080
         Left            =   2160
         SmallChange     =   1080
         TabIndex        =   7
         Top             =   4200
         Width           =   5295
      End
      Begin VB.PictureBox Cont2 
         Height          =   4020
         Left            =   2200
         ScaleHeight     =   3960
         ScaleWidth      =   5370
         TabIndex        =   3
         Top             =   120
         Width           =   5430
         Begin VB.PictureBox picStyleBox 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   7935
            Left            =   120
            ScaleHeight     =   7935
            ScaleWidth      =   6510
            TabIndex        =   4
            Top             =   0
            Width           =   6510
            Begin VB.CheckBox chkSet 
               Caption         =   "Style 00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   0
               Width           =   1080
            End
         End
      End
      Begin VB.VScrollBar StyleScroll 
         Height          =   4020
         LargeChange     =   255
         Left            =   7680
         SmallChange     =   255
         TabIndex        =   2
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblind 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPatSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSet_Click(Index As Integer)
Dim ckhCnt As Integer
Dim curChks As Integer

If chkSet(Index).Value = 1 Then
cStyle(Index) = 1
Else
cStyle(Index) = 0
End If
Screen.MousePointer = 11
MakeGrid picWorkSet, 16
curChks = 0
For ckhCnt = 0 To nObject - 1
    If chkSet(ckhCnt).Value = 1 Then
        curChks = curChks + 1
    End If
Next ckhCnt
 If curChks = 0 Then
    cmds(0).Enabled = False
 Else
    cmds(0).Enabled = True
 End If
Screen.MousePointer = 0
End Sub

Private Sub chkSet_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblind.Caption = Index
Preview_Grid picPrev, 8, Index
End Sub

Private Sub cmdColorOK_Click()
contCOlors.Visible = False
cmds(0).Enabled = True
cmds(2).Enabled = True
End Sub

Private Sub cmds_Click(Index As Integer)
Dim chkI
Select Case Index
Case 0 'OK
Unload Me
Case 1 'Clear All Styles
Screen.MousePointer = 11
 For chkI = 0 To nObject - 1
   chkSet(chkI).Value = 0
   DoEvents
 Next chkI
Screen.MousePointer = 0
Case 2 'colors
contCOlors.Visible = True
cmds(0).Enabled = False
cmds(2).Enabled = False
End Select
End Sub

Private Sub cmds_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picPrev.Picture = LoadPicture()
End Sub

Private Sub ColBar_Change()
TotColrs = ColBar.Value
Screen.MousePointer = 11
MakeGrid picWorkSet, 16
Screen.MousePointer = 0
End Sub

Private Sub ColBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picPrev.Picture = LoadPicture()
End Sub

Private Sub ColBlock_Click(Index As Integer)
selectCol.BackColor = ColBlock(Index).BackColor
cPatColor = selectCol.BackColor 'selected color
MakeGrid picWorkSet, 16
End Sub

Private Sub ColBlock_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ColBlock(Index).MousePointer = 99
End Sub

Private Sub ColTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picPrev.Picture = LoadPicture()
End Sub

Private Sub Load_CBlocks()
Dim Bx
'first part
For Bx = 1 To 7
    Load ColBlock(Bx)
    ColBlock(Bx).Move (160 * Bx), 0
    ColBlock(Bx).Visible = True
Next Bx
'second part
For Bx = 8 To 15
    Load ColBlock(Bx)
    ColBlock(Bx).Move (160 * (Bx - 8)), 160
    ColBlock(Bx).Visible = True
Next Bx
Color_Blocks 0
End Sub

Sub Color_Blocks(ByVal cMode As Integer)
Dim BI
For BI = 0 To 15
If cMode = 0 Then
ColBlock(BI).BackColor = QBColor(BI)
Else
ColBlock(BI).BackColor = QBColor(8)
End If
Next BI
End Sub





Private Sub cont1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picPrev.Picture = LoadPicture()
End Sub

Private Sub Cont2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picPrev.Picture = LoadPicture()
End Sub

Private Sub Form_Load()
frmMain.stbar.Text = "Loading Pattern Settings..."
frmMain.Enabled = False
Load_CBlocks 'color pallete
optColMode(pColMode).Value = True
optColMode_Click (pColMode)
ColBar.Value = TotColrs
Init_Styles 'styles
frmMain.stbar.Text = "Pattern Settings"
StyleScroll.Max = picStyleBox.Height - Cont2.ScaleHeight
StyleScrollH.Max = picStyleBox.Width - Cont2.ScaleWidth
Me.Refresh
End Sub


Sub Init_Styles()
Dim styleI
Screen.MousePointer = 11
    Load_StyleBoxes 'styles checkbox
 For styleI = 0 To UBound(cStyle)
    frmMain.stbar.Text = "Initializing Styles..." & Format((styleI * 100) / UBound(cStyle), "0") & "% Complete"
   chkSet(styleI).Value = cStyle(styleI)
   DoEvents
 Next styleI
 'disable the rest
 For styleI = 0 To nObject - 1
   If chkSet(styleI).Caption = "" Then
     frmMain.stbar.Text = "Disabling Unavailables..." & styleI
     chkSet(styleI).Caption = "<N/A>!"
     chkSet(styleI).Enabled = False
   End If
   DoEvents
 Next styleI
 frmMain.stbar.Text = "Done."
 Screen.MousePointer = 0
End Sub

Sub Init_Boxes()
Cont2.Height = 3660
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.stbar.Text = ""
End Sub

Private Sub MainTimer_Timer()
If Screen.MousePointer = 11 Then
Me.Enabled = False
Else
Me.Enabled = True
End If
ColBar.Enabled = (optColMode(1).Value = True)
lblTcols.Enabled = (optColMode(1).Value = True)
If ColBar.Enabled = False Then
ColTxt.BackColor = &H80000003
Else
ColTxt.BackColor = vbWhite
End If
picSelColors.Enabled = (optColMode(0).Value = True)
selectCol.BackColor = cPatColor
End Sub

Private Sub optColMode_Click(Index As Integer)
Color_Blocks Index
pColMode = Index
Screen.MousePointer = 11
MakeGrid picWorkSet, 16
Screen.MousePointer = 0
If Val(GetIniVal("General", "SaveSettings", IniName)) = 1 Then
SetIniVal "General", "ColorMode", Index, IniName
End If
End Sub

Private Sub picStyleBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picPrev.Picture = LoadPicture()
End Sub

Private Sub picWorkSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picPrev.Picture = LoadPicture()
End Sub

Private Sub StyleScroll_Change()
StyleScroll_Scroll
End Sub

Private Sub StyleScroll_Scroll()
picStyleBox.Top = -StyleScroll.Value
picPrev.Picture = LoadPicture()
End Sub

Private Sub Load_StyleBoxes()
    Dim tX As Single, tY As Single
    Dim StartX As Single, StartY As Single
    
    Dim MaximumX As Single, MaximumY As Single
    Dim TileWidth As Integer, TileHeight As Integer
    
    MaximumX = picStyleBox.Width - chkSet(0).Width
    MaximumY = picStyleBox.Height - chkSet(0).Height
    
    TileWidth = chkSet(0).Width
    TileHeight = chkSet(0).Height
    StartX = 0
    StartY = TileHeight
    nObject = 1
    chkSet(0).Caption = "Style " & Format(nObject, "00")
    frmMain.stbar.Text = "Creating Style Boxes..."
        For tX = StartX To MaximumX Step TileWidth
            For tY = StartY To MaximumY Step TileHeight
                Load chkSet(nObject)
                chkSet(nObject).Move tX, tY
                chkSet(nObject).Visible = True
                If nObject >= UBound(cStyle) + 1 Then
                chkSet(nObject).Caption = ""
                Else
                chkSet(nObject).Caption = "Style " & Format(nObject + 1, "00")
                End If
                nObject = nObject + 1
             StartY = 0
         DoEvents
        Next
      DoEvents
    Next
frmMain.stbar.Text = "Done."
End Sub

Private Sub StyleScrollH_Change()
StyleScrollH_Scroll
End Sub

Private Sub StyleScrollH_Scroll()
picStyleBox.Left = -StyleScrollH.Value
picPrev.Picture = LoadPicture()
End Sub
