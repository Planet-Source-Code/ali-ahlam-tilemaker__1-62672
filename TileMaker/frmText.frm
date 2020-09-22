VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picCopy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   4320
      MouseIcon       =   "frmText.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shadow Offset"
      Height          =   1095
      Left            =   2160
      TabIndex        =   15
      Top             =   2400
      Width           =   1935
      Begin VB.TextBox ST 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "1"
         ToolTipText     =   "Shadow Offset - Top"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox SL 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "1"
         ToolTipText     =   "Shadow Offset - Left"
         Top             =   360
         Width           =   375
      End
      Begin ComCtl2.UpDown USL 
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         ToolTipText     =   "Shadow Offset - Left"
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "SL"
         BuddyDispid     =   196611
         OrigLeft        =   1575
         OrigTop         =   360
         OrigRight       =   1815
         OrigBottom      =   645
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UST 
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         ToolTipText     =   "Shadow Offset - Top"
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "ST"
         BuddyDispid     =   196610
         OrigLeft        =   1560
         OrigTop         =   720
         OrigRight       =   1800
         OrigBottom      =   1005
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Top:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Left:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   315
      End
   End
   Begin MSComDlg.CommonDialog CMDLG 
      Left            =   4920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&Font"
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   11
      ToolTipText     =   "Change Font"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame fram 
      Caption         =   "Colors"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1935
      Begin VB.CheckBox optShades 
         Caption         =   "Shadow"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Enable/Disable Shadow"
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox optShades 
         Caption         =   "Highlight"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Enable/Disable Highlight"
         Top             =   600
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.PictureBox picColors 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         DrawMode        =   6  'Mask Pen Not
         DrawStyle       =   2  'Dot
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1320
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   10
         ToolTipText     =   "Shadow Color"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox picColors 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   6  'Mask Pen Not
         DrawStyle       =   2  'Dot
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1320
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   9
         ToolTipText     =   "Highlight Color"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picColors 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         DrawMode        =   6  'Mask Pen Not
         DrawStyle       =   2  'Dot
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1320
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   8
         ToolTipText     =   "Fore Color"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ForeColor:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame fram 
      Caption         =   "Output"
      Height          =   1335
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1935
      Begin VB.PictureBox picSample 
         AutoRedraw      =   -1  'True
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   1635
         TabIndex        =   5
         ToolTipText     =   "Sample"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "OK"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtEntry 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter Text Here"
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Highlight Offset"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1935
      Begin VB.TextBox HT 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "1"
         ToolTipText     =   "Highlight Offset - Top"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox HL 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "1"
         ToolTipText     =   "Highlight Offset - Left"
         Top             =   360
         Width           =   375
      End
      Begin ComCtl2.UpDown UHL 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         ToolTipText     =   "Highlight Offset - Left"
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "HL"
         BuddyDispid     =   196624
         OrigLeft        =   1575
         OrigTop         =   360
         OrigRight       =   1815
         OrigBottom      =   645
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UHT 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         ToolTipText     =   "Highlight Offset - Top"
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "HT"
         BuddyDispid     =   196623
         OrigLeft        =   960
         OrigTop         =   720
         OrigRight       =   1200
         OrigBottom      =   1005
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Top:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Left:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Enter Text:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private curX As Single
Private curY As Single


Private Sub cmds_Click(Index As Integer)
Select Case Index
Case 0 'ok
Apply_Text picCopy, frmMain.cxPos, frmMain.cyPos
frmMain.picWork.Picture = picCopy.Image
frmMain.Save_UndoAction
firstTimetxt = False
Unload Me
Case 1 'cancel
Unload Me
Case 2 'font
On Error GoTo handleFontErr
With CMDLG
.Flags = cdlCFEffects Or _
cdlCFWYSIWYG Or cdlCFBoth Or _
cdlCFScalableOnly
.CancelError = True
.FontName = cFontName
.FontSize = cFontSize
.FontBold = cFontBold
.FontItalic = cFontItalic
.FontUnderline = cFontUnderline
.FontStrikethru = cFontStrikeThru
.Color = cFontColor
.ShowFont

cFontName = .FontName
cFontSize = .FontSize
cFontBold = .FontBold
cFontItalic = .FontItalic
cFontUnderline = .FontUnderline
cFontStrikeThru = .FontStrikethru
cFontColor = .Color
picColors(0).BackColor = cFontColor

With txtEntry
.FontName = cFontName
.FontSize = cFontSize
.FontBold = cFontBold
.FontItalic = cFontItalic
.FontUnderline = cFontUnderline
.FontStrikethru = cFontStrikeThru
.ForeColor = cFontColor
End With
End With
Change_Sample
End Select
Exit Sub

handleFontErr:
Exit Sub
End Sub

Private Sub Form_Load()
If firstTimetxt = True Then
cFontName = "ARIAL BLACK"
cFontSize = 12
cFontBold = False
cFontItalic = False
cFontUnderline = False
cFontStrikeThru = False
cFontColor = vbRed
End If
Display_Sample "ARIAL BLACK", 30, False, False, False, False

With picCopy
.Width = frmMain.picWork.Width
.Height = frmMain.picWork.Height
.Picture = frmMain.picWork.Picture
End With
End Sub


Private Sub optShades_Click(Index As Integer)
Change_Sample

UHL.Enabled = optShades(0).Value
UHT.Enabled = optShades(0).Value

USL.Enabled = optShades(1).Value
UST.Enabled = optShades(1).Value

End Sub

Private Sub picColors_DblClick(Index As Integer)
Dim SelColor As Long
On Error GoTo handleColErr
With CMDLG
.CancelError = True
.Flags = cdlCCRGBInit
.ShowColor
SelColor = .Color
End With
picColors(Index).BackColor = SelColor
txtEntry.ForeColor = SelColor
Change_Sample
Exit Sub

handleColErr:
Exit Sub
End Sub


Private Sub picColors_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picColors(Index).Line (0, 0)-(22, 14), vbBlue, B
End Sub

Private Sub picColors_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picColors(0).Cls
picColors(1).Cls
picColors(2).Cls
End Sub


Private Sub UHL_Change()
Change_Sample
End Sub

Private Sub UHT_Change()
Change_Sample
End Sub

Private Sub USL_Change()
Change_Sample
End Sub

Private Sub UST_Change()
Change_Sample
End Sub


Private Sub Display_Sample(ByVal fName As String, _
ByVal fSize As Integer, ByVal fBold As Boolean, _
ByVal fItalic As Boolean, ByVal fUline As Boolean, _
ByVal fStrike As Boolean)
picSample.Cls
picSample.FontName = fName
picSample.FontSize = fSize
picSample.FontBold = fBold
picSample.FontItalic = fItalic
picSample.FontUnderline = fUline
picSample.FontStrikethru = fStrike

picSample.CurrentX = 10
picSample.CurrentY = 0
Text3d picSample, picSample.CurrentX, picSample.CurrentY, "ABC", _
picColors(0).BackColor, picColors(1).BackColor, picColors(2).BackColor, _
UHL.Value, USL.Value, UHT.Value, UST.Value, _
optShades(0).Value, optShades(1).Value
End Sub

Private Sub Change_Sample()
Display_Sample "ARIAL BLACK", 30, _
False, False, False, _
False
End Sub

Private Sub Apply_Text(pic As PictureBox, ByVal X As Single, ByVal Y As Single)
With pic
.DrawMode = vbCopyPen
.FillStyle = frmMain.cFill_Style
.FontName = cFontName
.FontSize = cFontSize
.FontBold = cFontBold
.FontItalic = cFontItalic
.FontUnderline = cFontUnderline
.FontStrikethru = cFontStrikeThru
.ForeColor = cFontColor
End With
Text3d pic, X, (Y - 6), Trim$(txtEntry.Text), _
picColors(0).BackColor, picColors(1).BackColor, _
picColors(2).BackColor, UHL.Value, USL.Value, _
UHT.Value, UST.Value, optShades(0).Value, optShades(1).Value
End Sub

'call Text3d Picture1, txt, vbRed, vbBlack, vbWhite
Sub Text3d(pic As Object, ByVal X As Single, ByVal Y As Single, _
ByVal txt As String, ByVal fore As Long, _
ByVal highlight As Long, ByVal shadow As Long, _
ByVal AdjustHlightL As Integer, ByVal AdjustShadowL As Integer, _
ByVal AdjustHlightT As Integer, ByVal AdjustShadowT As Integer, _
Optional HLighted As Boolean = True, Optional Shaded As Boolean = True)

Dim oldSmode As Integer

Dim oldcolor As Long

    oldcolor = pic.ForeColor
    oldSmode = pic.ScaleMode
    pic.ScaleMode = vbPixels
 '   x = pic.CurrentX
 '   y = pic.CurrentY
        
    If HLighted = True Then
    'first shade
    pic.ForeColor = highlight
    pic.CurrentX = X - AdjustHlightL
    pic.CurrentY = Y - AdjustHlightT
    'pic.Print txt
    TextOut pic.hDC, pic.CurrentX, pic.CurrentY, txt, Len(txt)
    End If
    
    If Shaded = True Then
    'second shade
    pic.ForeColor = shadow
    pic.CurrentX = X + AdjustShadowL
    pic.CurrentY = Y + AdjustShadowT
    'pic.Print txt
    TextOut pic.hDC, pic.CurrentX, pic.CurrentY, txt, Len(txt)
    End If
    
    'forecolor
    pic.ForeColor = fore
    pic.CurrentX = X
    pic.CurrentY = Y
    'pic.Print txt
    TextOut pic.hDC, pic.CurrentX, pic.CurrentY, txt, Len(txt)

    pic.ForeColor = oldcolor
    pic.ScaleMode = oldSmode
End Sub



