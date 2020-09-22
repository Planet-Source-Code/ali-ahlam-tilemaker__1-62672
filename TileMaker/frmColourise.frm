VERSION 5.00
Begin VB.Form frmColourise 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colourise..."
   ClientHeight    =   3255
   ClientLeft      =   5205
   ClientTop       =   4320
   ClientWidth     =   3915
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColourise.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picCustom 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   3735
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox picCombination 
         Height          =   615
         Left            =   2040
         ScaleHeight     =   555
         ScaleWidth      =   1590
         TabIndex        =   19
         ToolTipText     =   "Click to Preview"
         Top             =   390
         Width           =   1650
      End
      Begin VB.HScrollBar ColorBar 
         Height          =   200
         Index           =   2
         LargeChange     =   10
         Left            =   240
         Max             =   255
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.HScrollBar ColorBar 
         Height          =   200
         Index           =   1
         LargeChange     =   10
         Left            =   240
         Max             =   255
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.HScrollBar ColorBar 
         Height          =   200
         Index           =   0
         LargeChange     =   10
         Left            =   240
         Max             =   255
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select the color of the Lens:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   2025
      End
      Begin VB.Label lblPerc 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Index           =   2
         Left            =   1380
         TabIndex        =   22
         ToolTipText     =   "Channel Blue"
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lblPerc 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Index           =   1
         Left            =   1380
         TabIndex        =   21
         ToolTipText     =   "Channel Green"
         Top             =   600
         Width           =   435
      End
      Begin VB.Label lblPerc 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Index           =   0
         Left            =   1380
         TabIndex        =   20
         ToolTipText     =   "Channel Red"
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "B"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   18
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "G"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   17
         Top             =   600
         Width           =   105
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "R"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   16
         Top             =   360
         Width           =   105
      End
   End
   Begin VB.PictureBox picDefault 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   3735
      TabIndex        =   7
      Top             =   1560
      Width           =   3735
      Begin VB.PictureBox picSelected 
         Height          =   315
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   10
         Top             =   720
         Width           =   675
      End
      Begin VB.PictureBox picHue 
         AutoRedraw      =   -1  'True
         Height          =   315
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   255
         ScaleWidth      =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   3660
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         Caption         =   "Selected Color:"
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Select the color of the Lens:"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2025
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optChannel 
         Caption         =   "&Lens Effect"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optChannel 
         Caption         =   "C&olourise"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Timer previewTimer 
      Interval        =   10
      Left            =   3000
      Top             =   2520
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   4920
      MousePointer    =   99  'Custom
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   2520
      MousePointer    =   99  'Custom
      ScaleHeight     =   86
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   86
      TabIndex        =   2
      Top             =   120
      Width           =   1290
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1275
   End
End
Attribute VB_Name = "frmColourise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_cImage As New cImageProcessDIB
Attribute m_cImage.VB_VarHelpID = -1
Private m_cDib As New cDIBSection
Private m_cDibBuffer As New cDIBSection

Private m_fHue As Single

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
frmMain.picWork.Picture = picImage.Image
   Unload Me
End Sub

Private Sub ColorBar_Change(Index As Integer)
ColorBar_Scroll Index
End Sub

Private Sub ColorBar_Scroll(Index As Integer)
lblPerc(Index).Caption = Format((ColorBar(Index).Value * 100) / 255, "0") & "%"
picCombination.BackColor = RGB(ColorBar(0).Value, ColorBar(1).Value, ColorBar(2).Value)
End Sub

Private Sub Form_Load()
    picImage.Picture = frmMain.picWork.Image
    picImage.Width = frmMain.picWork.Width
    picImage.Height = frmMain.picWork.Height
    optChannel_Click 0
End Sub



Private Sub optChannel_Click(Index As Integer)
If Index = 0 Then 'default
RGBPallete picHue
picDefault.Visible = True
picCustom.Visible = False
Else
picDefault.Visible = False
picCustom.Visible = True
End If
End Sub

Private Sub picCombination_Click()
ColorLens_Image frmMain.picWork, picImage, ColorBar(0).Value, _
ColorBar(1).Value, ColorBar(2).Value
picImage.Picture = picImage.Image
End Sub

Private Sub picHue_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim r As Long, g As Long, b As Long
   If (x > 0) And (y > 0) And (x < picHue.ScaleWidth) And (y < picHue.ScaleHeight) Then
      m_fHue = ((x \ Screen.TwipsPerPixelX) - 40) / 40
      HLSToRGB m_fHue, 1, 0.5, r, g, b
      picSelected.BackColor = RGB(r, g, b)
   End If
End Sub

Private Sub picHue_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If (Button And vbLeftButton) = vbLeftButton Then
      picHue_MouseDown Button, Shift, x, y
   End If

End Sub

Public Sub Colourise(ByVal fHue As Single)
   ' Colourise takes hue (-1 to 5)
   m_cImage.Colourise m_cDib, fHue, 0.5
   Render
End Sub


Public Sub Render()
    picImage.Picture = LoadPicture()
    picImage.Width = m_cDib.Width * Screen.TwipsPerPixelX
    picImage.Height = m_cDib.Height * Screen.TwipsPerPixelY
    m_cDib.PaintPicture picImage.hDC
    picImage.Refresh
    picImage.Picture = picImage.Image
End Sub

Private Sub picHue_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   picImage.Picture = frmMain.picWork.Image
   Tmp_Save_LoadPrev
   Colourise m_fHue
End Sub

Private Sub previewTimer_Timer()
    Draw_Preview picImage, picPreview
End Sub

Private Function OpenFile(ByVal sFIle As String, Optional ByVal bIsTemp As Boolean = False) As Boolean
Dim sPicPrev As StdPicture
On Error GoTo OpenFileError
    
    Set sPicPrev = LoadPicture(sFIle)
    
    m_cDib.CreateFromPicture sPicPrev
    m_cDibBuffer.Create m_cDib.Width, m_cDib.Height
    
    OpenFile = True
    Exit Function
OpenFileError:
    MsgBox "An error occured trying to open this file: " & Err.Description, vbExclamation
    Exit Function
End Function

Private Sub Tmp_Save_LoadPrev()
Dim tmpPrevFile As String
tmpPrevFile = App.Path & "\tmp§Prev~"
On Error GoTo HandleTmpErr
'save to file
SavePicture picImage.Picture, tmpPrevFile
'update the class
OpenFile tmpPrevFile
'delete the tmp file
Kill tmpPrevFile
Exit Sub

HandleTmpErr: 'handle any ERRs
MsgBox Err.Description, vbCritical
Exit Sub
End Sub


Sub RGBPallete(pic As PictureBox)
Dim h As Single
Dim r As Long, g As Long, b As Long
Dim lH As Long
Dim x As Long

pic.Picture = LoadPicture()

   lH = pic.ScaleHeight
   For h = -40 To 200
      HLSToRGB h / 40, 1, 0.5, r, g, b
      pic.Line (x, 0)-(x + Screen.TwipsPerPixelX, lH), RGB(r, g, b), BF
      x = x + Screen.TwipsPerPixelX
   Next h
   pic.Refresh
   'picHue_MouseDown 1, 0, 0, 0
End Sub

