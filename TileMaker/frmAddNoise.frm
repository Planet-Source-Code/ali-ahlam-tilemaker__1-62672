VERSION 5.00
Begin VB.Form frmAddNoise 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Noise"
   ClientHeight    =   3240
   ClientLeft      =   4755
   ClientTop       =   2460
   ClientWidth     =   3930
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMosaic 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   3375
      TabIndex        =   22
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
      Begin VB.HScrollBar pixSize 
         Height          =   255
         Left            =   120
         Max             =   25
         Min             =   1
         TabIndex        =   24
         Top             =   360
         Value           =   1
         Width           =   3135
      End
      Begin VB.Label lblmos 
         AutoSize        =   -1  'True
         Caption         =   "Pixel Size:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblPixelVal 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   960
         TabIndex        =   23
         Top             =   120
         Width           =   90
      End
   End
   Begin VB.PictureBox picSample 
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
      TabIndex        =   4
      Top             =   240
      Width           =   1290
   End
   Begin VB.PictureBox picNoise 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3735
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   1440
         Width           =   1290
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   960
         TabIndex        =   18
         Text            =   "20"
         Top             =   1980
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "Colors"
         Height          =   1815
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1935
         Begin VB.OptionButton optType 
            Caption         =   "&Selection"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   540
            Width           =   1155
         End
         Begin VB.OptionButton optType 
            Caption         =   "&Random"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.HScrollBar RBar 
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            Max             =   128
            TabIndex        =   12
            Top             =   1410
            Width           =   1215
         End
         Begin VB.PictureBox RPic 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   11
            Top             =   1410
            Width           =   255
         End
         Begin VB.HScrollBar GBar 
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            Max             =   128
            TabIndex        =   10
            Top             =   1120
            Width           =   1215
         End
         Begin VB.PictureBox Gpic 
            BackColor       =   &H80000008&
            Height          =   255
            Left            =   1560
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   9
            Top             =   1120
            Width           =   255
         End
         Begin VB.HScrollBar BBar 
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            Max             =   128
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.PictureBox BPic 
            BackColor       =   &H80000008&
            Height          =   255
            Left            =   1560
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   7
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lbls 
            AutoSize        =   -1  'True
            Caption         =   "R:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   165
         End
         Begin VB.Label lbls 
            AutoSize        =   -1  'True
            Caption         =   "G:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   1120
            Width           =   165
         End
         Begin VB.Label lbls 
            AutoSize        =   -1  'True
            Caption         =   "B:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   1410
            Width           =   150
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1560
         TabIndex        =   20
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         Caption         =   "Percent:"
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.Timer PrevTimer 
      Interval        =   10
      Left            =   4800
      Top             =   1680
   End
   Begin VB.PictureBox picPrev 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   3960
      MousePointer    =   99  'Custom
      ScaleHeight     =   86
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   86
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame fraSep 
      Height          =   75
      Left            =   -300
      TabIndex        =   0
      Top             =   2520
      Width           =   4755
   End
End
Attribute VB_Name = "frmAddNoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bRandom As Boolean
Private m_lPercent As Long

Private m_cImage As New cImageProcessDIB
Attribute m_cImage.VB_VarHelpID = -1
Private m_cDib As New cDIBSection
Private m_cDibBuffer As New cDIBSection


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
        frmMain.picWork.Picture = picPrev.Image
    Unload Me
End Sub




Private Sub cmdPreview_Click()
    picPrev.Picture = frmMain.picWork.Image
    AddNoise m_bRandom, (CLng(txtAmount.Text))
End Sub

Private Sub Form_Load()
    picSample.Picture = picSample.Image
    picPrev.Width = frmMain.picWork.Width
    picPrev.Height = frmMain.picWork.Height
    picPrev.Picture = frmMain.picWork.Image
    If Me.Caption = "Mosaic" Then pixSize_Change
End Sub





Private Sub optType_Click(Index As Integer)
If optType(1).Value = True Then
    m_bRandom = True
Else
    m_bRandom = False
End If
    RBar.Enabled = m_bRandom
    GBar.Enabled = m_bRandom
    BBar.Enabled = m_bRandom
End Sub


Private Sub AddNoise(ByVal bRandom As Boolean, ByVal lAmount As Long)
Tmp_Save_LoadPrev

    m_cImage.AddNoise m_cDib, lAmount, bRandom
    RenderPrev
End Sub

Private Sub RenderPrev()
    'actual image preview
    picPrev.Width = m_cDib.Width * Screen.TwipsPerPixelX
    picPrev.Height = m_cDib.Height * Screen.TwipsPerPixelY
    m_cDib.PaintPicture picPrev.hDC
    picPrev.Refresh
    picPrev.Picture = picPrev.Image
End Sub

Private Sub Tmp_Save_LoadPrev()
Dim tmpPrevFile As String
tmpPrevFile = App.Path & "\tmp§Prev~"
On Error GoTo HandleTmpErr
'save to file
SavePicture picPrev.Picture, tmpPrevFile
'update the class
OpenFile tmpPrevFile
'delete the tmp file
Kill tmpPrevFile
Exit Sub

HandleTmpErr: 'handle any ERRs
MsgBox Err.Description, vbCritical
Exit Sub
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

Private Sub pixSize_Change()
pixSize_Scroll
End Sub

Private Sub pixSize_Scroll()
    picPrev.Picture = LoadPicture()
    Mosaic_Image frmMain.picWork, picPrev, pixSize.Value
    picPrev.Picture = picPrev.Image
    lblPixelVal.Caption = pixSize.Value
End Sub

Private Sub RBar_Change()
RBar_Scroll
End Sub

Private Sub RBar_Scroll()
RPic.BackColor = RGB(0, 0, RBar.Value)
    m_cImage.RSelection = RBar.Value
End Sub

Private Sub GBar_Change()
GBar_Scroll
End Sub

Private Sub GBar_Scroll()
Gpic.BackColor = RGB(0, GBar.Value, 0)
    m_cImage.GSelection = GBar.Value
End Sub

Private Sub BBar_Change()
BBar_Scroll
End Sub

Private Sub BBar_Scroll()
BPic.BackColor = RGB(BBar.Value, 0, 0)
    m_cImage.BSelection = BBar.Value
End Sub

Private Sub PrevTimer_Timer()
    Draw_Preview picPrev, picSample
    If Me.Caption = "Add Noise" Then
    picNoise.Visible = True
    picMosaic.Visible = False
    picSample.Move 2520, 240
    Else
    picNoise.Visible = False
    picMosaic.Visible = True
    picSample.Move 1320, 240
    End If
End Sub
