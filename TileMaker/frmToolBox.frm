VERSION 5.00
Begin VB.Form frmToolBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   600
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolBox 
      BorderStyle     =   0  'None
      Height          =   2500
      Left            =   0
      ScaleHeight     =   2505
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Width           =   600
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   10
         Left            =   300
         MouseIcon       =   "frmToolBox.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":0152
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   16
         ToolTipText     =   "Eraser"
         Top             =   1570
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   8
         Left            =   300
         MouseIcon       =   "frmToolBox.frx":01CE
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":0320
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   15
         ToolTipText     =   "Filled Chord"
         Top             =   1260
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   7
         Left            =   300
         MouseIcon       =   "frmToolBox.frx":069F
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":07F1
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   14
         ToolTipText     =   "Filled Circle/Ellipse"
         Top             =   950
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   6
         Left            =   300
         MouseIcon       =   "frmToolBox.frx":0B6C
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":0CBE
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   13
         ToolTipText     =   "Filled Box/Rectangle"
         Top             =   630
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   1
         Left            =   300
         MouseIcon       =   "frmToolBox.frx":1034
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":1186
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   12
         ToolTipText     =   "Line"
         Top             =   320
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   11
         Left            =   300
         MouseIcon       =   "frmToolBox.frx":11E4
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":1336
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   11
         ToolTipText     =   "Fill"
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   5
         Left            =   0
         MouseIcon       =   "frmToolBox.frx":13B5
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":1507
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   10
         ToolTipText     =   "Arc"
         Top             =   1570
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   4
         Left            =   0
         MouseIcon       =   "frmToolBox.frx":1565
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":16B7
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   9
         ToolTipText     =   "Chord"
         Top             =   1260
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   3
         Left            =   0
         MouseIcon       =   "frmToolBox.frx":1719
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":186B
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   8
         ToolTipText     =   "Circle/Ellipse"
         Top             =   950
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   2
         Left            =   0
         MouseIcon       =   "frmToolBox.frx":18D0
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":1A22
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   7
         ToolTipText     =   "Box/Rectangle"
         Top             =   630
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   9
         Left            =   0
         MouseIcon       =   "frmToolBox.frx":1A80
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":1BD2
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   6
         ToolTipText     =   "Text"
         Top             =   320
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmToolBox.frx":1C30
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":1D82
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   5
         ToolTipText     =   "Free Hand"
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   12
         Left            =   0
         MouseIcon       =   "frmToolBox.frx":1DFC
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":1F4E
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   4
         ToolTipText     =   "Selector"
         Top             =   1890
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   13
         Left            =   300
         MouseIcon       =   "frmToolBox.frx":1FB3
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":2105
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   3
         ToolTipText     =   "Color Picker"
         Top             =   1890
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   14
         Left            =   0
         MouseIcon       =   "frmToolBox.frx":2167
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":22B9
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   2
         ToolTipText     =   "Stroke"
         Top             =   2200
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Enabled         =   0   'False
         Height          =   300
         Index           =   15
         Left            =   300
         MouseIcon       =   "frmToolBox.frx":25FB
         MousePointer    =   99  'Custom
         Picture         =   "frmToolBox.frx":274D
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   1
         ToolTipText     =   "Stroke"
         Top             =   2200
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
picTool(cTool).Line (0, 0)-(15, 15), vbRed, B
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If
End Sub



Private Sub Form_Load()
Me.Move frmLens.Left + frmLens.Width, frmLens.Top
End Sub

Private Sub picTool_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 frmMain.picTool_Mouse_Down Index, Button, Shift, X, Y
 Unload Me
End Sub
