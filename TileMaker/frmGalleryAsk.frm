VERSION 5.00
Begin VB.Form frmGalleryAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tile Gallery"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Gallery"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.Timer GalleryTimer 
         Interval        =   30
         Left            =   1800
         Top             =   600
      End
      Begin VB.OptionButton optGallery 
         Caption         =   "Custom Gallery"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Loads the custom gallery"
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optGallery 
         Caption         =   "Default Gallery"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Loads the default gallery"
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGalleryAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmds_Click(Index As Integer)
If Index = 0 Then
Unload Me
frmMain.Gallerymnu_Click galMode
Else
Unload Me
End If
End Sub

Private Sub Form_Load()
galMode = 0
End Sub

Private Sub GalleryTimer_Timer()
optGallery(1).Enabled = (Total_CustomTiles > 0)
End Sub

Private Sub optGallery_Click(Index As Integer)
galMode = Index
End Sub

Private Sub optGallery_DblClick(Index As Integer)
cmds_Click 0
End Sub
