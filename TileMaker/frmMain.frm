VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tile Maker"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7635
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ToolbarButtonBox 
      Height          =   2295
      Left            =   6840
      ScaleHeight     =   2235
      ScaleWidth      =   795
      TabIndex        =   151
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
      Begin VB.PictureBox tButtonB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   360
         Picture         =   "frmMain.frx":030A
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   163
         ToolTipText     =   "Help"
         Top             =   360
         Width           =   360
      End
      Begin VB.PictureBox tButtonB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   360
         Picture         =   "frmMain.frx":0968
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   162
         ToolTipText     =   "Help"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButtonB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   360
         Picture         =   "frmMain.frx":0FC6
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   161
         ToolTipText     =   "Help"
         Top             =   1800
         Width           =   360
      End
      Begin VB.PictureBox tButtonB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   360
         Picture         =   "frmMain.frx":1624
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   160
         ToolTipText     =   "Help"
         Top             =   1440
         Width           =   360
      End
      Begin VB.PictureBox tButtonB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   360
         Picture         =   "frmMain.frx":1C82
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   159
         ToolTipText     =   "Help"
         Top             =   1080
         Width           =   360
      End
      Begin VB.PictureBox tButtonB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   360
         Picture         =   "frmMain.frx":22E0
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   158
         ToolTipText     =   "Help"
         Top             =   720
         Width           =   360
      End
      Begin VB.PictureBox tButtonA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   0
         Picture         =   "frmMain.frx":293E
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   157
         ToolTipText     =   "Help"
         Top             =   360
         Width           =   360
      End
      Begin VB.PictureBox tButtonA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   0
         Picture         =   "frmMain.frx":2F9C
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   156
         ToolTipText     =   "Help"
         Top             =   1440
         Width           =   360
      End
      Begin VB.PictureBox tButtonA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   0
         Picture         =   "frmMain.frx":35FA
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   155
         ToolTipText     =   "Help"
         Top             =   1080
         Width           =   360
      End
      Begin VB.PictureBox tButtonA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   0
         Picture         =   "frmMain.frx":3C58
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   154
         ToolTipText     =   "Help"
         Top             =   720
         Width           =   360
      End
      Begin VB.PictureBox tButtonA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   0
         Picture         =   "frmMain.frx":42B6
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   153
         ToolTipText     =   "Help"
         Top             =   1800
         Width           =   360
      End
      Begin VB.PictureBox tButtonA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   0
         Picture         =   "frmMain.frx":4914
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   152
         ToolTipText     =   "Help"
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.Timer MainToolBarTimer 
      Interval        =   30
      Left            =   6840
      Top             =   120
   End
   Begin VB.PictureBox picToolBar 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   7365
      TabIndex        =   129
      Top             =   50
      Width           =   7365
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   5040
         Picture         =   "frmMain.frx":4F72
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   164
         ToolTipText     =   "Get from Gallery"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   4680
         Picture         =   "frmMain.frx":55D0
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   145
         ToolTipText     =   "Add to Gallery"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   5520
         Picture         =   "frmMain.frx":5C2E
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   144
         ToolTipText     =   "Pattern Settings"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   5880
         Picture         =   "frmMain.frx":628C
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   143
         ToolTipText     =   "Page Settings"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   3600
         Picture         =   "frmMain.frx":68EA
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   142
         ToolTipText     =   "View Zoomed Tile"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   3960
         Picture         =   "frmMain.frx":6F48
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   141
         ToolTipText     =   "Show Tiled Background"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   4320
         Picture         =   "frmMain.frx":75A6
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   140
         ToolTipText     =   "Load From Image..."
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   6960
         Picture         =   "frmMain.frx":7C04
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   139
         ToolTipText     =   "Help"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":8262
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   138
         ToolTipText     =   "New"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":88C0
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   137
         ToolTipText     =   "Open"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   720
         Picture         =   "frmMain.frx":8F1E
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   136
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   1080
         Picture         =   "frmMain.frx":957C
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   135
         ToolTipText     =   "Print"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   1560
         Picture         =   "frmMain.frx":9BDA
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   134
         ToolTipText     =   "Cut"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   1920
         Picture         =   "frmMain.frx":A238
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   133
         ToolTipText     =   "Copy"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   2280
         Picture         =   "frmMain.frx":A896
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   132
         ToolTipText     =   "Paste"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":AEF4
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   131
         ToolTipText     =   "Undo"
         Top             =   0
         Width           =   360
      End
      Begin VB.PictureBox tButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":B552
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   130
         ToolTipText     =   "Redo"
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox FillStyleBox 
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   120
      Top             =   2520
      Visible         =   0   'False
      Width           =   1500
      Begin VB.PictureBox FStyle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":BBB0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   126
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox FStyle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   240
         Picture         =   "frmMain.frx":BC10
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   125
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox FStyle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   480
         Picture         =   "frmMain.frx":BC6B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   124
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox FStyle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   720
         Picture         =   "frmMain.frx":BCD0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   123
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox FStyle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmMain.frx":BD33
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   122
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox FStyle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   1200
         Picture         =   "frmMain.frx":BD97
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   121
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox redoCont 
      Height          =   2535
      Left            =   -7005
      ScaleHeight     =   2475
      ScaleWidth      =   6795
      TabIndex        =   95
      Top             =   1920
      Width           =   6855
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   23
         Left            =   5880
         MouseIcon       =   "frmMain.frx":BE01
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   119
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   22
         Left            =   5040
         MouseIcon       =   "frmMain.frx":BF53
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   118
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   21
         Left            =   4200
         MouseIcon       =   "frmMain.frx":C0A5
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   117
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   20
         Left            =   3360
         MouseIcon       =   "frmMain.frx":C1F7
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   116
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   19
         Left            =   2520
         MouseIcon       =   "frmMain.frx":C349
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   115
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   18
         Left            =   1680
         MouseIcon       =   "frmMain.frx":C49B
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   114
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   17
         Left            =   840
         MouseIcon       =   "frmMain.frx":C5ED
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   113
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   16
         Left            =   0
         MouseIcon       =   "frmMain.frx":C73F
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   112
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   15
         Left            =   5880
         MouseIcon       =   "frmMain.frx":C891
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   111
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   14
         Left            =   5040
         MouseIcon       =   "frmMain.frx":C9E3
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   110
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   13
         Left            =   4200
         MouseIcon       =   "frmMain.frx":CB35
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   109
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   12
         Left            =   3360
         MouseIcon       =   "frmMain.frx":CC87
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   108
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   11
         Left            =   2520
         MouseIcon       =   "frmMain.frx":CDD9
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   107
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   10
         Left            =   1680
         MouseIcon       =   "frmMain.frx":CF2B
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   106
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   9
         Left            =   840
         MouseIcon       =   "frmMain.frx":D07D
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   105
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   8
         Left            =   0
         MouseIcon       =   "frmMain.frx":D1CF
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   104
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   7
         Left            =   5880
         MouseIcon       =   "frmMain.frx":D321
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   103
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   6
         Left            =   5040
         MouseIcon       =   "frmMain.frx":D473
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   102
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   5
         Left            =   4200
         MouseIcon       =   "frmMain.frx":D5C5
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   101
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   4
         Left            =   3360
         MouseIcon       =   "frmMain.frx":D717
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   100
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   3
         Left            =   2520
         MouseIcon       =   "frmMain.frx":D869
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   99
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   2
         Left            =   1680
         MouseIcon       =   "frmMain.frx":D9BB
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   98
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   1
         Left            =   840
         MouseIcon       =   "frmMain.frx":DB0D
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   97
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picRedo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmMain.frx":DC5F
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   96
         Top             =   0
         Width           =   810
      End
   End
   Begin VB.Timer mainTimer 
      Interval        =   10
      Left            =   5640
      Top             =   6120
   End
   Begin VB.Timer presetTimer 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6960
      Tag             =   "Auto Start"
      Top             =   6120
   End
   Begin VB.Timer ToolBoxTimer 
      Interval        =   500
      Left            =   6120
      Top             =   6120
   End
   Begin VB.PictureBox picColorBox 
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   2925
      TabIndex        =   1
      ToolTipText     =   "Color Pallete"
      Top             =   3240
      Width           =   2920
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   18
         Top             =   0
         Width           =   495
         Begin VB.PictureBox SColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   50
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   19
            ToolTipText     =   "Left Color"
            Top             =   50
            Width           =   255
         End
         Begin VB.PictureBox SColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   150
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   20
            ToolTipText     =   "Right Color"
            Top             =   150
            Width           =   255
         End
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   0
         Left            =   660
         MouseIcon       =   "frmMain.frx":DDB1
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   2
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   8
         Left            =   660
         MouseIcon       =   "frmMain.frx":DF03
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   10
         Top             =   240
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   9
         Left            =   950
         MouseIcon       =   "frmMain.frx":E055
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   11
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   10
         Left            =   1230
         MouseIcon       =   "frmMain.frx":E1A7
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   12
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   11
         Left            =   1520
         MouseIcon       =   "frmMain.frx":E2F9
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   12
         Left            =   1800
         MouseIcon       =   "frmMain.frx":E44B
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   14
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   13
         Left            =   2090
         MouseIcon       =   "frmMain.frx":E59D
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   15
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   14
         Left            =   2370
         MouseIcon       =   "frmMain.frx":E6EF
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   16
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   15
         Left            =   2650
         MouseIcon       =   "frmMain.frx":E841
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   17
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   1
         Left            =   950
         MouseIcon       =   "frmMain.frx":E993
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   3
         Top             =   240
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   2
         Left            =   1230
         MouseIcon       =   "frmMain.frx":EAE5
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   4
         Top             =   240
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   3
         Left            =   1520
         MouseIcon       =   "frmMain.frx":EC37
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   5
         Top             =   240
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   4
         Left            =   1800
         MouseIcon       =   "frmMain.frx":ED89
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   6
         Top             =   240
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   5
         Left            =   2090
         MouseIcon       =   "frmMain.frx":EEDB
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   7
         Top             =   240
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   6
         Left            =   2370
         MouseIcon       =   "frmMain.frx":F02D
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   8
         Top             =   240
         Width           =   270
      End
      Begin VB.PictureBox picCol 
         Height          =   255
         Index           =   7
         Left            =   2650
         MouseIcon       =   "frmMain.frx":F17F
         MousePointer    =   99  'Custom
         ScaleHeight     =   195
         ScaleWidth      =   210
         TabIndex        =   9
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.PictureBox undoCont 
      Height          =   2655
      Left            =   -7005
      ScaleHeight     =   2595
      ScaleWidth      =   6795
      TabIndex        =   67
      Top             =   4800
      Width           =   6855
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   23
         Left            =   5880
         MouseIcon       =   "frmMain.frx":F2D1
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   94
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   22
         Left            =   5040
         MouseIcon       =   "frmMain.frx":F423
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   93
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   21
         Left            =   4200
         MouseIcon       =   "frmMain.frx":F575
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   92
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   20
         Left            =   3360
         MouseIcon       =   "frmMain.frx":F6C7
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   91
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   19
         Left            =   2520
         MouseIcon       =   "frmMain.frx":F819
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   90
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   18
         Left            =   1680
         MouseIcon       =   "frmMain.frx":F96B
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   89
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   17
         Left            =   840
         MouseIcon       =   "frmMain.frx":FABD
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   88
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   16
         Left            =   0
         MouseIcon       =   "frmMain.frx":FC0F
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   87
         Top             =   1680
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   15
         Left            =   5880
         MouseIcon       =   "frmMain.frx":FD61
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   86
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   14
         Left            =   5040
         MouseIcon       =   "frmMain.frx":FEB3
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   85
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   13
         Left            =   4200
         MouseIcon       =   "frmMain.frx":10005
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   84
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   12
         Left            =   3360
         MouseIcon       =   "frmMain.frx":10157
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   83
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   11
         Left            =   2520
         MouseIcon       =   "frmMain.frx":102A9
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   82
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   10
         Left            =   1680
         MouseIcon       =   "frmMain.frx":103FB
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   81
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   9
         Left            =   840
         MouseIcon       =   "frmMain.frx":1054D
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   80
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   8
         Left            =   0
         MouseIcon       =   "frmMain.frx":1069F
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   79
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   7
         Left            =   5880
         MouseIcon       =   "frmMain.frx":107F1
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   78
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   6
         Left            =   5040
         MouseIcon       =   "frmMain.frx":10943
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   77
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   5
         Left            =   4200
         MouseIcon       =   "frmMain.frx":10A95
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   76
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   4
         Left            =   3360
         MouseIcon       =   "frmMain.frx":10BE7
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   72
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   3
         Left            =   2520
         MouseIcon       =   "frmMain.frx":10D39
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   71
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   2
         Left            =   1680
         MouseIcon       =   "frmMain.frx":10E8B
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   70
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmMain.frx":10FDD
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   69
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   1
         Left            =   840
         MouseIcon       =   "frmMain.frx":1112F
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   68
         Top             =   0
         Width           =   810
      End
   End
   Begin VB.PictureBox statContainer 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7635
      TabIndex        =   60
      Top             =   5355
      Width           =   7635
      Begin VB.PictureBox ProgBar 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   80
         ScaleHeight     =   240
         ScaleWidth      =   5205
         TabIndex        =   128
         Top             =   30
         Visible         =   0   'False
         Width           =   5205
      End
      Begin VB.TextBox timBar 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   6500
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Stat"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox datBar 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   5350
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Stat"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox stbar 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   50
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Stat"
         ToolTipText     =   "Status Bar"
         Top             =   0
         Width           =   5250
      End
   End
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   65
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picToolBox 
      BorderStyle     =   0  'None
      Height          =   2500
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   600
      TabIndex        =   49
      Top             =   600
      Width           =   600
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Enabled         =   0   'False
         Height          =   300
         Index           =   15
         Left            =   300
         MouseIcon       =   "frmMain.frx":11281
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":113D3
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   150
         ToolTipText     =   "Stroke"
         Top             =   2200
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   14
         Left            =   0
         MouseIcon       =   "frmMain.frx":11427
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":11579
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   147
         ToolTipText     =   "Stroke"
         Top             =   2200
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   13
         Left            =   300
         MouseIcon       =   "frmMain.frx":118BB
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":11A0D
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   74
         ToolTipText     =   "Color Picker"
         Top             =   1890
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   12
         Left            =   0
         MouseIcon       =   "frmMain.frx":11A6F
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":11BC1
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   73
         ToolTipText     =   "Selector"
         Top             =   1890
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmMain.frx":11C26
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":11D78
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   50
         ToolTipText     =   "Free Hand"
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   9
         Left            =   0
         MouseIcon       =   "frmMain.frx":11DF2
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":11F44
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   59
         ToolTipText     =   "Text"
         Top             =   320
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   2
         Left            =   0
         MouseIcon       =   "frmMain.frx":11FA2
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":120F4
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   52
         ToolTipText     =   "Box/Rectangle"
         Top             =   630
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   3
         Left            =   0
         MouseIcon       =   "frmMain.frx":12152
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":122A4
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   53
         ToolTipText     =   "Circle/Ellipse"
         Top             =   950
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   4
         Left            =   0
         MouseIcon       =   "frmMain.frx":12309
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1245B
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   54
         ToolTipText     =   "Chord"
         Top             =   1260
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   5
         Left            =   0
         MouseIcon       =   "frmMain.frx":124BD
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1260F
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   55
         ToolTipText     =   "Arc"
         Top             =   1570
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   11
         Left            =   300
         MouseIcon       =   "frmMain.frx":1266D
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":127BF
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   66
         ToolTipText     =   "Fill"
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   1
         Left            =   300
         MouseIcon       =   "frmMain.frx":1283E
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":12990
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   51
         ToolTipText     =   "Line"
         Top             =   320
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   6
         Left            =   300
         MouseIcon       =   "frmMain.frx":129EE
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":12B40
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   56
         ToolTipText     =   "Filled Box/Rectangle"
         Top             =   630
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   7
         Left            =   300
         MouseIcon       =   "frmMain.frx":12EB6
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":13008
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   57
         ToolTipText     =   "Filled Circle/Ellipse"
         Top             =   950
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   8
         Left            =   300
         MouseIcon       =   "frmMain.frx":13383
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":134D5
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   58
         ToolTipText     =   "Filled Chord"
         Top             =   1260
         Width           =   300
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         DrawMode        =   6  'Mask Pen Not
         Height          =   300
         Index           =   10
         Left            =   300
         MouseIcon       =   "frmMain.frx":13854
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":139A6
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   64
         ToolTipText     =   "Eraser"
         Top             =   1570
         Width           =   300
      End
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      Height          =   3270
      Left            =   3300
      ScaleHeight     =   214
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   0
      ToolTipText     =   "Preview Area"
      Top             =   555
      Width           =   4215
      Begin VB.Label lblSample 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Text"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   600
         TabIndex        =   35
         Top             =   120
         Width           =   2985
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   3300
      TabIndex        =   32
      Top             =   3840
      Width           =   4215
      Begin VB.CommandButton cmdBtns 
         Caption         =   "&Zoom  [F2]"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   146
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Show Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   48
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtSample 
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   38
         Text            =   "Sample Text"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdBtns 
         Caption         =   "&Color"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox ChkPrev 
         Caption         =   "Preview Always"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Always show preview"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdBtns 
         Caption         =   "&Preview"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   33
         ToolTipText     =   "Preview Tile"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sample Text:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1170
      End
   End
   Begin MSComDlg.CommonDialog CMDLG 
      Left            =   360
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.PictureBox PicBack 
      Height          =   2580
      Left            =   795
      MouseIcon       =   "frmMain.frx":13A22
      MousePointer    =   99  'Custom
      ScaleHeight     =   2520
      ScaleWidth      =   2400
      TabIndex        =   22
      ToolTipText     =   "Work Area"
      Top             =   555
      Width           =   2460
      Begin VB.PictureBox picFillDisplay 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   600
         TabIndex        =   148
         Top             =   2280
         Width           =   600
         Begin VB.PictureBox CurFillStyle 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            DrawMode        =   6  'Mask Pen Not
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   149
            ToolTipText     =   "Current Fill Style"
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.PictureBox picFilterApply 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   2400
         MouseIcon       =   "frmMain.frx":13B74
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   127
         Top             =   0
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.PictureBox picClipBoard 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   -1300
         MouseIcon       =   "frmMain.frx":13CC6
         MousePointer    =   99  'Custom
         ScaleHeight     =   810
         ScaleWidth      =   810
         TabIndex        =   75
         Top             =   120
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.PictureBox picCord 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   2415
         TabIndex        =   45
         Top             =   2310
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Label lblPsiz 
            AutoSize        =   -1  'True
            Caption         =   "Pagesize"
            Height          =   195
            Left            =   1300
            TabIndex        =   47
            Top             =   0
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label lblCoords 
            AutoSize        =   -1  'True
            Caption         =   "X:0, Y:0"
            Height          =   195
            Left            =   25
            TabIndex        =   46
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   -1500
         MouseIcon       =   "frmMain.frx":13E18
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.PictureBox picWork 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   45
         MouseIcon       =   "frmMain.frx":13F6A
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   23
         Top             =   45
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   4995
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Image curTool 
         Height          =   480
         Index           =   6
         Left            =   3360
         Picture         =   "frmMain.frx":140BC
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image curTool 
         Height          =   480
         Index           =   5
         Left            =   2760
         Picture         =   "frmMain.frx":1420E
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image curTool 
         Height          =   480
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":14360
         Top             =   0
         Width           =   480
      End
      Begin VB.Image selector 
         Height          =   480
         Left            =   3840
         Picture         =   "frmMain.frx":144B2
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Cpicker 
         Height          =   480
         Left            =   4320
         Picture         =   "frmMain.frx":14604
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image curTool 
         Height          =   480
         Index           =   3
         Left            =   1440
         Picture         =   "frmMain.frx":14756
         Top             =   0
         Width           =   480
      End
      Begin VB.Image curTool 
         Height          =   480
         Index           =   1
         Left            =   480
         Picture         =   "frmMain.frx":148A8
         Top             =   0
         Width           =   480
      End
      Begin VB.Image curTool 
         Height          =   480
         Index           =   4
         Left            =   2280
         Picture         =   "frmMain.frx":149FA
         Top             =   0
         Width           =   480
      End
      Begin VB.Image curTool 
         Height          =   480
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":14B4C
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   24
      Top             =   3840
      Width           =   3120
      Begin VB.ComboBox preSets 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox LSize 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "1"
         ToolTipText     =   "Pen Thickness"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox nPat 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   790
         Width           =   495
      End
      Begin VB.CheckBox chkClear 
         Caption         =   "Clear Page"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Clear Picture First"
         Top             =   1080
         Width           =   1095
      End
      Begin ComCtl2.UpDown Pat1 
         Height          =   285
         Left            =   600
         TabIndex        =   30
         ToolTipText     =   "Patterns#2"
         Top             =   790
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   2
         BuddyControl    =   "nPat"
         BuddyDispid     =   196662
         OrigLeft        =   600
         OrigTop         =   960
         OrigRight       =   840
         OrigBottom      =   1245
         Increment       =   2
         Max             =   40
         Min             =   2
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdFlakes 
         Caption         =   "Flakes"
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         ToolTipText     =   "Flakes"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox nSteps 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin ComCtl2.UpDown Pat 
         Height          =   285
         Left            =   600
         TabIndex        =   25
         ToolTipText     =   "Patterns#1"
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "nSteps"
         BuddyDispid     =   196666
         OrigLeft        =   960
         OrigTop         =   4680
         OrigRight       =   1200
         OrigBottom      =   5535
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown LinSiz 
         Height          =   285
         Left            =   1560
         TabIndex        =   41
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "LSize"
         BuddyDispid     =   196661
         OrigLeft        =   2640
         OrigTop         =   600
         OrigRight       =   2880
         OrigBottom      =   885
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Presets:"
         Height          =   195
         Left            =   1920
         TabIndex        =   44
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pen Size:"
         Height          =   195
         Left            =   1080
         TabIndex        =   42
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Pat&terns:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Image NotAllowed 
      Height          =   480
      Left            =   4680
      Picture         =   "frmMain.frx":14C9E
      Top             =   5400
      Width           =   480
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save As"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print"
         Index           =   6
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   8
      End
   End
   Begin VB.Menu mnuE 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "U&ndo"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "C&ut"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuV 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Zoomed Tile"
         Index           =   0
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
         Index           =   1
      End
   End
   Begin VB.Menu mnuBr 
      Caption         =   "&Brush"
      Begin VB.Menu mnuBrush 
         Caption         =   "&Normal"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "&Circle"
         Index           =   1
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "&Box"
         Index           =   2
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "&Rectangle"
         Index           =   3
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "Filled B&ox"
         Index           =   4
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "Filled R&ectangle"
         Index           =   5
      End
      Begin VB.Menu spac 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrushSet 
         Caption         =   "&Settings..."
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuTl 
      Caption         =   "&Tool"
      Begin VB.Menu mnuTool 
         Caption         =   "&Pencil"
         Index           =   0
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Fill"
         Index           =   1
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Text"
         Index           =   2
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Line"
         Index           =   3
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Box Transparent"
         Index           =   4
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Box Filled"
         Index           =   5
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Circle Transparent"
         Index           =   6
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Circle Filled"
         Index           =   7
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Chord Transparent"
         Index           =   8
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Chord Filled"
         Index           =   9
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Arc"
         Index           =   10
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Eraser"
         Index           =   11
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Selector"
         Index           =   12
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Color Pic&ker"
         Index           =   13
      End
      Begin VB.Menu mnuTool 
         Caption         =   "St&roke"
         Index           =   14
      End
   End
   Begin VB.Menu mnuO 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Page Settings..."
         Index           =   0
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "P&attern Settings..."
         Index           =   2
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Tile Gallery"
         Index           =   4
         Begin VB.Menu mnuGallery 
            Caption         =   "&Default Gallery"
            Index           =   0
         End
         Begin VB.Menu mnuGallery 
            Caption         =   "&Custom Gallery"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "P&references..."
         Index           =   6
      End
   End
   Begin VB.Menu mnuFilt 
      Caption         =   "F&ilters"
      Begin VB.Menu mnuFilter 
         Caption         =   "S&often"
         Index           =   0
         Begin VB.Menu mnuSoftenFilter 
            Caption         =   "&Soften"
            Index           =   0
         End
         Begin VB.Menu mnuSoftenFilter 
            Caption         =   "Soften &More"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "S&harpen"
         Index           =   1
         Begin VB.Menu mnuSharpenFilter 
            Caption         =   "&Sharpen"
            Index           =   0
         End
         Begin VB.Menu mnuSharpenFilter 
            Caption         =   "Sharpen &More"
            Index           =   1
         End
         Begin VB.Menu mnuSharpenFilter 
            Caption         =   "&Unsharpen"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "B&lur"
         Index           =   2
         Begin VB.Menu mnuBlurFilter 
            Caption         =   "&Blur"
            Index           =   0
         End
         Begin VB.Menu mnuBlurFilter 
            Caption         =   "Blur &More"
            Index           =   1
         End
         Begin VB.Menu mnuBlurFilter 
            Caption         =   "M&otion Blur"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "&Artistic"
         Index           =   3
         Begin VB.Menu mnuArtisticFilter 
            Caption         =   "E&mboss"
            Index           =   0
         End
         Begin VB.Menu mnuArtisticFilter 
            Caption         =   "E&ngrave"
            Index           =   1
         End
         Begin VB.Menu mnuArtisticFilter 
            Caption         =   "N&oise..."
            Index           =   2
         End
         Begin VB.Menu mnuArtisticFilter 
            Caption         =   "Mo&saic..."
            Index           =   3
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "&Color"
         Index           =   4
         Begin VB.Menu mnuColor 
            Caption         =   "&Darken"
            Index           =   0
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Lighten"
            Index           =   1
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Negative"
            Index           =   2
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Grayscale"
            Index           =   3
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Black and White"
            Index           =   4
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Colourise..."
            Index           =   5
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Replace..."
            Index           =   6
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "S&pecial"
         Index           =   5
         Begin VB.Menu mnuSpecialFilter 
            Caption         =   "&Cloth"
            Index           =   0
         End
         Begin VB.Menu mnuSpecialFilter 
            Caption         =   "&Disabled"
            Index           =   1
         End
         Begin VB.Menu mnuSpecialFilter 
            Caption         =   "&Wave"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuImge 
      Caption         =   "&Image"
      Begin VB.Menu mnuImage 
         Caption         =   "Flip &Vertical"
         Index           =   0
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuImage 
         Caption         =   "Flip &Horizontal"
         Index           =   1
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuImage 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuImage 
         Caption         =   "&Rotate"
         Index           =   3
         Begin VB.Menu mnuRotate 
            Caption         =   "&Clockwise"
            Index           =   0
         End
         Begin VB.Menu mnuRotate 
            Caption         =   "&AntiClockwise"
            Index           =   1
         End
      End
      Begin VB.Menu mnuImage 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuImage 
         Caption         =   "&Mirror"
         Index           =   5
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuImage 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuImage 
         Caption         =   "&Offset..."
         Index           =   7
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuImage 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuImage 
         Caption         =   "R&eflection"
         Index           =   9
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuImage 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuImage 
         Caption         =   "Bac&k Color"
         Index           =   11
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuImage 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuImage 
         Caption         =   "&Tiled Background"
         Index           =   13
         Shortcut        =   {F2}
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadfrom 
         Caption         =   "&Load From..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddtoCGallery 
         Caption         =   "&Add To Gallery"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Contents"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Index"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Search for Help on..."
         Index           =   2
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   4
      End
   End
   Begin VB.Menu mnuFnt 
      Caption         =   "&Fonts"
      Visible         =   0   'False
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IsDown As Boolean
Public stopflag As Boolean
Public cFill_Style As Integer 'fill style

Private strokeCol As Long

Dim FirstX As Single
Dim FirstY As Single
Dim SecondX As Single
Dim SecondY As Single
Dim wD As Single 'aspect x
Dim HT As Single 'aspect y

Public cxPos As Single
Public cyPos As Single


Public UNDOS As Integer
Public REDOS As Integer

Public MoveSelection As Boolean

'******************************
Public m_sFIleName As String
Public m_sFIleTitle As String
Private m_bDirty As Boolean

Private WithEvents m_cImage As cImageProcessDIB
Attribute m_cImage.VB_VarHelpID = -1
Private m_cDib As New cDIBSection
Private m_cDibBuffer As New cDIBSection


Sub Setup_Envir()
On Error Resume Next
stbar.Text = "Setting Pallete..."
frmSplash.lblstat.Caption = "Setting Pallete..."
SetupPallete 'color pallete
stbar.Text = "Setting Styles..."
frmSplash.lblstat.Caption = "Setting Styles..."
SetupStyles 'style
stbar.Text = "Setting Presets..."
frmSplash.lblstat.Caption = "Setting Presets..."
SetupPresets 'presets
stbar.Text = "Setting Page..."
frmSplash.lblstat.Caption = "Setting Page..."
SetupPage 'page setup
stbar.Text = "Loading Environment Settings..."
frmSplash.lblstat.Caption = "Loading Environment Settings..."
mnuBrush_Click (Val(GetIniVal("General", "cBrush", IniName))) 'brush type

cTool = 0 'tool

TotColrs = 10 'total colors in pattern
pColMode = Val(GetIniVal("General", "ColorMode", IniName)) 'current style color mode to combined

If pColMode > 1 Then
pColMode = 1
End If

BcRad = Val(GetIniVal("Brush", "BcRad", IniName)) 'radius
BbSiz = Val(GetIniVal("Brush", "BbSiz", IniName)) 'box
BrSiz = Val(GetIniVal("Brush", "BrSiz", IniName)) 'rect
LinSiz.Value = Val(GetIniVal("General", "DrawSize", IniName))  'pen size
ChkPrev.Value = Val(GetIniVal("General", "AlwaysPrev", IniName))  'Always preview
chkShow.Value = Val(GetIniVal("General", "ShowText", IniName))  'Show text
SColor(0).BackColor = Val(GetIniVal("Colors", "Left", IniName))
SColor(1).BackColor = Val(GetIniVal("Colors", "Right", IniName))

'animation
Anim = Val(GetIniVal("General", "Animation", IniName))
If Anim < 0 Or Anim > 1 Then
Anim = 1
End If

'CustomGallery
Data_file = GetIniVal("Gallery", "CustomGallery", IniName)
If Dir(Data_file) = "" Then
Data_file = App.Path & "\CGallery.DAT" 'data file
End If

'UndoRedo
maxLevel = Val(GetIniVal("General", "MaxLevelUR", IniName))

If maxLevel >= 23 Then maxLevel = 23
If maxLevel <= 0 Then maxLevel = 5

Max_Undo = maxLevel
Max_Redo = maxLevel

UNDOS = 1 'undo actions
REDOS = 1 'undo actions

firstTimetxt = True 'text dialog first time load

tmpFName = App.Path & "\Tile¤tmp¯File" 'for tmp saving

stbar.Text = "Done."
Exit Sub
End Sub


Sub SetupPallete()
Dim colI
 For colI = 0 To 15
  picCol(colI).BackColor = QBColor(colI)
   DoEvents
 Next colI
End Sub

Sub SetupStyles()
Dim styleI
'first initialize
 For styleI = 0 To UBound(cStyle)
  cStyle(styleI) = 0
  DoEvents
 Next styleI
'apply some
 For styleI = 0 To 11
  cStyle(styleI) = 1
  DoEvents
 Next styleI
End Sub

Sub SetupPresets()
Dim PreI
For PreI = 1 To 75
preSets.AddItem "#" & PreI
DoEvents
Next PreI
End Sub

Sub SetupPage()
Dim PWid, PHgt As Single

PWid = Val((GetIniVal("PageSettings", "Width", IniName)))
PHgt = Val((GetIniVal("PageSettings", "Height", IniName)))

If PWid > 150 Then PWid = 150
If PHgt > 150 Then PHgt = 150

With picWork
.Width = PWid * 15
.Height = PHgt * 15
.FillStyle = 1
.DrawMode = vbCopyPen
End With

SetSizeInInches picPrint, 8.5, 15 'set size
'picPrint.Scale (-1, -1.5)-(7.5, 13) 'set scale

End Sub

Private Sub chkClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Clears Page Always for Patterns #2"
End Sub

Private Sub ChkPrev_Click()
If Val(GetIniVal("General", "SaveSettings", IniName)) = 1 Then
SetIniVal "General", "AlwaysPrev", ChkPrev.Value, IniName
End If
End Sub

Private Sub ChkPrev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Preview the changes Always"
End Sub

Private Sub chkShow_Click()
If Val(GetIniVal("General", "SaveSettings", IniName)) = 1 Then
SetIniVal "General", "ShowText", chkShow.Value, IniName
End If

End Sub

Private Sub cmdBtns_Click(Index As Integer)
Select Case Index
Case 0
BmpTile picView, picWork
Case 1
ApplyColor lblSample, "f"
Case 2
imageMnu_Click (13)
End Select
End Sub

Private Sub cmdBtns_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
stbar.Text = "Preview Tiled"
Case 1
stbar.Text = "Change Color of Sample Text"
Case 2
stbar.Text = "View Zoomed Background"
End Select
End Sub

Private Sub cmdFlakes_Click()
Dim cX, cY As Single
If cmdFlakes.Caption = "Flakes" Then
stopflag = False
cmdFlakes.Caption = "Stop"
Else
stopflag = True
picWork.Refresh
picWork.Picture = picWork.Image
picTmp.Picture = picWork.Image
picTmp.Refresh
BmpTile picView, picWork
cmdFlakes.Caption = "Flakes"
End If

Do While stopflag = False
Randomize
cX = picWork.Width * Rnd
cY = picWork.Height * Rnd
picWork.PSet (cX, cY), QBColor(15 * Rnd)
DoEvents
Loop
Save_UndoAction
End Sub


Private Sub cmdFlakes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Start/Stop Flakes"
End Sub







Private Sub Form_Load()
On Error Resume Next
If Dir(IniName) = "" Then
frmSplash.lblstat.Caption = "Creating INI File..."
ReCreateSettings
frmSplash.lblstat.Caption = ""
End If
mnuFile(8).Caption = "E&xit" & vbTab & "Alt+F4"

Setup_Envir 'set environment variables...

mnuFile_Click (0)

Paste_Status 'paste check
frmSplash.Animator.Enabled = False 'splash animator
frmSplash.lblstat.Caption = "Done."
'****************************
Set m_cImage = New cImageProcessDIB
ProgBar.Scale (0, 0)-(100, 10) 'for progress
MakeIt3D picToolBar, 1, 1, 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = ""
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = ""
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = ""
End Sub



Private Sub FStyle_Click(Index As Integer)
Dim ButI
For ButI = 0 To 5
    FStyle(ButI).Cls
Next ButI
FStyle(Index).Line (-1, -1)-(240, 240), vbBlack, BF
cFill_Style = Index + 2
FillStyleBox.Visible = False
CurFillStyle.Picture = FStyle(Index).Image
End Sub

Private Sub lblSample_Click()
PopupMenu mnuFnt
End Sub

Private Sub lblSample_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Sample Text"
End Sub

Private Sub LinSiz_Change()
DrawSize = LSize.Text
picWork.DrawWidth = DrawSize
If Val(GetIniVal("General", "SaveSettings", IniName)) = 1 Then
SetIniVal "General", "DrawSize", DrawSize, IniName
End If


    FirstX = 0
    FirstY = 0
    SecondX = 0
    SecondY = 0
End Sub

Private Sub LinSiz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Change Pen size"
End Sub

Private Sub MainTimer_Timer()
If Dir(IniName) = "" Then 'keep track of INI
ReCreateSettings
End If

mnuGallery(1).Enabled = (Total_CustomTiles > 0)

picWork.Refresh
picWork.Picture = picWork.Image
picTmp.Picture = picWork.Image

picTmp.Width = picWork.Width 'tmp picture
picTmp.Height = picWork.Height

MakeIt3D picColorBox, 1, 1, True
MakeIt3D picToolBox, 1, 1, True


lblSample.Visible = chkShow.Value <> 0
lblSample.Caption = txtSample.Text
lblPsiz.Caption = ((picWork.Width) / 15) & " X " & ((picWork.Height) / 15)
datBar.Text = Date
timBar.Text = Time

If Screen.MousePointer = 11 Then
Me.Enabled = False
Else
Me.Enabled = True
End If


If UNDOS < 2 Then
mnuEdit(0).Caption = "Can't Undo"
mnuEdit(0).Enabled = False
End If

If REDOS < 2 Then
mnuEdit(1).Caption = "Can't Redo"
mnuEdit(1).Enabled = False
End If

End Sub



Private Sub MainToolBarTimer_Timer()
'enable buttons accordingly
tButton(4).Enabled = (mnuEdit(3).Enabled) 'cut
tButton(5).Enabled = (mnuEdit(4).Enabled) 'copy
tButton(6).Enabled = (mnuEdit(5).Enabled) 'paste
tButton(7).Enabled = (mnuEdit(0).Enabled) 'undo
tButton(8).Enabled = (mnuEdit(1).Enabled) 'redo
tButton(12).Enabled = (mnuAddtoCGallery.Enabled) 'add to gallery

If tButton(4).Enabled = True Then
tButton(4).Picture = tButtonA(4).Picture
Else
tButton(4).Picture = tButtonB(4).Picture
End If

If tButton(5).Enabled = True Then
tButton(5).Picture = tButtonA(5).Picture
Else
tButton(5).Picture = tButtonB(5).Picture
End If

If tButton(6).Enabled = True Then
tButton(6).Picture = tButtonA(6).Picture
Else
tButton(6).Picture = tButtonB(6).Picture
End If

If tButton(7).Enabled = True Then
tButton(7).Picture = tButtonA(7).Picture
Else
tButton(7).Picture = tButtonB(7).Picture
End If

If tButton(8).Enabled = True Then
tButton(8).Picture = tButtonA(8).Picture
Else
tButton(8).Picture = tButtonB(8).Picture
End If

If tButton(12).Enabled = True Then
tButton(12).Picture = tButtonA(12).Picture
Else
tButton(12).Picture = tButtonB(12).Picture
End If
End Sub

Private Sub mnuAddtoCGallery_Click()
add_to_Gallery_Click
End Sub

Public Sub add_to_Gallery_Click()
'max entry reached?
If Total_CustomTiles >= 45 Then
    If MsgBox("Current Custom Gallery has reached the Maximum Limit" & vbCrLf & _
    "You should create a new gallery when it reaches the Maximum limit." & vbCrLf & _
    "Would you like to create it now?", vbYesNo + vbDefaultButton1 + vbCritical) = vbYes Then
    Create_New_Gallery
    End If
Else
    'add to gallery
    Gallery_Add_Now
End If
End Sub

Sub Create_New_Gallery()
Dim Custom_FName As String
Custom_FName = ""
Custom_FName = UCase(Trim$(InputBox("Enter the Custom Gallery Name" & vbCrLf & "Example:   mygallery", "New Custom Gallery")))

If Custom_FName = "" Then Exit Sub
Custom_FName = App.Path & "\" & Custom_FName & ".DAT"

Open Custom_FName For Output As #1
Close #1
Data_file = Custom_FName
SetIniVal "Gallery", "CustomGallery", Data_file, IniName
Gallery_Add_Now
End Sub

Sub Gallery_Add_Now()
Dim Tnumber
Dim tmpPic As String

If MsgBox("Would you like to add the current image to the custom Gallery?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    If picWork.ScaleWidth > 54 Or picWork.ScaleHeight > 54 Then
        If MsgBox("Current tile is larger than the allowed size" & vbCrLf & _
        "If you add it to the gallery part of it would be lost" & vbCrLf & _
        "Continue anyway?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Screen.MousePointer = 11
     
            tmpPic = App.Path & "\~tSPic.bmp"
    
            SavePicture picWork.Picture, tmpPic 'save tmp pic

            Tnumber = (Total_CustomTiles) + 1
    
            PicDB_Add Data_file, tmpPic, ("Tile#" & Tnumber) 'save dat

            Kill tmpPic 'delete tmp
 
            Screen.MousePointer = 0
            Else
                MsgBox "Tile Not Added."
        End If
    Else
            Screen.MousePointer = 11
     
            tmpPic = App.Path & "\~tSPic.bmp"
    
            SavePicture picWork.Picture, tmpPic 'save tmp pic

            Tnumber = (Total_CustomTiles) + 1
    
            PicDB_Add Data_file, tmpPic, ("Tile#" & Tnumber) 'save dat

            Kill tmpPic 'delete tmp
 
            Screen.MousePointer = 0

    End If
End If
End Sub

Private Sub mnuBlurFilter_Click(Index As Integer)
Tmp_SaveLoad 'save to tmp file
Me.Refresh
Select Case Index
Case 0 'blur
ProcessImage eBlur
Case 1 'blur more
ProcessImage eBlurMore
Case 2
Screen.MousePointer = 11
    ProgBar.Visible = True
    ProgBar.Cls
    Process_Image picWork, picFilterApply, fMotionBlur
    picWork.Picture = picFilterApply.Image
    picFilterApply.Picture = LoadPicture()
    picFilterApply.Width = 50
    picFilterApply.Height = 50
    ProgBar.Cls
    ProgBar.Visible = False
    Screen.MousePointer = 0
End Select

Save_UndoAction

PauseFor 1

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If
End Sub

Private Sub mnuBrush_Click(Index As Integer)
Dim mnuI
  For mnuI = 0 To 5
    mnuBrush(mnuI).Checked = False
   DoEvents
  Next mnuI
mnuBrush(Index).Checked = True
cBrush = Index
ToolsButtonsUp
picTool_MouseDown 0, 1, 1, 1, 1

cTool = 0
picWork.DrawMode = vbCopyPen
picWork.FillStyle = 1

If Val(GetIniVal("General", "SaveSettings", IniName)) = 1 Then
SetIniVal "General", "cBrush", cBrush, IniName
End If
End Sub



Private Sub mnuBrushSet_Click()
Brush_set_Click
End Sub

Public Sub Brush_set_Click()
MousePointer = 11
frmPageSet.piccont(0).Visible = False
frmPageSet.piccont(1).Visible = True
frmPageSet.Show 1
MousePointer = 0
End Sub

Private Sub mnuColor_Click(Index As Integer)
Tmp_SaveLoad 'save to tmp file
Me.Refresh
Select Case Index
Case 0 'darken
   ' Fade 255 = no fading, 0 = all black
   m_cImage.Fade m_cDib, 240
   m_bDirty = True
   Render
Case 1 'lighten
   ' Lighten takes percentage:
   m_cImage.Lighten m_cDib, 20
   m_bDirty = True
   Render
Case 2 'negative
'picTmp.Picture = picWork.Image
'bmpNegative picTmp, picWork
    m_cImage.AddImages m_cDib, m_cDibBuffer, -1, -255, -255, -255, 0, 0, 0, 0
    m_cDibBuffer.PaintPicture m_cDib.hDC
    m_bDirty = True
    Render
Case 3 'grayscale
    m_cImage.GrayScale m_cDib
    m_bDirty = True
    Render
Case 4 'black & white
    m_cImage.BlackAndWhite m_cDib, m_cDibBuffer
    m_bDirty = True
    Render
Case 5 'colorise
frmColourise.Show 1
Case 6
    frmEffects.picEffectsCont(0).Visible = False 'cloth
    frmEffects.picEffectsCont(1).Visible = True 'replace
    frmEffects.picEffectsCont(2).Visible = False 'wave
    frmEffects.Caption = "Replace Color..."
    frmEffects.Show 1
End Select


Save_UndoAction

PauseFor 1

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If
End Sub

Private Sub mnuEdit_Click(Index As Integer)
editMnu_Click Index
End Sub

Public Sub editMnu_Click(Index As Integer)
Select Case Index
Case 0 'undo
Undo_Action
Case 1
Redo_Action
Case 3 'cut
picWork.DrawStyle = 2
picWork.DrawMode = vbInvert
picWork.FillStyle = 1
DrawItem picWork, "box", vbBlack, FirstX, FirstY, SecondX, SecondY
picWork.DrawStyle = 0
picWork.DrawMode = vbCopyPen
Paint_PClipBoard FirstX + 1, SecondX - 1, FirstY + 1, SecondY - 1, "cut"
FirstX = 0
FirstY = 0
SecondX = 0
SecondY = 0
Case 4 'copy
picWork.DrawStyle = 2
picWork.DrawMode = vbInvert
picWork.FillStyle = 1
DrawItem picWork, "box", vbBlack, FirstX, FirstY, SecondX, SecondY
picWork.DrawStyle = 0
picWork.DrawMode = vbCopyPen
Paint_PClipBoard FirstX + 1, SecondX - 1, FirstY + 1, SecondY - 1, "copy"
FirstX = 0
FirstY = 0
SecondX = 0
SecondY = 0
Case 5 'paste
Paint_PClipBoard FirstX, FirstY, FirstX + 1, FirstY + 1, "paste"
End Select
picWork.Picture = picWork.Image
picWork.Refresh

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

mnuEdit(3).Enabled = False
mnuEdit(4).Enabled = False
Paste_Status

If Index >= 3 And Index <= 5 Then
Save_UndoAction
End If
End Sub

Private Sub mnuFile_Click(Index As Integer)
fileMnu_Click Index
End Sub

Public Sub fileMnu_Click(Index As Integer)
Dim msg
Select Case Index
Case 0 'new
File_New
Case 1 'open
Open_Pic
If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
Case 3 'Save
Save_Pic picWork
Case 4  'Save As
Save_PicAs picWork
Case 6 'print
If picWork.Picture = 0 Then
stbar.Text = "Nothing to print!"
MsgBox "There's Nothing to Print!", vbCritical
stbar.Text = ""
Exit Sub
End If

Screen.MousePointer = 11
BmpTile picPrint, picWork
stbar.Text = "Print."
Screen.MousePointer = 0
frmPrint.Show 1
picPrint.Picture = LoadPicture()
Case 8 'exit
End
End Select
Screen.MousePointer = 0
End Sub

Public Sub File_New()
Screen.MousePointer = 11
picWork.Picture = LoadPicture()
picWork.Tag = "" 'clear the filename tag

picView.Picture = LoadPicture()
picPrint.Picture = LoadPicture()
picFilterApply.Picture = LoadPicture()
picFilterApply.Width = picWork.Width
picFilterApply.Height = picWork.Height

If picWork.Width > (150 * 15) Then
picWork.Width = 150 * 15
picWork.Height = 150 * 15
End If

    FirstX = 0
    FirstY = 0
    SecondX = 0
    SecondY = 0
    
picTool_MouseDown 0, 1, 0, 1, 1

Refresh_Undos_Redos 'undo redo boxes

picWork.FillStyle = 1
picWork.DrawMode = 13

If frmLens.Visible = True Then
frmLens.picLens.Picture = LoadPicture()
End If
ProgBar.Visible = False
Screen.MousePointer = 0

End Sub


Private Sub mnuFont_Click()
On Error GoTo handleFontErr
With CMDLG
.CancelError = True
.Flags = cdlCFApply Or cdlCFBoth Or cdlCFEffects
.FontName = lblSample.Font.Name
.FontSize = lblSample.Font.Size
.ShowFont
lblSample.Font.Name = .FontName
lblSample.Font.Size = .FontSize
End With
Exit Sub
handleFontErr:
Exit Sub
End Sub

Private Sub mnuGallery_Click(Index As Integer)
Gallerymnu_Click Index
End Sub

Public Sub Gallerymnu_Click(Index As Integer)
galMode = Index
frmPresetTiles.Show 1

If frmPresetTiles.GalCancel = False Then
If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If

'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

Save_UndoAction
End If
End Sub

Private Sub mnuHelp_Click(Index As Integer)
helpMnu_Click Index
End Sub

Public Sub helpMnu_Click(Index As Integer)
Dim FindHelpFor As String

If Index <= 2 Then
If Dir(TMHelpFile) = "" Then 'if the file is missing
MsgBox "Could not locate the Help file (TileMaker.hlp)." & vbCrLf & _
"Please do not remove the file from the TileMaker directory." & vbCrLf & _
"If the file has been deleted please try reinstalling the file.", vbCritical
Exit Sub
End If
End If

Select Case Index
Case 0 'contents
   HelpFunction Me.hWnd, HELP_INDEX, ""
Case 1 'Index
   HelpFunction Me.hWnd, HELP_PARTIALKEY, ""
Case 2 'search
FindHelpFor = InputBox("Search for Help on : ", "TileMaker Help")
If FindHelpFor = "" Then Exit Sub
   HelpFunction Me.hWnd, HELP_PARTIALKEY, FindHelpFor
Case 4 'about
frmAbout.Show 1
End Select
End Sub

Private Sub mnuImage_Click(Index As Integer)
imageMnu_Click Index
End Sub

Public Sub imageMnu_Click(Index As Integer)
Select Case Index
Case 0 'flip V
BmpFlip picTmp, picWork, "V"
Save_UndoAction
Case 1 'flip H
BmpFlip picTmp, picWork, "H"
Save_UndoAction
Case 4 'Mirror
BmpMirror picTmp, picWork
Save_UndoAction
Case 7 'offset
If picWork.Picture = 0 Then
MsgBox "Currently there's no image on the page.", vbCritical
Exit Sub
End If
FrmOffset.picPrev.Picture = picWork.Picture
FrmOffset.Show 1
If Not OffCancel Then
 If ChkPrev.Value = 1 Then
   picWork.Picture = picWork.Image
  cmdBtns_Click (0)
 Save_UndoAction
 End If
End If
Case 9
frmReflect.Show 1
Case 11 'back color
If picWork.Picture > 0 Then
MsgBox "This procedure will erase the contents of the current page. It is recommended" & vbCrLf & _
"that you should save the current work and proceed.", vbQuestion, "Tile Maker"
End If
picWork.Picture = LoadPicture()
ApplyColor picWork, "b"
Save_UndoAction
Case 13 'zoom
cmdBtns_Click (0)
frmZoom.Show 1
End Select

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If

'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

End Sub

Private Sub mnuOptions_Click(Index As Integer)
optionsMnu_Click Index
End Sub

Public Sub optionsMnu_Click(Index As Integer)
MousePointer = 11
Select Case Index
Case 0 'page settings
frmPageSet.piccont(0).Visible = True
frmPageSet.piccont(1).Visible = False
frmPageSet.Show 1
Screen.MousePointer = 11
Resize_UndoRedoPics
Screen.MousePointer = 0
Case 2 'pattern settings
frmPatSet.Show 1
Case 6 'preferences
frmOptions.Show 1
Setup_Envir
End Select
MousePointer = 0
End Sub



Private Sub mnuRotate_Click(Index As Integer)
Dim RclockWise As Boolean
Select Case Index
Case 0 'clockwise
RclockWise = True
Case 1 'anti clockwise
RclockWise = False
End Select

Screen.MousePointer = 11
    ProgBar.Visible = True
    ProgBar.Cls
    Rotate_Image picWork, picFilterApply, RclockWise
    picWork.Picture = picFilterApply.Image
    picFilterApply.Picture = LoadPicture()
    picFilterApply.Width = 50
    picFilterApply.Height = 50
    ProgBar.Cls
    ProgBar.Visible = False

Save_UndoAction

PauseFor 1

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

Screen.MousePointer = 0
End Sub

Private Sub mnuSharpenFilter_Click(Index As Integer)
Tmp_SaveLoad 'save to tmp file
Me.Refresh
Select Case Index
Case 0 'sharpen
ProcessImage eSharpen
Case 1 'sharpen more
ProcessImage eSharpenMore
Case 2 'unsharpen
ProcessImage eUnSharp
End Select

Save_UndoAction

PauseFor 1

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If
End Sub

Private Sub mnuSoftenFilter_Click(Index As Integer)
Tmp_SaveLoad 'save to tmp file
Me.Refresh
Select Case Index
Case 0 'soften
ProcessImage eSoften
Case 1 'soften more
ProcessImage eSoftenMore
End Select

Save_UndoAction

PauseFor 1

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If
End Sub

Private Sub mnuArtisticFilter_Click(Index As Integer)
Tmp_SaveLoad 'save to tmp file
Me.Refresh
Select Case Index
Case 0 'emboss
'ProcessImage eEmboss
Screen.MousePointer = 11
    ProgBar.Visible = True
    ProgBar.Cls
    Process_Image picWork, picFilterApply, fEmboss
    picWork.Picture = picFilterApply.Image
    picFilterApply.Picture = LoadPicture()
    picFilterApply.Width = 50
    picFilterApply.Height = 50
    ProgBar.Cls
    ProgBar.Visible = False
    Screen.MousePointer = 0
Case 1 'engrave
Screen.MousePointer = 11
    ProgBar.Visible = True
    ProgBar.Cls
    Process_Image picWork, picFilterApply, fEngrave
    picWork.Picture = picFilterApply.Image
    picFilterApply.Picture = LoadPicture()
    picFilterApply.Width = 50
    picFilterApply.Height = 50
    ProgBar.Cls
    ProgBar.Visible = False
    Screen.MousePointer = 0
Case 2 'add noise
    frmAddNoise.Caption = "Add Noise"
    frmAddNoise.Show 1
Case 3
    frmAddNoise.Caption = "Mosaic"
    frmAddNoise.Show 1
End Select

Save_UndoAction

PauseFor 1

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If
End Sub

Private Sub mnuSpecialFilter_Click(Index As Integer)
Me.Refresh
Select Case Index
Case 0 'Cloth
    frmEffects.picEffectsCont(0).Visible = True 'cloth
    frmEffects.picEffectsCont(1).Visible = False 'replace
    frmEffects.picEffectsCont(2).Visible = False 'Wave
    frmEffects.Caption = "Cloth Effect"
    frmEffects.Show 1
    
    If frmEffects.Didcancel = False Then
    Save_UndoAction
    End If

Case 1 'Disabled
Screen.MousePointer = 11
    ProgBar.Visible = True
    ProgBar.Cls
    Disabled_Effect picWork, picFilterApply
    picWork.Picture = picFilterApply.Image
    picFilterApply.Picture = LoadPicture()
    picFilterApply.Width = 50
    picFilterApply.Height = 50
    ProgBar.Cls
    ProgBar.Visible = False
Screen.MousePointer = 0
Save_UndoAction
Case 2 'wave
    frmEffects.picEffectsCont(0).Visible = False 'cloth
    frmEffects.picEffectsCont(1).Visible = False 'replace
    frmEffects.picEffectsCont(2).Visible = True 'Wave
    frmEffects.Caption = "Wave Effect"
    frmEffects.Show 1
    
    If frmEffects.Didcancel = False Then
    Save_UndoAction
    End If
End Select

PauseFor 1

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If
End Sub

Private Sub mnuTool_Click(Index As Integer)
toolMnu_Click Index
End Sub

Public Sub toolMnu_Click(Index As Integer)
Select Case Index
Case 0 'pencil
picTool_MouseDown 0, 1, 0, 0, 0
Case 1 'fill
picTool_MouseDown 11, 1, 0, 0, 0
Case 2 'text
picTool_MouseDown 9, 1, 0, 0, 0
Case 3 'line
picTool_MouseDown 1, 1, 0, 0, 0
Case 4 'rectangle
picTool_MouseDown 2, 1, 0, 0, 0
Case 5 'rectangle filled
picTool_MouseDown 6, 1, 0, 0, 0
Case 6 'circle
picTool_MouseDown 3, 1, 0, 0, 0
Case 7 'circle filled
picTool_MouseDown 7, 1, 0, 0, 0
Case 8 'chord
picTool_MouseDown 4, 1, 0, 0, 0
Case 9 'chord filled
picTool_MouseDown 8, 1, 0, 0, 0
Case 10 'arc
picTool_MouseDown 5, 1, 0, 0, 0
Case 11 'eraser
picTool_MouseDown 10, 1, 0, 0, 0
Case 12 'selector
picTool_MouseDown 12, 1, 0, 0, 0
Case 13 'color picker
picTool_MouseDown 13, 1, 0, 0, 0
End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
viewMnu_Click Index
End Sub

Public Sub viewMnu_Click(Index As Integer)
Select Case Index
Case 0
mnuView(0).Checked = Not mnuView(0).Checked
If mnuView(0).Checked = True Then
frmLens.Show
Draw_Preview frmMain.picWork, frmLens.picLens
Else
frmLens.picLens.Picture = LoadPicture()
Unload frmLens
End If
Case 1
mnuView(1).Checked = Not mnuView(1).Checked

If mnuView(1).Checked = True Then
Me.Height = 6330
statContainer.Visible = True
Else
Me.Height = 6060
statContainer.Visible = False
End If

End Select
End Sub

Private Sub mnuLoadfrom_Click()
loadfrom_Click
End Sub

Public Sub loadfrom_Click()
MousePointer = 11
frmPicSel.Show
MousePointer = 0

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

If frmPicSel.SelectCancel = False Then
Save_UndoAction
End If
End Sub



Private Sub Pat_Change()
MakeGrid picWork, nSteps.Text
If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If

'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

picTmp.Picture = picWork.Image
picTmp.Refresh
Save_UndoAction

    FirstX = 0
    FirstY = 0
    SecondX = 0
    SecondY = 0

End Sub

Private Sub Pat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Change Patterns #1"
End Sub

Private Sub Pat1_Change()
If chkClear.Value = 1 Then
picWork.Picture = LoadPicture()
End If
Draw_Pattern picWork, Pat1.Value, (picWork.ScaleWidth / 2)
If ChkPrev.Value = 1 Then
BmpTile picView, picWork
End If
picTmp.Picture = picWork.Picture
Save_UndoAction

    FirstX = 0
    FirstY = 0
    SecondX = 0
    SecondY = 0
    
End Sub

Private Sub Pat1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Change Patterns #2"
End Sub

Private Sub PicBack_Click()
Beep
End Sub

Private Sub PicBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = ""
End Sub

Private Sub picCol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ColOver
If Button And 1 Then
SColor(0).BackColor = picCol(Index).Point(X, Y)

If Val(GetIniVal("General", "SaveSettings", IniName)) = 1 Then
SetIniVal "Colors", "Left", SColor(0).BackColor, IniName
End If
Else
SColor(1).BackColor = picCol(Index).BackColor
If Val(GetIniVal("General", "SaveSettings", IniName)) = 1 Then
SetIniVal "Colors", "Right", SColor(1).BackColor, IniName
End If
End If
End Sub

Private Sub picCol_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cCol As String
Select Case Index
Case 0
cCol = "Black"
Case 1
cCol = "Navy Blue"
Case 2
cCol = "Dark Green"
Case 3
cCol = "Dark Cyan"
Case 4
cCol = "Dark Red"
Case 5
cCol = "Purple"
Case 6
cCol = "Dark Yellow"
Case 7
cCol = "Light Grey"
Case 8
cCol = "Dark Grey"
Case 9
cCol = "Blue"
Case 10
cCol = "Fuse Green"
Case 11
cCol = "Cyan"
Case 12
cCol = "Red"
Case 13
cCol = "Magenta"
Case 14
cCol = "Yellow"
Case 15
cCol = "White"
End Select
stbar.Text = "Color: " & cCol
End Sub





Private Sub picColorBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Color Pallete"
End Sub






Private Sub picTool_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 picTool_Mouse_Down Index, Button, Shift, X, Y
End Sub

Public Sub picTool_Mouse_Down(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tI

FillStyleBox.Visible = False 'pop box
picFillDisplay.Visible = False

ToolsButtonsUp

'check if there was  selection tool
If cTool = 12 Then
picWork.DrawStyle = 2
picWork.DrawMode = vbInvert
picWork.FillStyle = 1
DrawItem picWork, "box", vbBlack, FirstX, FirstY, SecondX, SecondY
picWork.DrawStyle = 0
picWork.DrawMode = vbCopyPen
    FirstX = 0
    FirstY = 0
    SecondX = 0
    SecondY = 0

End If
'set the tool
cTool = Index

If Index >= 6 And Index <= 8 Then 'solid Styles
    cFill_Style = 0
    CurFillStyle.Cls
    CurFillStyle.Refresh
    CurFillStyle.Line (2, 2)-(238, 238), vbBlack, BF
Else
    cFill_Style = 1
    CurFillStyle.Cls
    CurFillStyle.Refresh
    CurFillStyle.Line (2, 2)-(238, 238), vbBlack, B

End If

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

If Button = vbRightButton Then
If Index >= 6 And Index <= 8 Then 'fill styles pop
    FillStyleBox.Move picTool(Index).Left + picTool(Index).Width + 20, picTool(Index).Top + picTool(Index).Height + 20
    FillStyleBox.Visible = True
    picFillDisplay.Visible = True
End If
End If
End Sub

Private Sub picTool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = picTool(Index).ToolTipText
End Sub

Private Sub picToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = ""
End Sub

Private Sub picToolBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = ""
End Sub

Private Sub picView_Click()
Beep
End Sub

Private Sub picView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Preview Background"
End Sub

Private Sub picWork_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicWorkMouseDown Button, Shift, X, Y
End Sub

Private Sub picWork_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 0 Then
X = 0
End If
If Y < 0 Then
Y = 0
End If
If X > picWork.ScaleWidth Then
X = picWork.ScaleWidth
End If
If Y > picWork.ScaleHeight Then
Y = picWork.ScaleHeight
End If

PicWorkMouseMove Button, Shift, X, Y, True
End Sub

Private Sub picWork_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 0 Then
X = 0
End If
If Y < 0 Then
Y = 0
End If
If X > picWork.ScaleWidth Then
X = picWork.ScaleWidth
End If
If Y > picWork.ScaleHeight Then
Y = picWork.ScaleHeight
End If

PicWorkMouseUp Button, Shift, X, Y
End Sub

'public event for mousedown on picWork
Public Sub PicWorkMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ColOver As Long
    IsDown = True
    
'clear old box
If cTool = 12 Then 'selector
picWork.DrawStyle = 2
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
DrawItem picWork, "box", vbBlack, FirstX, FirstY, SecondX, SecondY
picWork.DrawStyle = 0
picWork.DrawMode = vbCopyPen
End If

'define new coordinates
    FirstX = X
    FirstY = Y
    SecondX = X
    SecondY = Y
    
MoveSelection = False
Select Case cTool
Case 0 'pencil
Select Case cBrush
Case 0 'normal brush
picWork.PSet (X, Y), (SColor(Button - 1).BackColor)
Case 1 'circle
picWork.Circle (X, Y), BcRad, (SColor(Button - 1).BackColor)
Case 2 'box
picWork.Line (X - BbSiz, Y - BbSiz)-(X + BbSiz, Y + BbSiz), (SColor(Button - 1).BackColor), B
Case 3 'rectangle
picWork.Line (X - BrSiz - 5, Y - BrSiz - 2)-(X + BrSiz + 5, Y + BrSiz + 2), (SColor(Button - 1).BackColor), B
Case 4 'box
picWork.Line (X - BbSiz, Y - BbSiz)-(X + BbSiz, Y + BbSiz), (SColor(Button - 1).BackColor), BF
Case 5 'rectangle
picWork.Line (X - BrSiz - 5, Y - BrSiz - 2)-(X + BrSiz + 5, Y + BrSiz + 2), (SColor(Button - 1).BackColor), BF
End Select
Case 1 'line
DrawItem picWork, "line", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 2 'box
DrawItem picWork, "box", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY

Case 11 'fill**********************
Call Fill_Area(picWork, X, Y, (SColor(Button - 1).BackColor))

Case 13 'color picker
SColor(Button - 1).BackColor = picWork.Point(X, Y)

Case 14 'stroke tool
strokeCol = picWork.Point(X, Y)
End Select
picWork.Picture = picWork.Image
picTmp.Picture = picWork.Image
picTmp.Refresh
End Sub

'public event for mousemove on picWork
Public Sub PicWorkMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal ShowPointer As Boolean)
Dim curMIcon As Integer

If cTool = 0 Then
curMIcon = 0
End If
If cTool >= 1 And cTool <= 8 Or cTool = 12 Then
curMIcon = 1
End If
If cTool = 9 Then
curMIcon = 2
End If
If cTool > 9 And cTool <= 10 Then
curMIcon = 3
End If
If cTool > 10 And cTool <= 11 Then
curMIcon = 4
End If
If cTool > 12 And cTool <= 13 Then
curMIcon = 5
End If
If cTool > 13 And cTool <= 14 Then
curMIcon = 6
End If

'if zoom tile is visible
If mnuView(0).Checked = True Then
With frmLens
.Pointer.Visible = ShowPointer
.Pointer.Move X * (.picLens.Width / picWork.ScaleWidth) - 210, _
Y * (.picLens.Height / picWork.ScaleHeight) - 220
frmLens.picLens.MouseIcon = curTool(curMIcon).Picture
End With
End If

Paste_Status 'paste check

picWork.MousePointer = 99
picWork.MouseIcon = curTool(curMIcon).Picture

If DrawSize <= 0 Then
DrawSize = 1
End If
picWork.DrawWidth = DrawSize

lblCoords.Caption = "X: " & X & ", Y: " & Y
stbar.Text = "X: " & X & ", Y: " & Y & "           Page Size : " & picWork.ScaleWidth & "x" & picWork.ScaleHeight
On Error Resume Next
If IsDown Then

Select Case cTool
Case 0 'pencil
Select Case cBrush
Case 0 'normal brush
picWork.PSet (X, Y), (SColor(Button - 1).BackColor)
Case 1 'circle
picWork.Circle (X, Y), BcRad, (SColor(Button - 1).BackColor)
Case 2 'box
picWork.Line (X - BbSiz, Y - BbSiz)-(X + BbSiz, Y + BbSiz), (SColor(Button - 1).BackColor), B
Case 3 'rectangle
picWork.Line (X - BrSiz - 5, Y - BrSiz - 2)-(X + BrSiz + 5, Y + BrSiz + 2), (SColor(Button - 1).BackColor), B
Case 4 'box
picWork.Line (X - BbSiz, Y - BbSiz)-(X + BbSiz, Y + BbSiz), (SColor(Button - 1).BackColor), BF
Case 5 'rectangle
picWork.Line (X - BrSiz - 5, Y - BrSiz - 2)-(X + BrSiz + 5, Y + BrSiz + 2), (SColor(Button - 1).BackColor), BF
End Select
Case 1 'line
picWork.DrawMode = vbInvert
DrawItem picWork, "line", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
    SecondX = X
    SecondY = Y
DrawItem picWork, "line", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
picWork.DrawMode = vbCopyPen
Case 2 'box
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
picWork.ForeColor = (SColor(Button - 1).BackColor)
DrawItem picWork, "box", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
'    SecondX = X
'    SecondY = Y
picWork.FillStyle = cFill_Style
picWork.ForeColor = (SColor(Button - 1).BackColor)
DrawItem picWork, "box", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
picWork.DrawMode = vbCopyPen
Case 3 'circle
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
DrawItem picWork, "circle", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
DrawItem picWork, "circle", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
picWork.DrawMode = vbCopyPen
Case 4 'chord
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
DrawItem picWork, "chord", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
DrawItem picWork, "chord", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
picWork.DrawMode = vbCopyPen
Case 5 'arc
picWork.DrawMode = vbInvert
DrawItem picWork, "arc", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
DrawItem picWork, "arc", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
picWork.DrawMode = vbCopyPen
Case 6 'box filled
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
DrawItem picWork, "box", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
DrawItem picWork, "box", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
picWork.DrawMode = vbCopyPen
Case 7 'circle filled
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
DrawItem picWork, "circle", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
DrawItem picWork, "circle", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 8 'chord filled
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
DrawItem picWork, "chord", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
DrawItem picWork, "chord", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
picWork.DrawMode = vbCopyPen
Case 10 'eraser
picWork.DrawWidth = 1
picWork.Line (X - 7, Y - 7)-(X + 7, Y + 7), vbWhite, BF
picWork.DrawWidth = DrawSize
Case 12 'selector
If Not MoveSelection Then
picWork.DrawStyle = 2
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
DrawItem picWork, "box", vbBlack, FirstX, FirstY, SecondX, SecondY
    SecondX = X
    SecondY = Y
picWork.FillStyle = cFill_Style
DrawItem picWork, "box", vbBlack, FirstX, FirstY, SecondX, SecondY
picWork.DrawStyle = 0
Else
picWork.MousePointer = 5
End If
Case 13 'color picker
If IsDown Then
SColor(Button - 1).BackColor = picWork.Point(X, Y)
End If
Case 14 'Stroke Tool
If IsDown Then
picWork.PSet (X, Y), strokeCol
End If
End Select
End If

picWork.Picture = picWork.Image
picTmp.Picture = picWork.Image
picTmp.Refresh

'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If
End Sub

'public event for mouseup on picWork
Public Sub PicWorkMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
IsDown = False

cxPos = X
cyPos = Y

Select Case cTool
Case 1 'line
picWork.DrawMode = vbCopyPen
    SecondX = X
    SecondY = Y
DrawItem picWork, "line", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 2 'box
picWork.DrawMode = vbCopyPen
picWork.FillStyle = cFill_Style
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
picWork.ForeColor = (SColor(Button - 1).BackColor)
DrawItem picWork, "box", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 3 'circle
picWork.DrawMode = vbCopyPen
picWork.FillStyle = cFill_Style
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
picWork.ForeColor = (SColor(Button - 1).BackColor)
DrawItem picWork, "circle", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 4 'chord
picWork.DrawMode = vbCopyPen
picWork.FillStyle = cFill_Style
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
picWork.ForeColor = (SColor(Button - 1).BackColor)
DrawItem picWork, "chord", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 5 'arc
picWork.DrawMode = vbCopyPen
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
picWork.ForeColor = (SColor(Button - 1).BackColor)
DrawItem picWork, "arc", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 6 'box filled
picWork.DrawMode = vbCopyPen
picWork.FillStyle = cFill_Style
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
picWork.ForeColor = (SColor(0).BackColor)
picWork.FillColor = (SColor(1).BackColor)
DrawItem picWork, "box", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 7 'circle filled
picWork.DrawMode = vbCopyPen
picWork.FillStyle = cFill_Style
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
picWork.ForeColor = (SColor(0).BackColor)
picWork.FillColor = (SColor(1).BackColor)
DrawItem picWork, "circle", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 8 'chord filled
picWork.DrawMode = vbCopyPen
picWork.FillStyle = cFill_Style
    wD = X - FirstX
    HT = Y - FirstY
    If Shift Then
        AdjustAspect wD, HT, 1 'with aspect
    End If
    SecondX = FirstX + wD
    SecondY = FirstY + HT
    'SecondX = X
    'SecondY = Y
picWork.ForeColor = (SColor(0).BackColor)
picWork.FillColor = (SColor(1).BackColor)
DrawItem picWork, "chord", (SColor(Button - 1).BackColor), FirstX, FirstY, SecondX, SecondY
Case 9 'text
frmText.Show 1
'If ctxtEnt = "" Then Exit Sub
'With picWork
'.DrawMode = vbCopyPen
'.FillStyle = cFill_Style
'.ForeColor = (SColor(Button - 1).BackColor)
'.Font.Name = ctxtFont
'.Font.Size = ctxtSize
'.FontBold = ctxtBold
'.FontItalic = ctxtItalic
'.FontStrikethru = ctxtStrike
'.FontUnderline = ctxtUline
'DrawText .hDC, x, y, ctxtEnt
'End With
Case 10 'eraser
picWork.DrawWidth = 1
picWork.Line (X - 7, Y - 7)-(X + 7, Y + 7), vbWhite, BF
picWork.DrawWidth = DrawSize
Case 12 'selector
picWork.DrawStyle = 2
picWork.DrawMode = vbInvert
picWork.FillStyle = cFill_Style
DrawItem picWork, "box", vbBlack, FirstX, FirstY, SecondX, SecondY
    SecondX = X
    SecondY = Y
picWork.FillStyle = cFill_Style
DrawItem picWork, "box", vbBlack, FirstX, FirstY, SecondX, SecondY
picWork.DrawStyle = 0
picWork.DrawMode = vbCopyPen
mnuEdit(3).Enabled = True 'cut
mnuEdit(4).Enabled = True 'copy
MoveSelection = True
End Select

picWork.Picture = picWork.Image
picTmp.Picture = picWork.Image
picTmp.Refresh

If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If
'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

If cTool = 13 Then 'color picker
SColor(Button - 1).BackColor = picWork.Point(X, Y)
Exit Sub
End If
If cTool <> 9 Then 'not text
Save_UndoAction
End If
End Sub


Private Sub preSets_Click()
If preSets.Text <> "" Then
pColMode = 1

    FirstX = 0
    FirstY = 0
    SecondX = 0
    SecondY = 0
    
presetTimer.Enabled = True
Save_UndoAction
End If
End Sub

Private Sub presetTimer_Timer()
Dim PresetVal As Integer
Dim PresetW As Single
Dim PresetH As Single
Dim chk As Integer
Dim SelStyle As Integer
Dim PreDWidth As Integer
Screen.MousePointer = 11
With frmPatSet ' no need to show

'clear the styles
For chk = 0 To UBound(cStyle) + 1
.chkSet(chk).Value = 0
Next chk

'change pen size
PreDWidth = picWork.DrawWidth
LinSiz.Value = 1
LinSiz_Change
'picWork.DrawWidth = 1

'get selected style
SelStyle = Right(preSets.Text, Len(preSets.Text) - 1)

Select Case SelStyle
Case 1
PresetW = 54
PresetH = 54
PresetVal = 15
.chkSet(4).Value = 1
.chkSet(5).Value = 1
.chkSet(6).Value = 1
.chkSet(7).Value = 1
Case 2
PresetW = 54
PresetH = 54
PresetVal = 30
.chkSet(4).Value = 1
.chkSet(5).Value = 1
Case 3
PresetW = 54
PresetH = 54
PresetVal = 18
.chkSet(6).Value = 1
.chkSet(7).Value = 1
.chkSet(8).Value = 1
.chkSet(9).Value = 1
Case 4
PresetVal = 14
Case 5
PresetVal = 15
Case 6
PresetVal = 23
Case 7
PresetVal = 24
Case 8
PresetVal = 25
Case 9
PresetVal = 11
Case 10
PresetVal = 12
Case 11
PresetVal = 18
Case 12
PresetVal = 22
Case 13
PresetVal = 20
Case 14
PresetVal = 30
Case 15
PresetVal = 86
Case 16
PresetVal = 9
Case 17
PresetVal = 18
Case 18
PresetVal = 40
Case 19
PresetVal = 52
Case 20
PresetVal = 65
Case 21
PresetVal = 80
Case 22
PresetVal = 27
Case 23
PresetVal = 40
Case 24
PresetVal = 55
Case 25
PresetVal = 80
Case 26
PresetVal = 15
Case 27
PresetVal = 18
Case 28
PresetVal = 32
Case 29
PresetVal = 6
Case 30
PresetVal = 9
Case 31
PresetVal = 20
Case 32
PresetVal = 23
Case 33
PresetVal = 53
Case 34
PresetVal = 94
Case 35
PresetVal = 6
Case 36
PresetVal = 9
Case 37
PresetVal = 14
Case 38
PresetVal = 94
Case 39
PresetVal = 16
Case 40
PresetVal = 19
Case 41
PresetVal = 32
Case 42
PresetVal = 38
Case 43
PresetVal = 40
Case 44
PresetVal = 41
Case 45
PresetVal = 43
Case 46
PresetVal = 57
Case 47
PresetVal = 59
Case 48
PresetVal = 69
Case 49
PresetVal = 75
Case 50
PresetVal = 79
Case 51
PresetVal = 80
Case 52
PresetVal = 100
Case 53
PresetVal = 53
Case 54
PresetVal = 100
Case 55
PresetVal = 75
Case 56
PresetVal = 79
Case 57
PresetVal = 94
Case 58
PresetVal = 94
Case 59
PresetVal = 99
Case 60
PresetVal = 20
Case 61
PresetVal = 25
Case 62
PresetVal = 35
Case 63
PresetVal = 100
Case 64
PresetVal = 36
Case 65
PresetVal = 45
Case 66
PresetVal = 54
Case 67
PresetVal = 81
Case 68
PresetVal = 11
Case 69
PresetVal = 100
Case 70
PresetVal = 23
Case 71
PresetVal = 30
Case 72
PresetVal = 35
Case 73
PresetVal = 40
Case 74
PresetVal = 41
Case 75
PresetVal = 80
End Select
'determine size and style() within the range...
If SelStyle >= 4 And SelStyle <= 8 Then
PresetW = 54
PresetH = 54
.chkSet(14).Value = 1
.chkSet(15).Value = 1
.chkSet(16).Value = 1
.chkSet(17).Value = 1
End If
'*
If SelStyle >= 9 And SelStyle <= 12 Then
PresetW = 54
PresetH = 54
.chkSet(18).Value = 1
.chkSet(19).Value = 1
.chkSet(20).Value = 1
.chkSet(21).Value = 1
.chkSet(22).Value = 1
.chkSet(23).Value = 1
.chkSet(24).Value = 1
.chkSet(25).Value = 1
End If
'*
If SelStyle >= 13 And SelStyle <= 15 Then
PresetW = 54
PresetH = 54
.chkSet(26).Value = 1
.chkSet(27).Value = 1
End If
'*
If SelStyle >= 16 And SelStyle <= 21 Then
PresetW = 54
PresetH = 54
.chkSet(34).Value = 1
.chkSet(35).Value = 1
End If
'*
If SelStyle >= 22 And SelStyle <= 25 Then
PresetW = 54
PresetH = 54
.chkSet(30).Value = 1
.chkSet(31).Value = 1
End If
'*
If SelStyle >= 26 And SelStyle <= 28 Then
PresetW = 54
PresetH = 54
.chkSet(66).Value = 1
.chkSet(67).Value = 1
.chkSet(68).Value = 1
.chkSet(69).Value = 1
End If
'*
If SelStyle >= 29 And SelStyle <= 34 Then
PresetW = 54
PresetH = 54
.chkSet(70).Value = 1
.chkSet(71).Value = 1
.chkSet(72).Value = 1
.chkSet(73).Value = 1
.chkSet(74).Value = 1
.chkSet(75).Value = 1
.chkSet(76).Value = 1
.chkSet(77).Value = 1
End If
'*
If SelStyle >= 35 And SelStyle <= 38 Then
PresetW = 54
PresetH = 54
.chkSet(78).Value = 1
.chkSet(79).Value = 1
.chkSet(80).Value = 1
.chkSet(81).Value = 1
.chkSet(82).Value = 1
.chkSet(83).Value = 1
.chkSet(84).Value = 1
.chkSet(85).Value = 1
End If
'*
If SelStyle >= 39 And SelStyle <= 52 Then
PresetW = 54
PresetH = 54
.chkSet(90).Value = 1
.chkSet(91).Value = 1
.chkSet(92).Value = 1
.chkSet(93).Value = 1
.chkSet(94).Value = 1
.chkSet(95).Value = 1
.chkSet(96).Value = 1
.chkSet(97).Value = 1
End If
'*
If SelStyle >= 53 And SelStyle <= 54 Then
PresetW = 54
PresetH = 54
.chkSet(98).Value = 1
.chkSet(99).Value = 1
.chkSet(100).Value = 1
.chkSet(101).Value = 1
End If
'*
If SelStyle >= 55 And SelStyle <= 57 Then
PresetW = 54
PresetH = 54
.chkSet(102).Value = 1
.chkSet(103).Value = 1
.chkSet(104).Value = 1
.chkSet(105).Value = 1
End If
'*
If SelStyle >= 58 And SelStyle <= 59 Then
PresetW = 54
PresetH = 54
.chkSet(106).Value = 1
.chkSet(107).Value = 1
.chkSet(108).Value = 1
.chkSet(109).Value = 1
End If
'*
If SelStyle >= 60 And SelStyle <= 63 Then
PresetW = 54
PresetH = 54
.chkSet(110).Value = 1
.chkSet(111).Value = 1
.chkSet(112).Value = 1
.chkSet(113).Value = 1
End If
'*
If SelStyle >= 64 And SelStyle <= 67 Then
PresetW = 54
PresetH = 54
.chkSet(118).Value = 1
End If
'*
If SelStyle >= 68 And SelStyle <= 69 Then
PresetW = 54
PresetH = 54
.chkSet(119).Value = 1
.chkSet(120).Value = 1
.chkSet(121).Value = 1
.chkSet(122).Value = 1
.chkSet(123).Value = 1
.chkSet(124).Value = 1
.chkSet(125).Value = 1
.chkSet(126).Value = 1
End If
'*
If SelStyle >= 70 And SelStyle <= 75 Then
PresetW = 54
PresetH = 54
.chkSet(127).Value = 1
.chkSet(128).Value = 1
.chkSet(129).Value = 1
.chkSet(130).Value = 1
End If
End With

Unload frmPatSet
picWork.Width = PresetW * 15
picWork.Height = PresetH * 15
Pat.Value = PresetVal
picWork.Picture = picWork.Image

'old pen size
LinSiz.Value = PreDWidth
LinSiz_Change


If ChkPrev.Value = 1 Then
cmdBtns_Click (0)
End If

'lens picture
If mnuView(0).Checked = True Then
Draw_Preview frmMain.picWork, frmLens.picLens
End If

Me.Refresh
Screen.MousePointer = 0
presetTimer.Enabled = False
End Sub

Private Sub SColor_DblClick(Index As Integer)
ApplyColor SColor(Index), "b"
End Sub

Private Sub SColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
stbar.Text = "Left Button Color   [Double click for custom colors]"
Case 1
stbar.Text = "Right Button Color  [Double click for custom colors]"
End Select
End Sub





Sub ReCreateSettings()
'general
SetIniVal "General", "SaveSettings", 0, IniName
SetIniVal "General", "AlwaysPrev", 0, IniName
SetIniVal "General", "ShowText", 0, IniName
SetIniVal "General", "DrawSize", 1, IniName
SetIniVal "General", "cBrush", 0, IniName
SetIniVal "General", "ColorMode", 1, IniName

SetIniVal "General", "Animation", 1, IniName
SetIniVal "General", "StartupDir", App.Path, IniName

'CustomGallery
SetIniVal "Gallery", "CustomGallery", App.Path & "\CGallery.DAT", IniName

'UndoRedo
SetIniVal "General", "MaxLevelUR", 5, IniName

'brush
SetIniVal "Brush", "BcRad", 5, IniName
SetIniVal "Brush", "BbSiz", 2, IniName
SetIniVal "Brush", "BrSiz", 2, IniName
'pageset
SetIniVal "PageSettings", "Width", 54, IniName
SetIniVal "PageSettings", "Height", 54, IniName
SetIniVal "PageSettings", "Proportional", 0, IniName
SetIniVal "PageSettings", "AutoSize", 0, IniName
'colors
SetIniVal "Colors", "Left", vbBlack, IniName
SetIniVal "Colors", "Right", vbWhite, IniName

End Sub

Private Sub tButton_Click(Index As Integer)
Select Case Index
Case 0 'new
mnuFile_Click 0
Case 1 'open
mnuFile_Click 1
Case 2 'save
mnuFile_Click 3
Case 3 'print
mnuFile_Click 6
Case 4 'cut
mnuEdit_Click 3
Case 5 'copy
mnuEdit_Click 4
Case 6 'paste
mnuEdit_Click 5
Case 7 'undo
mnuEdit_Click 0
Case 8 'redo
mnuEdit_Click 1
Case 9 'zoomed tile
mnuView_Click 0
Case 10 'background
mnuImage_Click 13
Case 11 'loadfrom
mnuLoadfrom_Click
Case 12 'add to gallery
mnuAddtoCGallery_Click
Case 13 'pattern set
mnuOptions_Click 2
Case 14 'page set
mnuOptions_Click 0
Case 15 'help
mnuHelp_Click 1
Case 16 'view gallery
Load_A_Gallery
End Select
End Sub

Private Sub tButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MainToolBarTimer.Enabled = False
tButton(Index).Line (-1, -1)-(23, 23), vbBlack, BF
End If
End Sub


Private Sub tButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = tButton(Index).ToolTipText
End Sub

Private Sub tButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
tButton(Index).Cls
tButton(Index).Refresh
MainToolBarTimer.Enabled = True
End Sub

Private Sub ToolBoxTimer_Timer()
'tool box selection
If Anim = 1 Then
picTool(cTool).DrawMode = vbInvert
picTool(cTool).Line (-1, -1)-(15, 15), , BF
Else
picTool(cTool).DrawMode = vbCopyPen
picTool(cTool).Line (0, 0)-(15, 15), vbRed, B
End If
End Sub

Private Sub txtSample_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
stbar.Text = "Change Sample Text"
End Sub

Sub ApplyColor(curObject As Object, ByVal ForeBack As String)
On Error GoTo handleColErr
stbar.Text = "Color Picker"
With CMDLG
.CancelError = True
.Flags = cdlCCRGBInit
.ShowColor
End With
If UCase(Left(ForeBack, 1)) = "F" Then
curObject.ForeColor = CMDLG.Color
Else
curObject.BackColor = CMDLG.Color
End If
stbar.Text = ""
Exit Sub

handleColErr:
stbar.Text = ""
Exit Sub
End Sub

Sub ToolsButtonsUp()
Dim ptlI
 For ptlI = 0 To Max_Tool 'change to button ups
    picTool(ptlI).Cls
 Next
End Sub

Sub Resize_UndoRedoPics()
Dim pURI
ProgBar.Visible = True
For pURI = 0 To Max_Undo
    'undo pics
    frmMain.picUndo(pURI).Width = frmMain.picWork.Width
    frmMain.picUndo(pURI).Height = frmMain.picWork.Height
    'redo pics
    frmMain.picRedo(pURI).Width = frmMain.picWork.Width
    frmMain.picRedo(pURI).Height = frmMain.picWork.Height
    Update_Progress (pURI * 100) / Max_Undo, "Processing..."
Next
ProgBar.Visible = False
End Sub

Sub Refresh_Undos_Redos()
Dim PUndo, PRedo
'undo boxes
For PUndo = 0 To Max_Undo
picUndo(PUndo).Picture = LoadPicture()
Next
'redo boxes
For PRedo = 0 To Max_Redo
picRedo(PRedo).Picture = LoadPicture()
Next
'disable menus
mnuEdit(0).Enabled = False
mnuEdit(0).Caption = "Can't Undo"
mnuEdit(1).Enabled = False
mnuEdit(1).Caption = "Can't Redo"

UNDOS = 1 'undo actions
REDOS = 1 'undo actions

End Sub


Sub Save_UndoAction()
Dim UndoI
If UNDOS <= Max_Undo Then
UNDOS = UNDOS + 1
End If
For UndoI = Max_Undo To 1 Step -1
picUndo(UndoI).Picture = picUndo(UndoI - 1).Image
picUndo(UndoI).Refresh
Next
picUndo(0).Picture = picWork.Image 'undo picture update
picUndo(0).Refresh
mnuEdit(0).Enabled = True
mnuEdit(0).Caption = "Undo"
End Sub

Sub Save_RedoAction()
Dim RedoI
If REDOS <= Max_Redo Then
REDOS = REDOS + 1
End If
For RedoI = Max_Redo To 1 Step -1
picRedo(RedoI).Picture = picRedo(RedoI - 1).Image
picRedo(RedoI).Refresh
Next
picRedo(0).Picture = picUndo(0).Image 'undo picture update
picRedo(0).Refresh
mnuEdit(1).Enabled = True
mnuEdit(1).Caption = "Redo"
End Sub


Sub Undo_Action()
Dim UndoI
If UNDOS < 2 Then Exit Sub

tButton(7).Cls
tButton(7).Refresh
tButton(7).DrawMode = vbInvert


UNDOS = UNDOS - 1

Save_RedoAction 'save redo actions

picWork.Picture = picUndo(1).Image 'undo the last action
picWork.Refresh

For UndoI = 0 To Max_Undo - 1
picUndo(UndoI).Picture = picUndo(UndoI + 1).Image
picUndo(UndoI).Refresh
picUndo(UndoI + 1).Picture = LoadPicture()
Next
End Sub

Sub Redo_Action()
Dim RedoI

If REDOS < 2 Then Exit Sub

tButton(8).Cls
tButton(8).Refresh
tButton(8).DrawMode = vbInvert

REDOS = REDOS - 1

picWork.Picture = picRedo(0).Image 'redo the last action
picWork.Refresh

Save_UndoAction 'save undo actions

For RedoI = 0 To Max_Redo - 1
picRedo(RedoI).Picture = picRedo(RedoI + 1).Image
picRedo(RedoI).Refresh
picRedo(RedoI + 1).Picture = LoadPicture()
Next
End Sub




'sets the width and height of clip picturebox the same as selection
Sub Paint_PClipBoard(ByVal X1 As Single, ByVal X2 As Single, _
ByVal Y1 As Single, ByVal Y2 As Single, ByVal curCommand As String)
Dim tmp As Single
Dim sMode
On Error Resume Next
sMode = picWork.ScaleMode
picWork.ScaleMode = 1
    'check the values to see if its OK
    If X1 > X2 Then
        tmp = X1
        X1 = X2
        X2 = tmp
    End If
    If Y1 > Y2 Then
        tmp = Y1
        Y1 = Y2
        Y2 = tmp
    End If
If curCommand = "paste" Then ' if pasting
    ' Make sure an image exists. This will happen
    ' if the clipboard does not contain a bitmap
    ' and the user presses ^V.
    If Not Clipboard.GetFormat(vbCFBitmap) Then Exit Sub
    
    picClipBoard.AutoSize = True
    picClipBoard.Picture = Clipboard.GetData(vbCFBitmap)
    picClipBoard.AutoSize = False

    picWork.PaintPicture _
        picClipBoard.Picture, _
        X1, Y1, _
        picClipBoard.ScaleWidth, picClipBoard.ScaleHeight, _
        0, 0, picClipBoard.ScaleWidth, picClipBoard.ScaleHeight
    picWork.Refresh
    picWork.Picture = picWork.Image
    picWork.ScaleMode = sMode

Exit Sub 'get out now
End If

'set sizes
picClipBoard.Width = (X2 - X1) * 15
picClipBoard.Height = (Y2 - Y1) * 15

'paint image
picClipBoard.Picture = LoadPicture()
    picClipBoard.PaintPicture _
        picWork.Picture, _
        0, 0, (X2 - X1 + 1) * 15, (Y2 - Y1 + 1) * 15, _
        X1 * 15, Y1 * 15, (X2 - X1 + 1) * 15, (Y2 - Y1 + 1) * 15
   
   
Select Case curCommand
Case "cut"
    Clipboard.Clear
    Clipboard.SetData picClipBoard.Image, vbCFBitmap
    
    Paste_Status
    
    picWork.Line _
        ((X1 + 1) * 15, (Y1 + 1) * 15)-((X2 - 1) * 15, (Y2 - 1) * 15), _
        picWork.BackColor, BF
Case "copy"
    ' Copy to the clipboard.
    Clipboard.Clear
    Clipboard.SetData picClipBoard.Image, vbCFBitmap
    
    Paste_Status
End Select
    ' Make the picture part of the background.
    picWork.Refresh
    picWork.Picture = picWork.Image
    picWork.ScaleMode = sMode
Exit Sub
End Sub

Sub Paste_Status()
    mnuEdit(5).Enabled = Clipboard.GetFormat(vbCFBitmap)
End Sub


Public Sub ProcessImage(ByVal eType As EFilterTypes)
    With m_cImage
        .FilterType = eType
        .ProcessImage m_cDib, m_cDibBuffer
        Render
        m_bDirty = True
    End With
End Sub

Public Sub Render()
    picWork.Width = m_cDib.Width * Screen.TwipsPerPixelX
    picWork.Height = m_cDib.Height * Screen.TwipsPerPixelY
    m_cDib.PaintPicture picWork.hDC
    picWork.Refresh
End Sub

Public Function OpenFile(ByVal sFIle As String, Optional ByVal bIsTemp As Boolean = False) As Boolean
Dim sPic As StdPicture
On Error GoTo OpenFileError
    
    Set sPic = LoadPicture(sFIle)
    
    m_cDib.CreateFromPicture sPic
    m_cDibBuffer.Create m_cDib.Width, m_cDib.Height
    
    OpenFile = True
    Exit Function
OpenFileError:
    MsgBox "An error occured trying to open this file: " & Err.Description, vbExclamation
    Exit Function
End Function

Sub Tmp_SaveLoad()
On Error GoTo HandleTmpErr
'save to file
SavePicture picWork.Picture, tmpFName
'update the class
OpenFile tmpFName
'delete the tmp file
Kill tmpFName
Exit Sub

HandleTmpErr: 'handle any ERRs
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

Public Sub AddNoise(ByVal bRandom As Boolean, _
ByVal lAmount As Long, ByVal RVal As Long, _
ByVal GVal As Long, ByVal BVal As Long)
    
    m_cImage.RSelection = RVal
    m_cImage.GSelection = GVal
    m_cImage.BSelection = BVal
    
    m_cImage.AddNoise m_cDib, lAmount, bRandom
    m_bDirty = True
    Render
    m_bDirty = True
End Sub


'Adjusts the aspect while drawing boxes
Private Sub AdjustAspect(ByRef ww_wid As Single, ByRef ww_hgt As Single, ByVal view_aspect As Single)
Dim ww_aspect As Single
Dim sign As Integer

    ' Don't divide by zero.
    If ww_wid = 0 Or ww_hgt = 0 Or view_aspect = 0 Then Exit Sub
    
    ww_aspect = ww_hgt / ww_wid
    If ww_aspect < 0 Then
        sign = -1
    Else
        sign = 1
    End If
    ww_aspect = Abs(ww_aspect)

    If ww_aspect > view_aspect Then
        ' The world window is too tall and thin. Make it wider.
        ww_wid = sign * ww_hgt / view_aspect
    Else
        ' The world window is too short and squat. Make it taller.
        ww_hgt = sign * view_aspect * ww_wid
    End If
End Sub

Private Sub Load_A_Gallery()
If mnuGallery(1).Enabled = True Then
frmGalleryAsk.Show 1
Else
mnuGallery_Click 0
End If
End Sub
