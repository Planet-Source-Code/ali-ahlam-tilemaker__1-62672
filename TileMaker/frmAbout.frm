VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4260
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5715
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":030A
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox piccont 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   90
      ScaleHeight     =   3375
      ScaleWidth      =   5535
      TabIndex        =   2
      Top             =   360
      Width           =   5535
      Begin VB.Timer movetimer 
         Interval        =   10
         Left            =   4830
         Top             =   488
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   5579
         Y1              =   2108
         Y2              =   2108
      End
      Begin VB.Label lblReg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   840
         Width           =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   0
         X2              =   5564
         Y1              =   2093
         Y2              =   2093
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Index           =   2
         Left            =   990
         TabIndex        =   3
         Top             =   8
         Width           =   3540
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":55C7
         ForeColor       =   &H00FFFFFF&
         Height          =   1185
         Left            =   165
         TabIndex        =   8
         Top             =   2273
         Width           =   5430
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   645
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   974
         TabIndex        =   6
         Top             =   0
         Width           =   3540
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":5709
         ForeColor       =   &H00FFFFFF&
         Height          =   1050
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   5445
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   990
         TabIndex        =   4
         Top             =   35
         Width           =   3540
      End
      Begin VB.Image picIcon 
         Height          =   480
         Left            =   150
         Picture         =   "frmAbout.frx":5854
         Top             =   8
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2280
      TabIndex        =   0
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Label lblCright 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Ahlam 1998 - 1999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5295
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LeftVal As Integer
Dim TopVal As Integer


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
frmMain.stbar.Text = "About Tile Maker"
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle(0).Caption = App.Title
    lblTitle(1).Caption = App.Title
    lblTitle(2).Caption = App.Title
    TopVal = 1
    LeftVal = 1
    lblReg.Caption = "Owned by: " & UserName
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.stbar.Text = ""
End Sub

Private Sub movetimer_Timer()
lblCright.Left = lblCright.Left + LeftVal
lblCright.Top = lblCright.Top + TopVal

If lblCright.Left >= 180 Then
LeftVal = -1
End If
If lblCright.Left <= 160 Then
LeftVal = 1
End If

If lblCright.Top >= 10 Then
TopVal = -1
End If
If lblCright.Top <= 5 Then
TopVal = 1
End If
End Sub
