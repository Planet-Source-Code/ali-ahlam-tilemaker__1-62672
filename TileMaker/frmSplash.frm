VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3825
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Animator 
      Interval        =   50
      Left            =   5520
      Top             =   3600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label lblCright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Ahlam 1998"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   225
      Left            =   3960
      TabIndex        =   8
      Top             =   840
      Width           =   2040
   End
   Begin VB.Label lblstat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to TileMaker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   1905
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   765
      Index           =   2
      Left            =   2650
      TabIndex        =   0
      Top             =   1090
      Width           =   2430
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   1110
      Width           =   2430
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed To Any Owner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   210
      Left            =   4965
      TabIndex        =   4
      Top             =   240
      Width           =   1980
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For 32 bit Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   360
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   1800
      Width           =   885
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is a freeware program. Please visit  www.geocities.com/researchtriangle/6731  for more."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   300
      TabIndex        =   1
      Top             =   3360
      Width           =   6840
   End
   Begin VB.Image imgLogo 
      Height          =   2025
      Left            =   360
      Picture         =   "frmSplash.frx":822C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   765
      Index           =   0
      Left            =   2645
      TabIndex        =   6
      Top             =   1080
      Width           =   2430
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim TopVal
Dim LeftVal

Private Sub Animator_Timer()
lblCright.Left = lblCright.Left + LeftVal
lblCright.Top = lblCright.Top + TopVal

If lblCright.Left >= 288 Then
LeftVal = -1
End If
If lblCright.Left <= 264 Then
LeftVal = 1
End If

If lblCright.Top >= 64 Then
TopVal = -1
End If
If lblCright.Top <= 56 Then
TopVal = 1
End If
End Sub


Private Sub Form_Activate()
frmSplash.Refresh
End Sub

Private Sub Form_Load()
TopVal = 1
LeftVal = 1
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName(0).Caption = App.Title
    lblProductName(1).Caption = App.Title
    lblProductName(2).Caption = App.Title
    lblLicenseTo.Caption = "Licensed to: " & UserName
End Sub

