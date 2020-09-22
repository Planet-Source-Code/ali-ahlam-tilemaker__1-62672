VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fram 
      Caption         =   "Number of copies"
      Height          =   855
      Index           =   2
      Left            =   3480
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
      Begin ComCtl2.UpDown UDCopies 
         Height          =   285
         Left            =   1935
         TabIndex        =   12
         ToolTipText     =   "Number of copies"
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196615
         OrigLeft        =   1080
         OrigTop         =   480
         OrigRight       =   1320
         OrigBottom      =   735
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
      End
      Begin VB.TextBox txtCopies 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "1"
         ToolTipText     =   "Numberof copies"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblcopies 
         AutoSize        =   -1  'True
         Caption         =   "Copies to Print:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmds 
      Caption         =   "Printer &Setup"
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   9
      ToolTipText     =   "Printer Setup"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      ToolTipText     =   "Print"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame fram 
      Caption         =   "Preview"
      Height          =   3375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox picPrev 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   120
         ScaleHeight     =   3015
         ScaleWidth      =   3015
         TabIndex        =   4
         ToolTipText     =   "Preview"
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fram 
      Caption         =   "Print Type"
      Height          =   1095
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      ToolTipText     =   "Print Type"
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox chkStrech 
         Caption         =   "Streched"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         ToolTipText     =   "Stretched Tile"
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton printType 
         Caption         =   "Tiled Background"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Print Tiled Background"
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton printType 
         Caption         =   "Single Tile"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Print Single Tile"
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Label lblmsg 
      Height          =   195
      Left            =   3480
      TabIndex        =   8
      Top             =   2760
      Width           =   2445
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkStrech_Click()
If chkStrech.Value = 1 Then
Draw_Preview frmMain.picWork, picPrev
Else
picPrev.Picture = frmMain.picWork.Picture
End If
End Sub

Private Sub cmds_Click(Index As Integer)
Select Case Index
Case 0 'print
Screen.MousePointer = 11
If printType(0).Value = True Then
If chkStrech.Value = 1 Then
Print_Tile frmMain.picWork, 0, UDCopies.Value 'single tile
Else
Print_Tile frmMain.picWork, 1, UDCopies.Value 'single tile
End If
Else
Print_Tile frmMain.picPrint, 0, UDCopies.Value  'tiled back
End If
Unload Me
Screen.MousePointer = 0
Case 1 'cancel
Unload Me
Case 2 'printer setup
On Error GoTo handleprintsetErr
With frmMain.CMDLG
.Flags = cdlPDPrintSetup
.CancelError = True
.PrinterDefault = True
.ShowPrinter
End With
End Select

Exit Sub

handleprintsetErr:
Screen.MousePointer = 0
If Err <> 32755 Then
MsgBox Err.Description, vbCritical
End If
Exit Sub
End Sub

Private Sub Form_Load()
printType(0).Value = True
printType_Click (0)
End Sub


Private Sub printType_Click(Index As Integer)
picPrev.ScaleMode = 3
Select Case Index
 Case 0
 chkStrech.Enabled = True
 If Not chkStrech.Value = 1 Then
 picPrev.Picture = frmMain.picWork.Picture
 Else
 Draw_Preview frmMain.picWork, picPrev
 End If
 Case 1
 chkStrech.Enabled = False
 BmpTile picPrev, frmMain.picWork
End Select
End Sub



