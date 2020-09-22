VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings..."
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApplyDefSet 
      Caption         =   "Reset Settings"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      ToolTipText     =   "Load Default Settings"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "OK"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   50
      ScaleHeight     =   3375
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.PictureBox TabPic 
         BorderStyle     =   0  'None
         Height          =   2535
         Index           =   2
         Left            =   2280
         ScaleHeight     =   2535
         ScaleWidth      =   4455
         TabIndex        =   8
         Top             =   1320
         Width           =   4455
         Begin ComCtl2.UpDown URMax 
            Height          =   285
            Left            =   2040
            TabIndex        =   18
            ToolTipText     =   "Undo Redo Level"
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   5
            BuddyControl    =   "txtUndoRedoLevels"
            BuddyDispid     =   196615
            OrigLeft        =   2160
            OrigTop         =   360
            OrigRight       =   2400
            OrigBottom      =   615
            Max             =   23
            Min             =   5
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtUndoRedoLevels 
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            ToolTipText     =   "Undo Redo Levels"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblURlevel 
            AutoSize        =   -1  'True
            Caption         =   "Maximum Level:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1140
         End
      End
      Begin VB.PictureBox TabPic 
         BorderStyle     =   0  'None
         Height          =   2535
         Index           =   1
         Left            =   1440
         ScaleHeight     =   2535
         ScaleWidth      =   4455
         TabIndex        =   7
         Top             =   1080
         Width           =   4455
         Begin VB.CommandButton cmdNewGalry 
            Caption         =   "&New"
            Height          =   375
            Left            =   1440
            TabIndex        =   21
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CommandButton cmdCustGalry 
            Caption         =   "C&hange"
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtGalName 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label lblGallry 
            AutoSize        =   -1  'True
            Caption         =   "Custom Gallery:"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.PictureBox TabPic 
         BorderStyle     =   0  'None
         Height          =   2535
         Index           =   0
         Left            =   240
         ScaleHeight     =   2535
         ScaleWidth      =   4455
         TabIndex        =   6
         Top             =   720
         Width           =   4455
         Begin VB.CommandButton cmdBrowseDir 
            Caption         =   "..."
            Height          =   280
            Left            =   1800
            TabIndex        =   13
            ToolTipText     =   "Change Startup Directory"
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtStartUp 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkAnimateTBar 
            Caption         =   "Animate Toolbar"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "Animate Toolbar"
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkSaveSet 
            Caption         =   "Save Settings"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            ToolTipText     =   "Save Settings"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblStartup 
            AutoSize        =   -1  'True
            Caption         =   "Startup Directory:"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   1230
         End
      End
      Begin VB.CommandButton cmdTab 
         Caption         =   "&Undo/Redo"
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdTab 
         Caption         =   "G&allery"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdTab 
         Caption         =   "&General"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BrowseInfo
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (bBrowse As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long

Private Const Tot_Tabs = 2

Private Sub chkSaveSet_Click()
SetIniVal "General", "SaveSettings", chkSaveSet.Value, IniName 'save settings
End Sub

Private Sub cmdApplyDefSet_Click()
frmMain.ReCreateSettings
chkSaveSet.Value = Val(GetIniVal("General", "SaveSettings", IniName)) 'save settings
chkAnimateTBar.Value = Val(GetIniVal("General", "Animation", IniName)) 'animation
txtStartUp.Text = GetIniVal("General", "StartupDir", IniName) 'start up
txtGalName.Text = GetIniVal("Gallery", "CustomGallery", IniName) 'custom gallery
txtUndoRedoLevels.Text = Val(GetIniVal("General", "MaxLevelUR", IniName)) 'Undo/Redo
End Sub

Private Sub cmdBrowseDir_Click()
Dim BrwsPath
BrwsPath = BrowseForDirectory(Me)
If BrwsPath <> "" Then
    txtStartUp.Text = BrwsPath
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCustGalry_Click()
With frmMain.CMDLG
.Filename = ""
.Filter = "Data files (*.DAT)|*.DAT"
.CancelError = False
.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNNoChangeDir
.DialogTitle = "Select the Custom Gallery"
.ShowOpen
End With
If frmMain.CMDLG.Filename = "" Then Exit Sub
Data_file = frmMain.CMDLG.Filename
txtGalName.Text = Data_file
End Sub

Private Sub cmdNewGalry_Click()
With frmMain.CMDLG
.Filename = ""
.Filter = "Data files (*.DAT)|*.DAT"
.CancelError = False
.Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt Or cdlOFNNoChangeDir Or cdlOFNNoLongNames Or cdlOFNNoChangeDir
.DialogTitle = "Enter New Gallery Name"
.ShowSave
End With
If frmMain.CMDLG.Filename = "" Then Exit Sub
Data_file = frmMain.CMDLG.Filename
Open Data_file For Output As #1
Close #1
txtGalName.Text = Data_file
End Sub

Private Sub cmdOK_Click()
'saveinivalues
SetIniVal "General", "SaveSettings", chkSaveSet.Value, IniName
SetIniVal "General", "Animation", chkAnimateTBar.Value, IniName
SetIniVal "General", "StartupDir", txtStartUp.Text, IniName
SetIniVal "Gallery", "CustomGallery", txtGalName.Text, IniName
SetIniVal "General", "MaxLevelUR", txtUndoRedoLevels.Text, IniName

Unload Me
End Sub

Private Sub cmdTab_GotFocus(Index As Integer)
Select_Tab Picture1, Index
Show_Tab Index
End Sub

Private Sub cmdTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdTab(Index).ToolTipText = cmdTab(Index).Caption
End Sub

Private Sub Form_Load()
Arrange_Tabs
Select_Tab Picture1, 0

chkSaveSet.Value = Val(GetIniVal("General", "SaveSettings", IniName)) 'save settings
chkAnimateTBar.Value = Val(GetIniVal("General", "Animation", IniName)) 'animation
txtStartUp.Text = GetIniVal("General", "StartupDir", IniName) 'start up
txtGalName.Text = GetIniVal("Gallery", "CustomGallery", IniName) 'custom gallery
txtUndoRedoLevels.Text = Val(GetIniVal("General", "MaxLevelUR", IniName)) 'Undo/Redo
End Sub


Sub Select_Tab(pic As PictureBox, ByVal nIndex)
Dim LeftOf, Wid, cmdI
Const BorderCol1 = &H8000000C 'dark grey
Const BorderCol2 = &H80000009 'white
Const BorderCol3 = &H8000000A 'grey
pic.Scale (0, 0)-(15, 15)
pic.Cls
pic.Line (0, 2)-(14, 14), BorderCol1, B
pic.Line (0.1, 2.1)-(14.1, 14.1), BorderCol2, B
For cmdI = 0 To Tot_Tabs
  LeftOf = cmdTab(cmdI).Left
  Wid = cmdTab(cmdI).Left + cmdTab(cmdI).Width
     pic.Line (LeftOf, 0)-(LeftOf, 2), BorderCol1
Next
  pic.Line (0, 0)-(Wid, 2), BorderCol1, B 'border of buttons
LeftOf = cmdTab(nIndex).Left
Wid = cmdTab(nIndex).Left + cmdTab(nIndex).Width
pic.Line (LeftOf, 0)-(LeftOf, 2), BorderCol1
pic.Line (LeftOf + 0.1, 0.1)-(LeftOf + 0.1, 2.1), BorderCol2
pic.Line (Wid, 0)-(Wid, 2), BorderCol1
pic.Line (Wid + 0.1, 0.1)-(Wid + 0.1, 2.1), BorderCol2
pic.Line (Wid - 0.05, 0.1)-(Wid - 0.05, 2.1), BorderCol3, BF
pic.Line (LeftOf + 0.15, 2)-(Wid, 2.2), BorderCol3, BF
End Sub


Sub Arrange_Tabs()
Dim i
For i = 0 To Tot_Tabs
cmdTab(i).Left = 1095 * i
cmdTab(i).Top = 0
TabPic(i).Left = 120
TabPic(i).Top = 600
Next
Show_Tab
End Sub

Sub Show_Tab(Optional ByVal cIndex = 0)
Dim i
For i = 0 To Tot_Tabs
TabPic(i).Visible = False
Next
TabPic(cIndex).Visible = True
End Sub


' Let the user browse for a directory. Return the
' selected directory. Return an empty string if
' the user cancels.
Private Function BrowseForDirectory(Bfrm As Form) As String
Dim browse_info As BrowseInfo
Dim item As Long
Dim dir_name As String
   
   browse_info.hWndOwner = Bfrm.hWnd
   browse_info.pidlRoot = 0
   browse_info.sDisplayName = Space$(260)
   browse_info.sTitle = "Select the Startup Directory"
   browse_info.ulFlags = 1 ' Return directory name.
   browse_info.lpfn = 0
   browse_info.lParam = 0
   browse_info.iImage = 0
   
   item = SHBrowseForFolder(browse_info)
   If item Then
       dir_name = Space$(260)
       If SHGetPathFromIDList(item, dir_name) Then
           BrowseForDirectory = Left(dir_name, InStr(dir_name, Chr$(0)) - 1)
       Else
           BrowseForDirectory = ""
       End If
   End If
End Function

