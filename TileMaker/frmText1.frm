VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ControlBox      =   0   'False
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   4215
      Begin VB.CommandButton cmds 
         Caption         =   "&Font"
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblfont 
         AutoSize        =   -1  'True
         Caption         =   "Font:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter your text here"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmds_Click(Index As Integer)
Select Case Index
Case 0 'OK
With txtEntry
.Text = Trim(txtEntry.Text)
cmds(0).Enabled = .Text <> ""
If .Text = "" Then Exit Sub
ctxtEnt = .Text
ctxtFont = .Font.Name
ctxtSize = .Font.Size
ctxtBold = .Font.Bold
ctxtItalic = .Font.Italic
ctxtStrike = .Font.Strikethrough
ctxtUline = .Font.Underline
End With
Unload Me
Case 1 'font
On Error GoTo HandleFontErr
With frmMain.CMDLG
.CancelError = True
.Flags = cdlCFApply Or cdlCFBoth Or cdlCFEffects
.FontName = txtEntry.Font.Name
.FontSize = txtEntry.Font.Size
.ShowFont
txtEntry.Font.Name = .FontName
txtEntry.Font.Size = .FontSize
txtEntry.Font.Bold = .FontBold
txtEntry.Font.Italic = .FontItalic
txtEntry.Font.Strikethrough = .FontStrikethru
txtEntry.Font.Underline = .FontUnderline
End With
lblfont.Caption = "Font: " & txtEntry.Font.Name & ", Size: " & Abs(txtEntry.Font.Size)
Exit Sub
HandleFontErr:
Exit Sub
Case 2 'cancel
Unload Me
End Select
End Sub

Private Sub Form_Load()
ctxtEnt = ""
ctxtFont = ""
ctxtSize = 0
ctxtBold = False
ctxtItalic = False
ctxtStrike = False
ctxtUline = False
txtEntry.Font.Name = txtEntry.FontName
txtEntry.Font.Size = txtEntry.FontSize
txtEntry.Font.Bold = txtEntry.FontBold
txtEntry.Font.Italic = txtEntry.FontItalic
txtEntry.Font.Strikethrough = txtEntry.FontStrikethru
txtEntry.Font.Underline = txtEntry.FontUnderline

lblfont.Caption = "Font: " & txtEntry.Font.Name & ", Size: " & Abs(txtEntry.Font.Size)
End Sub

Private Sub txtEntry_Change()
cmds(0).Enabled = txtEntry.Text <> ""
End Sub

Private Sub txtEntry_LostFocus()
txtEntry.Text = Trim(txtEntry.Text)
cmds(0).Enabled = txtEntry.Text <> ""
End Sub
