VERSION 5.00
Begin VB.Form frmPresetTiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preset Tiles"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox SelectTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   4320
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.VScrollBar VBar 
      Height          =   3975
      LargeChange     =   930
      Left            =   4920
      SmallChange     =   930
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Container 
      BackColor       =   &H80000009&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   120
      Width           =   4800
      Begin VB.PictureBox TileContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   0
         ScaleHeight     =   1050
         ScaleWidth      =   4800
         TabIndex        =   1
         Top             =   0
         Width           =   4800
         Begin VB.PictureBox picPreTile 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00404040&
            ForeColor       =   &H80000008&
            Height          =   810
            Index           =   0
            Left            =   120
            ScaleHeight     =   54
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   54
            TabIndex        =   2
            Tag             =   "Tile #1"
            Top             =   120
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmPresetTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nTiles As Integer 'variable for tilebox count


Private Sub cmds_Click(Index As Integer)
Select Case Index
Case 0
frmMain.picWork.Picture = SelectTile.Picture
Unload Me
Case 1
Unload Me
End Select
End Sub

Private Sub Form_Load()
Initialize_Tiles
End Sub

Private Sub picPreTile_DblClick(Index As Integer)
cmds_Click 0
End Sub

Private Sub picPreTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmds(0).Enabled = True
Draw_Borders
Make3DBorder picPreTile(Index), TileContainer, 2, 1, True
End Sub

Private Sub picPreTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picPreTile(Index).ToolTipText = picPreTile(Index).Tag
End Sub

Private Sub picPreTile_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SelectTile.Picture = LoadPicture()
SelectTile.Picture = picPreTile(Index).Image
End Sub

Private Sub VBar_Change()
VBar_Scroll
End Sub

Private Sub VBar_Scroll()
TileContainer.Top = -VBar.Value
End Sub


Sub Initialize_Tiles()
Load_TileBoxes
Add_Tiles
Draw_Borders
End Sub

Sub Draw_Borders()
Dim i As Integer
TileContainer.Picture = LoadPicture()
For i = 0 To nTiles - 1
Make3DBorder picPreTile(i), TileContainer, 1, 2, False
Next
End Sub

Sub Make3DBorder(Ctrl As Control, cParent As Control, nBevel%, nSpace%, Optional didSelect = False)
Dim PixX%, PixY%, CTop%, CLft%, CRgt%, CBtm%, i%
Dim AddX%, AddY%
Dim LinColor1, LinColor2

PixX% = Screen.TwipsPerPixelX
PixY% = Screen.TwipsPerPixelY
CTop% = Ctrl.Top - PixX%
CLft% = Ctrl.Left - PixY%
CRgt% = Ctrl.Left + Ctrl.Width
CBtm% = Ctrl.Top + Ctrl.Height

If didSelect = True Then
LinColor1 = &HC00000
LinColor2 = &HC00000
Else
LinColor1 = &HFFFFFF
LinColor2 = &H808080
End If

For i% = nSpace% To (nBevel% + nSpace% - 1)
AddX% = i% * PixX%
AddY% = i% * PixY%
cParent.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + AddX%, CTop% - AddY%), LinColor1 '&HFFFFFF
cParent.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - AddX%, CBtm% + AddY%), LinColor1 '&HFFFFFF
cParent.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + AddX% + PixX%, CBtm% + AddY%), LinColor2 '&H808080
cParent.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + AddX%, CBtm% + AddY%), LinColor2 '&H808080
Next
End Sub


Private Sub Load_TileBoxes()
    Dim tX As Single, tY As Single
    Dim StartX As Single, StartY As Single
    
    Dim MaximumX As Single, MaximumY As Single
    Dim TileWidth As Integer, TileHeight As Integer
    
    'first adjust the TileContainer height according to
    'the number of available tiles
    
    TileContainer.Height = ((TileRows(Total_Tiles - 1)) * (picPreTile(0).Height + 120))
    
    'proceed the process
    MaximumX = TileContainer.Width - picPreTile(0).Width
    MaximumY = TileContainer.Height - picPreTile(0).Height
    
    TileWidth = picPreTile(0).Width + 120
    TileHeight = picPreTile(0).Height + 120
    StartX = TileHeight + 120
    StartY = 120
    nTiles = 1
        
    For tY = StartY To MaximumY Step TileHeight
        For tX = StartX To MaximumX Step TileWidth
                Load picPreTile(nTiles)
                picPreTile(nTiles).Move tX, tY
                picPreTile(nTiles).Visible = True
                picPreTile(nTiles).Tag = "Tile #" & nTiles + 1
                nTiles = nTiles + 1
         DoEvents
        Next
        StartX = 120
      DoEvents
    Next
    
'set exact height of Tile container
TileContainer.Height = (picPreTile(nTiles - 1).Top + picPreTile(nTiles - 1).Height) + 120
'init the scroll bar
If TileContainer.Height > Container.Height Then
VBar.Max = TileContainer.Height - Container.Height
VBar.Enabled = True
Else
VBar.Enabled = False
End If
End Sub


Sub Add_Tiles()
Dim TileI
Dim Avail_Tiles
'call the tile count function
Avail_Tiles = (Total_Tiles) - 1 'dont include the first tile

For TileI = 0 To nTiles - 1
    If TileI < Avail_Tiles Then
        GetTile TileI + 102, picPreTile(TileI)
        picPreTile(TileI).Enabled = True
    Else
        GetTile 101, picPreTile(TileI)
        picPreTile(TileI).Enabled = False
    End If
Next TileI
End Sub

Sub GetTile(ByVal TileID As Integer, pic As PictureBox)
         pic.Picture = LoadResPicture(TileID, vbResBitmap)
         pic.Refresh
End Sub

'This function counts a set of bitmaps within a
'resource file attached to the project
'Call:      varName = Total_Tiles
Function Total_Tiles() As Single
Dim tmpPics() As Object
Dim Tilcount As Single
Dim Tcnt As Single

On Error GoTo handleCountERR

'intialize counts
Tcnt = 100
Tilcount = 0

Do
    Tcnt = Tcnt + 1 'initial counter
       ReDim tmpPics(Tcnt - 101) 'redimension array
       Set tmpPics(Tcnt - 101) = LoadResPicture(Tcnt, vbResBitmap) 'load
    Tilcount = Tilcount + 1 'count tiles
Loop
Exit Function

handleCountERR:
  Total_Tiles = Tilcount 'set the count
    'remove array of objects
    For Tcnt = 0 To UBound(tmpPics)
        Set tmpPics(Tcnt) = Nothing
    Next Tcnt
    Exit Function
End Function

Function TileRows(ByVal Total_Tiles As Single) As Single
Dim nRows As Single
Dim RowsInt As Single
On Error GoTo handleRowERR

nRows = Total_Tiles / 5
RowsInt = Int(nRows)

If nRows > RowsInt Then
RowsInt = RowsInt + 1
End If

TileRows = RowsInt
Exit Function

handleRowERR:
MsgBox Err.Description, vbCritical
TileRows = 0
Exit Function
End Function
