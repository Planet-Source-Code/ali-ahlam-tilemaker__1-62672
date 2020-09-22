VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmPresetTiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preset Tiles"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ControlBox      =   0   'False
   Icon            =   "frmPresetTiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPrintGal 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4335
      ScaleWidth      =   5055
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "Print &Setup"
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   3960
         Width           =   1215
      End
      Begin ComCtl2.UpDown GapBar 
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   4050
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327681
         Max             =   1
         Wrap            =   -1  'True
      End
      Begin VB.PictureBox collectionBox 
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3795
         ScaleWidth      =   4995
         TabIndex        =   11
         Top             =   0
         Width           =   5055
         Begin VB.VScrollBar BarV 
            Height          =   3735
            Left            =   4690
            TabIndex        =   13
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox picTileSet 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4455
            Left            =   0
            ScaleHeight     =   4455
            ScaleWidth      =   4560
            TabIndex        =   12
            Top             =   0
            Width           =   4560
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Spacing:"
         Height          =   195
         Left            =   3960
         TabIndex        =   14
         Top             =   4050
         Width           =   630
      End
   End
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmds 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
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
      Left            =   5280
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.VScrollBar VBar 
      Height          =   3855
      LargeChange     =   930
      Left            =   4920
      SmallChange     =   930
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Container 
      BackColor       =   &H80000004&
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
   Begin VB.CommandButton CmdPrintGal 
      Caption         =   "Print"
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "frmPresetTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nTiles As Integer 'variable for tilebox count
Private TilIndex As Integer 'selected tile to be deleted
Private GCancel As Boolean
Private MaxTileLimit As Integer

Property Get GalCancel() As Boolean
GalCancel = GCancel
End Property

Private Sub cmdPrintSet_Click()
On Error GoTo HandlePrintERR
frmMain.CMDLG.Flags = cdlPDPrintSetup
frmMain.CMDLG.PrinterDefault = True
frmMain.CMDLG.CancelError = True
frmMain.CMDLG.ShowPrinter
Exit Sub

HandlePrintERR:
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub cmds_Click(Index As Integer)
Select Case Index
Case 0 'OK
GCancel = False
With frmMain.picWork
.Width = SelectTile.Width
.Height = SelectTile.Height
.Picture = SelectTile.Picture
End With
Unload Me
Case 1 'delete
cmds(1).Enabled = False
Delete_Tile TilIndex
Case 2 'cancel
GCancel = True
Unload Me
End Select
End Sub



Private Sub Form_Load()
Initialize_Tiles galMode
Screen.MousePointer = 0
frmMain.ProgBar.Visible = False
End Sub

Private Sub picPreTile_DblClick(Index As Integer)
cmds_Click 0
End Sub

Private Sub picPreTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TilIndex = Index
cmds(0).Enabled = True
If galMode = 1 Then
cmds(1).Enabled = True
Else
cmds(1).Enabled = False
End If
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


Sub Initialize_Tiles(ByVal cMode As Integer)
Select Case cMode
Case 0 'default
Load_TileBoxes ((TileRows(Total_Tiles - 1)) * (picPreTile(0).Height + 120))
Add_Tiles
Case 1 'custom
Load_TileBoxes ((TileRows(Total_CustomTiles)) * (picPreTile(0).Height + 120))
AddCust_Tiles
End Select
Draw_Borders
End Sub



Private Sub Draw_Borders()
Dim i As Integer
TileContainer.Picture = LoadPicture()
For i = 0 To nTiles - 1
Make3DBorder picPreTile(i), TileContainer, 1, 2, False
Next
End Sub

Private Sub Make3DBorder(Ctrl As Control, cParent As Control, nBevel%, nSpace%, Optional didSelect = False)
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


Private Sub Load_TileBoxes(ByVal ContHeight As Single)
    Dim tX As Single, tY As Single
    Dim StartX As Single, StartY As Single
    
    Dim MaximumX As Single, MaximumY As Single
    Dim TileWidth As Integer, TileHeight As Integer
    
    'first adjust the TileContainer height according to
    'the number of available tiles
    Screen.MousePointer = 11
    frmMain.ProgBar.Visible = True
    frmMain.ProgBar.Cls
    
    TileContainer.Height = ContHeight '((TileRows(Total_Tiles - 1)) * (picPreTile(0).Height + 120))
    
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
      Update_Progress ((tY * 100) / MaximumY), "Initiaizing Gallery..."
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

Screen.MousePointer = 0
frmMain.ProgBar.Visible = False
End Sub


Private Sub Add_Tiles()
Dim TileI
Dim Avail_Tiles
'call the tile count function
Screen.MousePointer = 11
frmMain.ProgBar.Visible = True
frmMain.ProgBar.Cls

Avail_Tiles = (Total_Tiles) - 1 'dont include the first tile

For TileI = 0 To nTiles - 1
    If TileI < Avail_Tiles Then
        GetTile TileI + 102, picPreTile(TileI)
        picPreTile(TileI).Enabled = True
    Else
        GetTile 101, picPreTile(TileI)
        picPreTile(TileI).Enabled = False
    End If
    
Update_Progress ((TileI * 100) / (nTiles - 1)), "Loading Gallery..."
Next TileI
Screen.MousePointer = 0
frmMain.ProgBar.Visible = False
End Sub

Private Sub GetTile(ByVal TileID As Integer, pic As PictureBox)
         pic.Picture = LoadResPicture(TileID, vbResBitmap)
         pic.Refresh
End Sub

'This function counts a set of bitmaps within a
'resource file attached to the project
'Call:      varName = Total_Tiles
Private Function Total_Tiles() As Single
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

Private Function TileRows(ByVal Total_Tiles As Single) As Single
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

'add the custom tiles to the tile boxes
Private Sub AddCust_Tiles()
Dim CustI As Single
Dim P$

    Screen.MousePointer = 11
    frmMain.ProgBar.Visible = True
    frmMain.ProgBar.Cls

'add tiles
For CustI = 0 To nTiles - 1
    If (CustI + 1) <= Total_CustomTiles Then
        PICDB_Load Data_file, (CustI + 1), P$, picPreTile(CustI)
    Else
        GetTile 101, picPreTile(CustI)
        picPreTile(CustI).Enabled = False
    End If
 Update_Progress ((CustI * 100) / (nTiles - 1)), "Loading Custom Tiles..."
Next CustI

    Screen.MousePointer = 0
    frmMain.ProgBar.Visible = False
End Sub


Private Sub Delete_Tile(ByVal TileIndex As Integer)
Dim DelI, Tnumber As Integer
Dim tmpPic As String
Dim tmpGalry As String

'define files
tmpPic = App.Path & "\~tSPic.bmp"
tmpGalry = App.Path & "\~tGal.DAT"

Open tmpGalry For Output As #1 'open the tmp file
Close #1

'Update the gallery without the selected tile
For DelI = 0 To (Total_CustomTiles - 1)
If (DelI) = TileIndex Then
    picPreTile(DelI).Picture = LoadPicture()
Else
    SavePicture picPreTile(DelI).Image, tmpPic 'save tmp pic
    Tnumber = Total_CustomTiles + 1
    PicDB_Add tmpGalry, tmpPic, ("Tile#" & Tnumber) 'save dat
    Kill tmpPic 'delete tmp
End If
Next DelI

'update files
Kill Data_file 'delete old Gallery
Name tmpGalry As Data_file 'create new

'refresh tile box
AddCust_Tiles
Draw_Borders
cmds(0).Enabled = False
End Sub

'=============================================printing
Private Sub GapBar_Change()
picTileSet.Height = (9 * 825)
picTileSet.Width = (5 * 825)
If galMode = 0 Then
MaxTileLimit = Total_Tiles - 1 'default
Else
MaxTileLimit = Total_CustomTiles - 1 'custom
End If
PrintTile_Gallery picTileSet, MaxTileLimit, GapBar.Value
BarV.Max = (picTileSet.Height - collectionBox.ScaleHeight) + 255
End Sub


Private Sub BarV_Change()
 BarV_Scroll
End Sub

Private Sub BarV_Scroll()
picTileSet.Top = -BarV.Value
End Sub

Private Sub cmdCancel_Click()
picPrintGal.Visible = False
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
PrintOutGallery picTileSet
picPrintGal.Visible = False
Screen.MousePointer = 0
End Sub

Private Sub CmdPrintGal_Click()
picTileSet.Height = 9 * 825
picTileSet.Width = 5 * 825
If galMode = 0 Then
MaxTileLimit = Total_Tiles - 1 'default
Else
MaxTileLimit = Total_CustomTiles - 1 'custom
End If
PrintTile_Gallery picTileSet, MaxTileLimit, GapBar.Value
BarV.Max = (picTileSet.Height - collectionBox.ScaleHeight) + 255
picPrintGal.Visible = True
End Sub

Private Sub PrintTile_Gallery(Targetobj As Object, ByVal Ttotals As Integer, ByVal gapSize As Integer)
Dim Wid As Single
Dim Hgt As Single
Dim SWid As Single
Dim SHgt As Single

Dim X As Single
Dim Y As Single
Dim cTile As Integer
Dim maxX As Single
Dim maxY As Single


    SWid = picPreTile(0).ScaleWidth
    SHgt = picPreTile(0).ScaleHeight

    Wid = picPreTile(cTile).Width + gapSize
    Hgt = picPreTile(cTile).Height + gapSize
    
    Y = 0
    cTile = 0
    
    maxX = (5 * SWid)
    maxY = (9 * SHgt)
    
    Targetobj.Picture = LoadPicture()
    Targetobj.ScaleMode = vbPixels
    
    Do While Y < maxY 'Targetobj.ScaleHeight
        X = 0
        Do While X < maxX 'Targetobj.ScaleWidth
            Targetobj.PaintPicture picPreTile(cTile).Picture, _
                X, Y, SWid, SHgt
              X = X + SWid + gapSize
             If cTile = Ttotals Then GoTo DoneAll
            cTile = cTile + 1
        Loop
        Y = Y + SHgt + gapSize
    Loop
    
DoneAll:
    Targetobj.Refresh
  Targetobj.Picture = Targetobj.Image
Exit Sub
End Sub


'prints the gallery
Private Sub PrintOutGallery(picDest As PictureBox)
Dim CCopies As Integer

On Error GoTo HandlePrintERR

Printer.Scale (0, 0)-(7, 11)
picDest.Scale (0, 0)-(7, 11)

Printer.PaintPicture picDest.Picture, 1, 1
Printer.EndDoc 'start print
Exit Sub

HandlePrintERR:
MsgBox Err.Description, vbCritical
Printer.EndDoc
Exit Sub
End Sub


