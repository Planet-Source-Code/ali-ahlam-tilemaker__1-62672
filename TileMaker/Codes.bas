Attribute VB_Name = "Codes"
Option Explicit
#If Win32 Then
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

'help **********

Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long

Public Const HELP_QUIT = 2
Public Const HELP_INDEX = 3
Public Const HELP_HELPONHELP = 4
Public Const HELP_PARTIALKEY = &H105
Public Const HELP_CONTENTS = &H3&
'***************

Private Const SWW_HPARENT = -8

#Else
End
#End If

'Program vars
Public DrawSize As Integer
Public cStyle(0 To 162) As Integer 'styles array
Public nObject As Integer 'style boxes
Public TotColrs As Integer 'colors used
Public pColMode As Integer 'color mode
Public cPatColor  As Long 'color selected
Public cTool As Integer 'current selected tool


Public cBrush As Integer
Public BcRad As Integer
Public BbSiz As Integer
Public BrSiz As Integer
Public IniName As String

Public ctxtEnt As String 'text entry
Public ctxtFont As String 'text font
Public ctxtSize As Integer 'text size
Public ctxtBold As Boolean 'text bold
Public ctxtItalic As Boolean 'text italic
Public ctxtStrike As Boolean 'text strikethru
Public ctxtUline As Boolean 'text underline

'text dialog
Public cFontName As String
Public cFontSize As Integer
Public cFontBold As Boolean
Public cFontItalic As Boolean
Public cFontUnderline As Boolean
Public cFontStrikeThru As Boolean
Public cFontColor As Long
Public firstTimetxt As Boolean

Public OffCancel As Boolean 'offset cancel flag
Public Const Max_Tool = 14 'tools used
Public Max_Undo  As Integer '= 23 '24 boxes'undo used max
Public Max_Redo  As Integer '= 23 '24 boxes'redo used max
Public maxLevel As Integer 'selected total of Undo/Redo
Public Anim As Integer


Public tmpFName As String 'for tmp saving
Public Data_file As String ' Data file name
Public galMode As Integer 'gallery Mode
Public TMHelpFile As String



'provide help
Public Sub HelpFunction(lhWnd As Long, HelpCmd As Integer, HelpKey As String)
   
Dim lRtn As Long 'declare the needed variables
   
If HelpCmd = HELP_PARTIALKEY Then
   lRtn = WinHelp(lhWnd, TMHelpFile, HelpCmd, HelpKey)
Else
   lRtn = WinHelp(lhWnd, TMHelpFile, HelpCmd, 0&)
End If
   
End Sub



Sub Open_Pic()
On Error GoTo HandlePicOpenErr
With frmMain.CMDLG
.Filename = ""
.Filter = "Bitmap files (*.bmp)|*.bmp|JPEG files (*.jpg)|*.jpg| GIFF files (*.gif)|*.gif| Icon files (*.ico)|*.ico| Cursor files (*.cur)|*.cur| Meta files (*.wmf)|*.wmf| PCX files (*.pcx)|*.pcx| All files (*.*)|*.*"
.CancelError = False
.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
.DialogTitle = "Select the Picture file"
.ShowOpen
End With
If frmMain.CMDLG.Filename = "" Then Exit Sub
Screen.MousePointer = 11
With frmMain
.picWork.Picture = LoadPicture(frmMain.CMDLG.Filename)
.picWork.Tag = frmMain.CMDLG.Filename 'set the filename tag
.picTmp.Picture = LoadPicture(frmMain.CMDLG.Filename)
.picWork.Refresh
.picTmp.Refresh
.picWork.Picture = .picWork.Image
.picTmp.Picture = .picWork.Image

frmMain.OpenFile frmMain.CMDLG.Filename


frmMain.Resize_UndoRedoPics
frmMain.Save_UndoAction 'undo action

End With
Screen.MousePointer = 0
Exit Sub

HandlePicOpenErr:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub


Sub Save_Pic(pic As PictureBox)
Dim curName As String

curName = pic.Tag

If curName = "" Then 'is a new one
Save_PicAs pic
Else 'saved already
SavePicture pic.Image, curName
End If
Exit Sub

SaveErr:
Screen.MousePointer = 0
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

Sub Save_PicAs(pic As PictureBox)
On Error GoTo SaveErr
With frmMain.CMDLG
.Filename = ""
.Filter = "Bitmap files (*.bmp)|*.bmp|JPEG files (*.jpg)|*.jpg| GIFF files (*.gif)|*.gif| Icon files (*.ico)|*.ico| Cursor files (*.cur)|*.cur| Meta files (*.wmf)|*.wmf| PCX files (*.pcx)|*.pcx"
.CancelError = False
.Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt Or cdlOFNNoChangeDir Or cdlOFNNoLongNames
.DialogTitle = "Save File As..."
.ShowSave
End With
If frmMain.CMDLG.Filename = "" Then Exit Sub
Screen.MousePointer = 11
With frmMain
SavePicture pic.Image, .CMDLG.Filename
pic.Tag = .CMDLG.Filename
End With
Screen.MousePointer = 0
Exit Sub

SaveErr:
Screen.MousePointer = 0
MsgBox Err.Description, vbCritical
Exit Sub
End Sub


Sub Main()
On Error GoTo handleLoadERR
'frmPicSel.Show
'frmPatSet.Show
'frmLens.Show
'frmPresetTiles.Show
'frmAddNoise.Show
frmSplash.Show
PauseFor 2
If Dir(CStr(GetIniVal("Gallery", "CustomGallery", IniName))) = "" Then
Data_file = App.Path & "\CGallery.DAT" 'custom gallery name
Else
Data_file = (GetIniVal("Gallery", "CustomGallery", IniName))
End If

TMHelpFile = App.Path & "\TileMaker.hlp" 'TileMaker Help file
IniName = App.Path & "\TileMaker.ini" 'Ini File

If Not Dir(CStr(GetIniVal("General", "StartupDir", IniName)), vbDirectory) = "" Then
    ChDir (CStr(GetIniVal("General", "StartupDir", IniName))) 'change dir
End If
frmMain.Show
Unload frmSplash
Exit Sub
handleLoadERR:
MsgBox "There's not enough memory to load TileMaker" & vbCrLf & _
"Please try again after freeing some memory." & vbCrLf & _
"If the problem persists try restarting windows.", vbCritical
End
End Sub

'get value
Function GetIniVal(ByVal cSection As String, ByVal cKey As String, ByVal IniFile As String)
Dim buf As String * 256
Dim length As Long

On Error Resume Next

Screen.MousePointer = 11
    length = GetPrivateProfileString( _
        cSection, cKey, "<no value>", _
        buf, Len(buf), IniFile)
    GetIniVal = Left$(buf, length)
Screen.MousePointer = 0
Exit Function
End Function

' Set the value.
Sub SetIniVal(ByVal cSection As String, ByVal cKey As String, ByVal cValue As String, ByVal IniFile As String)
On Error Resume Next

Screen.MousePointer = 11
    WritePrivateProfileString _
        cSection, cKey, _
        cValue, IniFile
Screen.MousePointer = 0
Exit Sub
End Sub

Sub PauseFor(ByVal nSecs As Integer)
Dim start
    start = Timer   ' Set start time.
    Do While Timer < start + nSecs
        DoEvents    ' Yield to other processes.
    Loop
End Sub

Public Function UserName() As String
Dim Bufstr As String
On Error Resume Next
Bufstr = Space$(50)
If GetUserName(Bufstr, 50) > 0 Then
    UserName = Bufstr
    UserName = RTrim(UserName)
Else
    UserName = ""
End If
Exit Function
End Function


'set an objects size in Inches
'Call:-    SetSizeInInches anyobject, CSng(3), CSng(1)
Public Sub SetSizeInInches(ByVal obj As Object, ByVal Wid As Single, ByVal Hgt As Single)
On Error Resume Next
    ' Convert into twips. Add room for the border.
    obj.Width = Wid * 1440 + obj.Width - _
        obj.ScaleX(obj.ScaleWidth, obj.ScaleMode, vbTwips)
    obj.Height = Hgt * 1440 + obj.Height - _
        obj.ScaleY(obj.ScaleHeight, obj.ScaleMode, vbTwips)
Exit Sub
End Sub

'makes a form float on another form
Sub Float_Form(FloatForm As Form, mainFrm As Form)
On Error Resume Next
    Call SetWindowWord(FloatForm.hWnd, SWW_HPARENT, mainFrm.hWnd)
Exit Sub
End Sub
 


'save the tile to dat
Public Sub PicDB_Add(Filename As String, Bitmapname As String, Caption As String)
Dim X, Y, NOR%, r, E%, O&, l$
    'Procces data
    X = FreeFile
    Open Filename For Binary As #X
    If LOF(X) = 0 Then GoTo npos
    Get #X, 1, NOR%
npos:
    NOR% = NOR% + 1
    Put #X, 1, NOR%
    r = LOF(X) + 1
    E% = Len(Caption)
    Put #X, r, E%
    Put #X, , Caption
    Y = FreeFile
    Open Bitmapname For Binary As #Y
    O& = LOF(Y)
    Put #X, , O&
    l$ = String$(O&, " ")
    Get #Y, , l$
    Put #X, , l$
    Close #Y
    Close #X
End Sub

'load a custom tile
Public Sub PICDB_Load(Filename As String, ByVal Recnum As Single, Caption As String, pic As Control)
Dim X, E%, i, r$, O&, Y
Dim tmpDBFile As String

    X = FreeFile
    tmpDBFile = App.Path & "\¤BMPDB¤.TMP"
    Open Filename For Binary As #X
    Get #X, 1, E%
    For i = 1 To (Recnum - 1)
    Get #X, , E%
    r$ = String$(E%, " ")
    Get #X, , r$
    Get #X, , O&
    r$ = String$(O&, " ")
    Get #X, , r$
    Next
    Get #X, , E%
    Caption = String$(E%, " ")
    Get #X, , Caption
    Get #X, , O&
    r$ = String$(O&, " ")
    Get #X, , r$
    Y = FreeFile
    Open tmpDBFile For Output As #Y
    Print #Y, r$
    Close #Y
    Close #X
    pic.Picture = LoadPicture(tmpDBFile)
    pic.Refresh
    Kill tmpDBFile
End Sub


'count the tiles in custom gallery
Public Function Total_CustomTiles() As Single
Dim Tot_Entries As Integer

' Check for Records
On Error GoTo HandleReadDatERR
Open Data_file For Binary As #1
    If LOF(1) = 0 Then
    Total_CustomTiles = 0
Close #1
Exit Function ' If no data Close File, and exit with 0
End If

Get #1, 1, Tot_Entries
'total entries
Total_CustomTiles = Tot_Entries
Close #1
Exit Function

HandleReadDatERR:
'MsgBox Err.Description, vbCritical
    Total_CustomTiles = 0
    Exit Function
End Function


