Attribute VB_Name = "BmpCodes"
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWid As Integer, ByVal nHt As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function SetPixel Lib "GDI32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "GDI32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer) As Long
Declare Function StretchBlt% Lib "GDI32" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop&)
Declare Function FloodFill Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Declare Function Rectangle Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function Ellipse Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function Chord Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function Arc Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function ArcTo Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function TextOut Lib "GDI32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long


'BmpFlip, BmpMirror, BmpRotate
Const pi = 3.14159265359

Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)


Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

'Public Aspect

Public Enum SpecialFilters
    fEmboss
    fEngrave
    fMotionBlur
End Enum


Public Sub BmpMirror(Pict1 As PictureBox, pict2 As PictureBox)
Dim px%
Dim py%
 
 On Error GoTo HandleErr
 
    'flip horizontal
    Pict1.ScaleMode = 3
    pict2.ScaleMode = 3

    pict2.Picture = LoadPicture()
    px% = Pict1.ScaleWidth
    py% = Pict1.ScaleHeight
    RetVal% = StretchBlt(pict2.hDC, px% - 1, 0, -px%, py%, Pict1.hDC, 0, 0, px%, py%, SRCCOPY)
    pict2.Refresh
    pict2.Picture = pict2.Image
    Pict1.Picture = pict2.Image
    Exit Sub
    
HandleErr:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub


'This is a revised edition of the bmpTile which I
'collected and changed a litle bit This one is much more
'efficient than the previous one which had the
'problem of over flow. Thanks to  www.vb-helper.com

Sub BmpTile(Targetobj As Object, picTile As PictureBox)
Dim Wid As Single
Dim Hgt As Single
Dim X As Single
Dim Y As Single

On Error GoTo HandleTileErr 'just in case

    Wid = picTile.ScaleWidth
    Hgt = picTile.ScaleHeight
    Y = 0
    
    Targetobj.Picture = LoadPicture()
    frmMain.stbar.Text = "Processing Tile..."
    frmMain.ProgBar.Visible = True
    Do While Y < Targetobj.ScaleHeight
        X = 0
        Do While X < Targetobj.ScaleWidth
            Targetobj.PaintPicture picTile.Picture, _
                X, Y, Wid, Hgt
            X = X + Wid
        Loop
        Y = Y + Hgt
        Update_Progress ((Y * 100) / Hgt), "Processing..."
    Loop
  
    Targetobj.Refresh
  Targetobj.Picture = Targetobj.Image
frmMain.stbar.Text = "Ready."
frmMain.ProgBar.Visible = False
Exit Sub
HandleTileErr:
frmMain.stbar.Text = "Ready"

Exit Sub
End Sub

'This is a revised edition of the bmpFlip which I
'changed a litle bit this procedure is fixed with
'the problem of a distorted image caused before with
'the old procedure.
Public Sub BmpFlip(Pict1 As PictureBox, pict2 As PictureBox, ByVal FlipBy As String)
Dim px%
Dim py%
    
On Error GoTo HandleErr
    'flip
    Pict1.ScaleMode = 3
    pict2.ScaleMode = 3
    
    pict2.Picture = LoadPicture()
    px% = Pict1.ScaleWidth
    py% = Pict1.ScaleHeight
    
    If UCase(FlipBy) = "H" Then
    pict2.PaintPicture Pict1.Picture, _
                0, py% - 1, px%, -py%, 0, 0, px%, py%, SRCCOPY
    Else
    pict2.PaintPicture Pict1.Picture, _
               px% - 1, 0, -px%, py%, 0, 0, px%, py%, SRCCOPY
    End If
   
    pict2.Refresh
    pict2.Picture = pict2.Image
    Pict1.Picture = pict2.Image
Exit Sub

HandleErr:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

'draws a picture in a preview form
Sub Paint_ClipSize(picParent As PictureBox, picClip As PictureBox)
Dim to_x As Single
Dim to_y As Single
Dim Wid As Single
Dim Hgt As Single

On Error GoTo HandleErr

    If picParent.Picture = 0 Then Exit Sub
   With picClip
   .Picture = LoadPicture()
    
    ' See if the image is too big to fit.
    Wid = picParent.ScaleWidth
    Hgt = picParent.ScaleHeight
    If Wid > .ScaleWidth Then
        Hgt = Hgt * .ScaleWidth / Wid
        Wid = .ScaleWidth
    End If
    If Hgt > .ScaleHeight Then
        Wid = Wid * .ScaleHeight / Hgt
        Hgt = .ScaleHeight
    End If

    ' See where we need to put the picture to center it.
    to_x = (.ScaleWidth - Wid) / 2
    to_y = (.ScaleHeight - Hgt) / 2

    ' Copy the picture centered on the form.
    .PaintPicture picParent.Picture, _
        to_x, to_y, Wid, Hgt
    .Refresh
    .Picture = .Image
    End With
Exit Sub

HandleErr:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

'negative form of image
Sub bmpNegative(picFrom As PictureBox, picTo As PictureBox)
Dim Wid, Hgt As Single
On Error Resume Next
    Wid = picFrom.ScaleWidth
    Hgt = picFrom.ScaleHeight
    picTo.Picture = LoadPicture()
     picTo.PaintPicture picFrom, 0, 0, Wid, Hgt, 0, 0, Wid, Hgt, SRCINVERT 'negative
    picTo.Refresh
    picTo.Picture = picTo.Image
    picFrom.Picture = picTo.Image
    Exit Sub
End Sub

'pattern
Sub Draw_Pattern(pic As PictureBox, ByVal shap, ByVal siz)
Dim w, h, r, X, Y
   w = pic.ScaleWidth / 2
   h = pic.ScaleHeight / 2
     pic.Cls
 On Error GoTo HandleErr
     
        For i = 1 To 360
            angle = (i * 3.141592654) / 180
            r = siz * Cos(shap * angle)
            X = r * Sin(angle)
            Y = r * Cos(angle)
          pic.PSet (X + w, Y + h), QBColor(i Mod 15)
          pic.Refresh
         DoEvents
        Next i
     pic.Picture = pic.Image
Exit Sub


HandleErr:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

'draws a lens
Sub Draw_Lens(pLens As PictureBox, pImage As PictureBox)
On Error Resume Next
pLens.Picture = LoadPicture()
pLens.PaintPicture pImage.Picture _
, 0, 0, pLens.Width, pLens.Height
Exit Sub
End Sub

'draws a pattern with a selected color
Sub Pattern_SColor(pic As PictureBox, ByVal numSteps, ByVal cselColr As Long)
Dim pScaleMode, i

 On Error GoTo HandleErr

  If numSteps = "" Then Exit Sub
   If numSteps <= 0 Then Exit Sub
    pScaleMode = pic.ScaleMode
    pic.ScaleMode = vbTwips
    pic.Picture = LoadPicture()
    pic.FillStyle = 1
    For i = 0 To pic.Width Step pic.Width / numSteps
    If cStyle(0) = 1 Then
        pic.Line (i, pic.Height)-(0, i), cselColr '3D effect
    End If
    If cStyle(1) = 1 Then
         pic.Line (i, 0)-(pic.Height, i), cselColr '3D effect
    End If
    If cStyle(2) = 1 Then
    pic.Line (0, pic.Height - i)-(i, 0), cselColr '3D effect
    End If
    If cStyle(3) = 1 Then
    pic.Line (pic.Height, pic.Height - i)-(i, pic.Height), cselColr  '3D effect
    End If
    If cStyle(4) = 1 Then
         pic.Line (i, pic.Height)-(0, 0), cselColr 'topleft
    End If
    If cStyle(5) = 1 Then
         pic.Line (0, 0)-(pic.Height, i), cselColr 'topleft
    End If
    If cStyle(6) = 1 Then
         pic.Line (pic.Height, pic.Height)-(i, 0), cselColr 'bottomleft
    End If
    If cStyle(7) = 1 Then
         pic.Line (pic.Height, pic.Height)-(0, i), cselColr 'bottomleft
    End If
    If cStyle(8) = 1 Then
         pic.Line (0, i)-(pic.Height, 0), cselColr 'topright
    End If
    If cStyle(9) = 1 Then
         pic.Line (i, pic.Height)-(pic.Height, 0), cselColr 'topright
    End If
    If cStyle(10) = 1 Then
         pic.Line (i, 0)-(0, pic.Height), cselColr 'bottomright
    End If
    If cStyle(11) = 1 Then
         pic.Line (pic.Height, i)-(0, pic.Height), cselColr 'bottomright
    End If
    If cStyle(12) = 1 Then
         pic.Line (0, i)-(pic.Height, i), cselColr 'horz
    End If
    If cStyle(13) = 1 Then
        pic.Line (i, pic.Height)-(i, 0), cselColr 'vert
    End If
    If cStyle(14) = 1 Then
       pic.Line (0, i)-(i, 0), cselColr 'mesh1
    End If
    If cStyle(15) = 1 Then
        pic.Line (i, pic.Height)-(pic.Height, i), cselColr 'mesh2
    End If
    If cStyle(16) = 1 Then
        pic.Line (pic.Height - i, 0)-(pic.Height, i), cselColr 'mesh3
    End If
    If cStyle(17) = 1 Then
        pic.Line (i, pic.Height)-(0, pic.Height - i), cselColr  'mesh4
    End If
    If cStyle(18) = 1 Then
        pic.Line (0, i)-(pic.Height - i, pic.Height - i), cselColr '3D effect 1
    End If
    If cStyle(19) = 1 Then
        pic.Line (pic.Height - i, pic.Height - i)-(i, 0), cselColr '3D effect 2
    End If
    If cStyle(20) = 1 Then
        pic.Line (pic.Height - i, 0)-(i, pic.Height - i), cselColr '3D effect 3
    End If
    If cStyle(21) = 1 Then
        pic.Line (i, pic.Height - i)-(pic.Height, i), cselColr '3D effect 4
    End If
    If cStyle(22) = 1 Then
        pic.Line (pic.Height - i, i)-(0, pic.Height - i), cselColr '3D effect 5
    End If
    If cStyle(23) = 1 Then
        pic.Line (i, pic.Height)-(pic.Height - i, i), cselColr '3D effect 6
    End If
    If cStyle(24) = 1 Then
        pic.Line (pic.Height - i, pic.Height - i)-(i, pic.Height), cselColr '3D effect 7
    End If
    If cStyle(25) = 1 Then
        pic.Line (pic.Height, i)-(pic.Height - i, pic.Height - i), cselColr '3D effect 8
    End If
    If cStyle(26) = 1 Then
        pic.Line (pic.Height - i, pic.Height - i)-(i, pic.Height - i), cselColr 'effect1
    End If
    If cStyle(27) = 1 Then
        pic.Line (pic.Height - i, i)-(pic.Height - i, pic.Height - i), cselColr 'effect2
    End If
    If cStyle(28) = 1 Then '
        pic.Line (pic.Height, 0)-(pic.Height - i, pic.Height - i), cselColr 'box effect
    End If
    If cStyle(29) = 1 Then '
        pic.Line (pic.Height - i, pic.Height - i)-(0, pic.Height), cselColr 'box effect
    End If
    If cStyle(30) = 1 Then
        pic.Line (0, 0)-(i, pic.Height - i), cselColr 'box effect
    End If
    If cStyle(31) = 1 Then
        pic.Line (pic.Height - i, i)-(pic.Height, pic.Height), cselColr 'box effect
    End If
    If cStyle(32) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height - i, i), cselColr '3D
    End If
    If cStyle(33) = 1 Then
         pic.Line (i, i)-(pic.Height / 2, pic.Height - i), cselColr '3D
    End If
    If cStyle(34) = 1 Then
        pic.Line (0, i)-(pic.Height, pic.Height - i), cselColr 'effect5
    End If
    If cStyle(35) = 1 Then
        pic.Line (pic.Height - i, 0)-(i, pic.Height), cselColr  'effect6
    End If
    If cStyle(36) = 1 Then
        pic.Line (i, pic.Height)-(i, i), cselColr   'line^1
    End If
    If cStyle(37) = 1 Then
        pic.Line (i, pic.Height - i)-(i, pic.Height), cselColr  'line^2
    End If
    If cStyle(38) = 1 Then
        pic.Line (i, 0)-(i, pic.Height - i), cselColr 'lineV1
    End If
    If cStyle(39) = 1 Then
        pic.Line (i, 0)-(i, i), cselColr   'lineV2
    End If
    If cStyle(40) = 1 Then
        pic.Line (pic.Height, i)-(i, i), cselColr  'line<1
    End If
    If cStyle(41) = 1 Then
        pic.Line (pic.Height, i)-(pic.Height - i, i), cselColr 'line<2
    End If
    If cStyle(42) = 1 Then
        pic.Line (0, i)-(pic.Height - i, i), cselColr 'line>1
    End If
    If cStyle(43) = 1 Then
        pic.Line (0, i)-(i, i), cselColr 'line>2
    End If
    If cStyle(44) = 1 Then
        pic.Line (pic.Height - i / 2, i)-(i, pic.Height - i), cselColr '3D
    End If
    If cStyle(45) = 1 Then
        pic.Line (i, (pic.Height / 2) - i / 2)-(pic.Height - i, i), cselColr '3D
    End If
    If cStyle(46) = 1 Then
        pic.Line (pic.Height - i, i)-(i, pic.Height - i / 2), cselColr '3D
    End If
    If cStyle(47) = 1 Then
        pic.Line (i, pic.Height - i)-((pic.Height / 2) - i / 2, i), cselColr '3D
    End If
    If cStyle(48) = 1 Then
        pic.Line (pic.Height - i, pic.Height - i / 2)-(i, i), cselColr '3D
    End If
    If cStyle(49) = 1 Then
        pic.Line (i, i)-(pic.Height - i / 2, pic.Height - i), cselColr '3D
    End If
    If cStyle(50) = 1 Then
        pic.Line ((pic.Height / 2) - i / 2, pic.Height - i)-(i, i), cselColr '3D
    End If
    If cStyle(51) = 1 Then
        pic.Line (i, i)-(pic.Height - i, (pic.Height / 2) - i / 2), cselColr '3D
    End If
    If cStyle(52) = 1 Then
        pic.Line (pic.Height / 2, i)-(i, pic.Height / 2), cselColr  'SlantBox
    End If
    If cStyle(53) = 1 Then
        pic.Line (pic.Height - i, pic.Height / 2)-(pic.Height / 2, i), cselColr 'SlantBox
    End If
    If cStyle(54) = 1 Then
        pic.Line (pic.Height - i, i)-(0, pic.Height - i / 2), cselColr '3D
    End If
    If cStyle(55) = 1 Then
        pic.Line (pic.Height / 2 - i / 2, pic.Height)-(i, pic.Height - i), cselColr '3D
    End If
    If cStyle(56) = 1 Then
        pic.Line (pic.Height - i / 2, 0)-(i, pic.Height - i), cselColr '3D
    End If
    If cStyle(57) = 1 Then
        pic.Line (pic.Height, pic.Height / 2 - i / 2)-(pic.Height - i, i), cselColr  '3D
    End If
    If cStyle(58) = 1 Then
        pic.Line (pic.Height, pic.Height - i / 2)-(i, i), cselColr '3D effect 8
    End If
    If cStyle(59) = 1 Then
        pic.Line (pic.Height - i / 2, pic.Height)-(i, i), cselColr '3D effect 8
    End If
    If cStyle(60) = 1 Then
        pic.Line (0, pic.Height / 2 - i / 2)-(i, i), cselColr '3D effect 8
    End If
    If cStyle(61) = 1 Then
        pic.Line (pic.Height / 2 - i / 2, 0)-(i, i), cselColr '3D effect 8
    End If
    If cStyle(62) = 1 Then
        pic.Line (pic.Height / 2, i)-(0, pic.Height / 2), cselColr 'Box
    End If
    If cStyle(63) = 1 Then
        pic.Line (pic.Height / 2, pic.Height - i)-(pic.Height, pic.Height / 2), cselColr 'Box
    End If
    If cStyle(64) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height / 2, 0), cselColr  'Box
    End If
    If cStyle(65) = 1 Then
        pic.Line (pic.Height / 2, pic.Height)-(i, pic.Height / 2), cselColr 'Box
    End If
    If cStyle(66) = 1 Then
        pic.Line (pic.Height - i, 0)-(i, pic.Height / 2), cselColr '3D
    End If
    If cStyle(67) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height - i, pic.Height), cselColr '3D
    End If
    If cStyle(68) = 1 Then
        pic.Line (0, i)-(pic.Height / 2, pic.Height - i), cselColr  '3D
    End If
    If cStyle(69) = 1 Then
        pic.Line (pic.Height, i)-(pic.Height / 2, pic.Height - i), cselColr '3D
    End If
    If cStyle(70) = 1 Then
        pic.Line (i, pic.Height / 2)-(0, 0), cselColr  'trianglestyle
    End If
    If cStyle(71) = 1 Then
        pic.Line (0, pic.Height)-(i, pic.Height / 2), cselColr  'trianglestyle
    End If
    If cStyle(72) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height, 0), cselColr 'trianglestyle
    End If
    If cStyle(73) = 1 Then
        pic.Line (pic.Height, pic.Height)-(i, pic.Height / 2), cselColr  'trianglestyle
    End If
    If cStyle(74) = 1 Then
        pic.Line (0, pic.Height / 2)-(i, pic.Height), cselColr 'trianglestyle
    End If
    If cStyle(75) = 1 Then
        pic.Line (i, 0)-(0, pic.Height / 2), cselColr  'trianglestyle
    End If
    If cStyle(76) = 1 Then
        pic.Line (pic.Height, pic.Height / 2)-(i, pic.Height), cselColr  'trianglestyle
    End If
    If cStyle(77) = 1 Then
        pic.Line (i, 0)-(pic.Height, pic.Height / 2), cselColr  'trianglestyle
    End If
    If cStyle(78) = 1 Then
        pic.Line (pic.Height / 2, i)-(pic.Height - i / 2, pic.Height), cselColr 'Isocelesstyle
    End If
    If cStyle(79) = 1 Then
        pic.Line (pic.Height / 2, pic.Height - i)-(pic.Height / 2 - i / 2, pic.Height), cselColr 'Isocelesstyle
    End If
    If cStyle(80) = 1 Then
        pic.Line (pic.Height - i / 2, 0)-(pic.Height / 2, pic.Height - i), cselColr 'Isocelesstyle
    End If
    If cStyle(81) = 1 Then
        pic.Line (pic.Height / 2 - i / 2, 0)-(pic.Height / 2, i), cselColr  'Isocelesstyle
    End If
    If cStyle(82) = 1 Then
        pic.Line (0, pic.Height / 2 - i / 2)-(i, pic.Height / 2), cselColr 'Isocelesstyle
    End If
    If cStyle(83) = 1 Then
        pic.Line (0, pic.Height / 2 + i / 2)-(i, pic.Height / 2), cselColr 'Isocelesstyle
    End If
    If cStyle(84) = 1 Then
        pic.Line (pic.Height - i, pic.Height / 2)-(pic.Height, pic.Height / 2 - i / 2), cselColr 'Isocelesstyle
    End If
    If cStyle(85) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height, pic.Height - i / 2), cselColr 'Isocelesstyle
    End If
    If cStyle(86) = 1 Then
        pic.Line (0, pic.Height / 2)-(pic.Height, i), cselColr 'comet
    End If
    If cStyle(87) = 1 Then
        pic.Line (0, i)-(pic.Height, pic.Height / 2), cselColr 'comet
    End If
    If cStyle(88) = 1 Then
        pic.Line (i, pic.Height)-(pic.Height / 2, 0), cselColr 'comet
    End If
    If cStyle(89) = 1 Then
        pic.Line (i, 0)-(pic.Height / 2, pic.Height), cselColr 'comet
    End If
    If cStyle(90) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 - i), i, cselColr 'circles
    End If
    If cStyle(91) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 - i), i, cselColr 'circles
    End If
    If cStyle(92) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 - i / 2), i, cselColr 'circles
    End If
    If cStyle(93) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 + i / 2), i, cselColr 'circles
    End If
    If cStyle(94) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 + i / 2), i, cselColr 'circles
    End If
    If cStyle(95) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 - i / 2), i, cselColr 'circles
    End If
    If cStyle(96) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 + i / 2, pic.ScaleWidth / 2 + i), i, cselColr 'circles
    End If
    If cStyle(97) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 - i / 2, pic.ScaleWidth / 2 + i), i, cselColr 'circles
    End If
    If cStyle(98) = 1 Then
        pic.Circle (pic.ScaleWidth / (i + 1), pic.ScaleHeight / (i + 1)), i, cselColr 'Corner circles
    End If
    If cStyle(99) = 1 Then
        pic.Circle (pic.ScaleWidth, 0), i, cselColr 'Corner circles
    End If
    If cStyle(100) = 1 Then
        pic.Circle (0, pic.ScaleHeight), i, cselColr 'Corner circles
    End If
    If cStyle(101) = 1 Then
        pic.Circle (pic.ScaleWidth, pic.ScaleWidth), i, cselColr 'Corner circles
    End If
    If cStyle(102) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - (i + 1), pic.ScaleHeight / 2 - (i + 1)), i, cselColr 'circles w/Box style
    End If
    If cStyle(103) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + (i + 1), pic.ScaleHeight / 2 - (i + 1)), i, cselColr 'circles w/Box style
    End If
    If cStyle(104) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 + i), i, cselColr 'circles w/Box style
    End If
    If cStyle(105) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 + i), i, cselColr 'circles w/Box style
    End If
    If cStyle(106) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 - i / 2), i, cselColr 'circles lunar style
    End If
    If cStyle(107) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 + i / 2), i, cselColr 'circles lunar style
    End If
    If cStyle(108) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 - i / 2), i, cselColr 'circles lunar style
    End If
    If cStyle(109) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 + i / 2), i, cselColr 'circles lunar style
    End If
    If cStyle(110) = 1 Then
        pic.Circle (pic.ScaleWidth / 2, pic.ScaleHeight), i, cselColr 'circles center Border style
    End If
    If cStyle(111) = 1 Then
        pic.Circle (pic.ScaleHeight, pic.ScaleWidth / 2), i, cselColr 'circles center Border style
    End If
    If cStyle(112) = 1 Then
        pic.Circle (pic.ScaleWidth / 2, 0), i, cselColr 'circles center Border style
    End If
    If cStyle(113) = 1 Then
        pic.Circle (0, pic.ScaleWidth / 2), i, cselColr 'circles center Border style
    End If
    If cStyle(114) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2), i, cselColr 'circles offset center
    End If
    If cStyle(115) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2), i, cselColr 'circles offset center
    End If
    If cStyle(116) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2 - i / 2), i, cselColr 'circles offset center
    End If
    If cStyle(117) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2 + i / 2), i, cselColr 'circles offset center
    End If
    If cStyle(118) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2), i, cselColr 'circle center
    End If
    If cStyle(119) = 1 Then
        pic.Circle (pic.ScaleHeight, i), i, cselColr 'circles corner to corner
    End If
    If cStyle(120) = 1 Then
        pic.Circle (pic.ScaleHeight - i, 0), i, cselColr 'circles corner to corner
    End If
    If cStyle(121) = 1 Then
        pic.Circle (0, pic.ScaleHeight - i), i, cselColr 'circles corner to corner
    End If
    If cStyle(122) = 1 Then
        pic.Circle (i, pic.ScaleHeight), i, cselColr 'circles corner to corner
    End If
    If cStyle(123) = 1 Then
        pic.Circle (i, 0), i, cselColr 'circles corner to corner
    End If
    If cStyle(124) = 1 Then
        pic.Circle (0, i), i, cselColr 'circles corner to corner
    End If
    If cStyle(125) = 1 Then
        pic.Circle (pic.ScaleHeight, pic.ScaleHeight - i), i, cselColr 'circles corner to corner
    End If
    If cStyle(126) = 1 Then
        pic.Circle (pic.ScaleHeight - i, pic.ScaleHeight), i, cselColr 'circles corner to corner
    End If
    If cStyle(127) = 1 Then
        pic.Circle (i, pic.ScaleHeight / 2), i, cselColr 'Circles Sides
    End If
    If cStyle(128) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, i), i, cselColr 'Circles Sides
    End If
    If cStyle(129) = 1 Then
        pic.Circle (pic.ScaleHeight - i, pic.ScaleHeight / 2), i, cselColr 'Circles Sides
    End If
    If cStyle(130) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleHeight - i), i, cselColr 'Circles Sides
    End If
    If cStyle(131) = 1 Then '*******************new ones
        pic.Line (pic.Width / 10, i + pic.Width / 10)-(0, i), cselColr 'border style 1a
    End If
    If cStyle(132) = 1 Then
        pic.Line (0, i + pic.Width / 10)-(pic.Width / 10, i), cselColr 'border style 1b
    End If
    If cStyle(133) = 1 Then
        pic.Line (i, pic.Width)-(i + pic.Width / 10, pic.Width - pic.Width / 10), cselColr  'border side 2a
    End If
    If cStyle(134) = 1 Then
        pic.Line (i + pic.Width / 10, pic.Width)-(i, pic.Width - pic.Width / 10), cselColr  'border side 2b
    End If
    If cStyle(135) = 1 Then
        pic.Line (pic.Width, i + pic.Width / 10)-(pic.Width - pic.Width / 10, i), cselColr 'border side 3a
    End If
    If cStyle(136) = 1 Then
        pic.Line (pic.Width - pic.Width / 10, i + pic.Width / 10)-(pic.Width, i), cselColr 'border side 3b
    End If
    If cStyle(137) = 1 Then
        pic.Line (i, pic.Width / 10)-(i + pic.Width / 10, 0), cselColr 'border side 4a
    End If
    If cStyle(138) = 1 Then
        pic.Line (i + pic.Width / 10, pic.Width / 10)-(i, 0), cselColr 'border side 4b
    End If
    If cStyle(139) = 1 Then '*********border straight style  new
        pic.Line (i, 0)-(i, pic.Width / 10), cselColr 'border straight 1
    End If
    If cStyle(140) = 1 Then
        pic.Line (0, i)-(pic.Width / 10, i), cselColr 'border straight 2
    End If
    If cStyle(141) = 1 Then
        pic.Line (i, pic.Width)-(i, pic.Width - pic.Width / 10), cselColr 'border straight 4
    End If
    If cStyle(142) = 1 Then
        pic.Line (pic.Width, i)-(pic.Width - pic.Width / 10, i), cselColr 'border straight 4
    End If
    If cStyle(143) = 1 Then '******** Slant centered style  new
        pic.Line (0, i)-(pic.Width / 2, i + pic.Width / 10), cselColr 'Slant centered 1
    End If
    If cStyle(144) = 1 Then
        pic.Line (pic.Width / 2, i + pic.Width / 10)-(pic.Width, i), cselColr 'Slant centered 2
    End If
    If cStyle(145) = 1 Then
        pic.Line (pic.Width / 2, i)-(pic.Width, i + pic.Width / 10), cselColr  'Slant centered 3
    End If
    If cStyle(146) = 1 Then
        pic.Line (0, i + pic.Width / 10)-(pic.Width / 2, i), cselColr 'Slant centered 4
    End If
    If cStyle(147) = 1 Then
        pic.Line (i, 0)-(i + pic.Width / 10, pic.Width / 2), cselColr 'Slant centered 5
    End If
    If cStyle(148) = 1 Then
        pic.Line (i + pic.Width / 10, pic.Width / 2)-(i, pic.Width), cselColr  'Slant centered 6
    End If
    If cStyle(149) = 1 Then
        pic.Line (i, pic.Width / 2)-(i + pic.Width / 10, pic.Width), cselColr  'Slant centered 7
    End If
    If cStyle(150) = 1 Then
        pic.Line (i + pic.Width / 10, 0)-(i, pic.Width / 2), cselColr 'Slant centered 8
    End If
            
'      DoEvents
     pic.Refresh
     Update_Progress ((i * 100) / pic.Width), "Processing..."
    Next i
    
    pic.ScaleMode = pScaleMode
    pic.Picture = pic.Image
Exit Sub

HandleErr:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

'choose where to go
Sub MakeGrid(pic As PictureBox, ByVal numSteps)
If pColMode = 0 Then
Pattern_SColor pic, numSteps, cPatColor
Else
Pattern_CombColor pic, numSteps
End If
End Sub

'draw pattern with a combination of given number of colors
Sub Pattern_CombColor(pic As PictureBox, ByVal numSteps)
Dim pScaleMode, i
 
 On Error GoTo HandleErr
  
  If numSteps = "" Then Exit Sub
   If numSteps <= 0 Then Exit Sub
    pScaleMode = pic.ScaleMode
    pic.ScaleMode = vbTwips
    pic.Picture = LoadPicture()
    pic.FillStyle = 1
    For i = 0 To pic.Width Step pic.Width / numSteps
    If cStyle(0) = 1 Then
        pic.Line (i, pic.Height)-(0, i), QBColor(i Mod TotColrs) '3D effect
    End If
    If cStyle(1) = 1 Then
         pic.Line (i, 0)-(pic.Height, i), QBColor(i Mod TotColrs) '3D effect
    End If
    If cStyle(2) = 1 Then
    pic.Line (0, pic.Height - i)-(i, 0), QBColor(i Mod TotColrs) '3D effect
    End If
    If cStyle(3) = 1 Then
    pic.Line (pic.Height, pic.Height - i)-(i, pic.Height), QBColor(i Mod TotColrs)  '3D effect
    End If
    If cStyle(4) = 1 Then
         pic.Line (i, pic.Height)-(0, 0), QBColor(i Mod TotColrs) 'topleft
    End If
    If cStyle(5) = 1 Then
         pic.Line (0, 0)-(pic.Height, i), QBColor(i Mod TotColrs) 'topleft
    End If
    If cStyle(6) = 1 Then
         pic.Line (pic.Height, pic.Height)-(i, 0), QBColor(i Mod TotColrs) 'bottomleft
    End If
    If cStyle(7) = 1 Then
         pic.Line (pic.Height, pic.Height)-(0, i), QBColor(i Mod TotColrs) 'bottomleft
    End If
    If cStyle(8) = 1 Then
         pic.Line (0, i)-(pic.Height, 0), QBColor(i Mod TotColrs) 'topright
    End If
    If cStyle(9) = 1 Then
         pic.Line (i, pic.Height)-(pic.Height, 0), QBColor(i Mod TotColrs) 'topright
    End If
    If cStyle(10) = 1 Then
         pic.Line (i, 0)-(0, pic.Height), QBColor(i Mod TotColrs) 'bottomright
    End If
    If cStyle(11) = 1 Then
         pic.Line (pic.Height, i)-(0, pic.Height), QBColor(i Mod TotColrs) 'bottomright
    End If
    If cStyle(12) = 1 Then
         pic.Line (0, i)-(pic.Height, i), QBColor(i Mod TotColrs) 'horz
    End If
    If cStyle(13) = 1 Then
        pic.Line (i, pic.Height)-(i, 0), QBColor(i Mod TotColrs) 'vert
    End If
    If cStyle(14) = 1 Then
       pic.Line (0, i)-(i, 0), QBColor(i Mod TotColrs) 'mesh1
    End If
    If cStyle(15) = 1 Then
        pic.Line (i, pic.Height)-(pic.Height, i), QBColor(i Mod TotColrs) 'mesh2
    End If
    If cStyle(16) = 1 Then
        pic.Line (pic.Height - i, 0)-(pic.Height, i), QBColor(i Mod TotColrs) 'mesh3
    End If
    If cStyle(17) = 1 Then
        pic.Line (i, pic.Height)-(0, pic.Height - i), QBColor(i Mod TotColrs)  'mesh4
    End If
    If cStyle(18) = 1 Then
        pic.Line (0, i)-(pic.Height - i, pic.Height - i), QBColor(i Mod TotColrs) '3D effect 1
    End If
    If cStyle(19) = 1 Then
        pic.Line (pic.Height - i, pic.Height - i)-(i, 0), QBColor(i Mod TotColrs) '3D effect 2
    End If
    If cStyle(20) = 1 Then
        pic.Line (pic.Height - i, 0)-(i, pic.Height - i), QBColor(i Mod TotColrs) '3D effect 3
    End If
    If cStyle(21) = 1 Then
        pic.Line (i, pic.Height - i)-(pic.Height, i), QBColor(i Mod TotColrs) '3D effect 4
    End If
    If cStyle(22) = 1 Then
        pic.Line (pic.Height - i, i)-(0, pic.Height - i), QBColor(i Mod TotColrs) '3D effect 5
    End If
    If cStyle(23) = 1 Then
        pic.Line (i, pic.Height)-(pic.Height - i, i), QBColor(i Mod TotColrs) '3D effect 6
    End If
    If cStyle(24) = 1 Then
        pic.Line (pic.Height - i, pic.Height - i)-(i, pic.Height), QBColor(i Mod TotColrs) '3D effect 7
    End If
    If cStyle(25) = 1 Then
        pic.Line (pic.Height, i)-(pic.Height - i, pic.Height - i), QBColor(i Mod TotColrs) '3D effect 8
    End If
    If cStyle(26) = 1 Then
        pic.Line (pic.Height - i, pic.Height - i)-(i, pic.Height - i), QBColor(i Mod TotColrs) 'effect1
    End If
    If cStyle(27) = 1 Then
        pic.Line (pic.Height - i, i)-(pic.Height - i, pic.Height - i), QBColor(i Mod TotColrs) 'effect2
    End If
    If cStyle(28) = 1 Then '
        pic.Line (pic.Height, 0)-(pic.Height - i, pic.Height - i), QBColor(i Mod TotColrs) 'box effect
    End If
    If cStyle(29) = 1 Then '
        pic.Line (pic.Height - i, pic.Height - i)-(0, pic.Height), QBColor(i Mod TotColrs) 'box effect
    End If
    If cStyle(30) = 1 Then
        pic.Line (0, 0)-(i, pic.Height - i), QBColor(i Mod TotColrs) 'box effect
    End If
    If cStyle(31) = 1 Then
        pic.Line (pic.Height - i, i)-(pic.Height, pic.Height), QBColor(i Mod TotColrs) 'box effect
    End If
    If cStyle(32) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height - i, i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(33) = 1 Then
         pic.Line (i, i)-(pic.Height / 2, pic.Height - i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(34) = 1 Then
        pic.Line (0, i)-(pic.Height, pic.Height - i), QBColor(i Mod TotColrs) 'effect5
    End If
    If cStyle(35) = 1 Then
        pic.Line (pic.Height - i, 0)-(i, pic.Height), QBColor(i Mod TotColrs)  'effect6
    End If
    If cStyle(36) = 1 Then
        pic.Line (i, pic.Height)-(i, i), QBColor(i Mod TotColrs)   'line^1
    End If
    If cStyle(37) = 1 Then
        pic.Line (i, pic.Height - i)-(i, pic.Height), QBColor(i Mod TotColrs)  'line^2
    End If
    If cStyle(38) = 1 Then
        pic.Line (i, 0)-(i, pic.Height - i), QBColor(i Mod TotColrs) 'lineV1
    End If
    If cStyle(39) = 1 Then
        pic.Line (i, 0)-(i, i), QBColor(i Mod TotColrs)   'lineV2
    End If
    If cStyle(40) = 1 Then
        pic.Line (pic.Height, i)-(i, i), QBColor(i Mod TotColrs)  'line<1
    End If
    If cStyle(41) = 1 Then
        pic.Line (pic.Height, i)-(pic.Height - i, i), QBColor(i Mod TotColrs)  'line<2
    End If
    If cStyle(42) = 1 Then
        pic.Line (0, i)-(pic.Height - i, i), QBColor(i Mod TotColrs) 'line>1
    End If
    If cStyle(43) = 1 Then
        pic.Line (0, i)-(i, i), QBColor(i Mod TotColrs) 'line>2
    End If
    If cStyle(44) = 1 Then
        pic.Line (pic.Height - i / 2, i)-(i, pic.Height - i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(45) = 1 Then
        pic.Line (i, (pic.Height / 2) - i / 2)-(pic.Height - i, i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(46) = 1 Then
        pic.Line (pic.Height - i, i)-(i, pic.Height - i / 2), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(47) = 1 Then
        pic.Line (i, pic.Height - i)-((pic.Height / 2) - i / 2, i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(48) = 1 Then
        pic.Line (pic.Height - i, pic.Height - i / 2)-(i, i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(49) = 1 Then
        pic.Line (i, i)-(pic.Height - i / 2, pic.Height - i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(50) = 1 Then
        pic.Line ((pic.Height / 2) - i / 2, pic.Height - i)-(i, i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(51) = 1 Then
        pic.Line (i, i)-(pic.Height - i, (pic.Height / 2) - i / 2), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(52) = 1 Then
        pic.Line (pic.Height / 2, i)-(i, pic.Height / 2), QBColor(i Mod TotColrs)  'SlantBox
    End If
    If cStyle(53) = 1 Then
        pic.Line (pic.Height - i, pic.Height / 2)-(pic.Height / 2, i), QBColor(i Mod TotColrs) 'SlantBox
    End If
    If cStyle(54) = 1 Then
        pic.Line (pic.Height - i, i)-(0, pic.Height - i / 2), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(55) = 1 Then
        pic.Line (pic.Height / 2 - i / 2, pic.Height)-(i, pic.Height - i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(56) = 1 Then
        pic.Line (pic.Height - i / 2, 0)-(i, pic.Height - i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(57) = 1 Then
        pic.Line (pic.Height, pic.Height / 2 - i / 2)-(pic.Height - i, i), QBColor(i Mod TotColrs)  '3D
    End If
    If cStyle(58) = 1 Then
        pic.Line (pic.Height, pic.Height - i / 2)-(i, i), QBColor(i Mod TotColrs) '3D effect 8
    End If
    If cStyle(59) = 1 Then
        pic.Line (pic.Height - i / 2, pic.Height)-(i, i), QBColor(i Mod TotColrs) '3D effect 8
    End If
    If cStyle(60) = 1 Then
        pic.Line (0, pic.Height / 2 - i / 2)-(i, i), QBColor(i Mod TotColrs) '3D effect 8
    End If
    If cStyle(61) = 1 Then
        pic.Line (pic.Height / 2 - i / 2, 0)-(i, i), QBColor(i Mod TotColrs) '3D effect 8
    End If
    If cStyle(62) = 1 Then
        pic.Line (pic.Height / 2, i)-(0, pic.Height / 2), QBColor(i Mod TotColrs) 'Box
    End If
    If cStyle(63) = 1 Then
        pic.Line (pic.Height / 2, pic.Height - i)-(pic.Height, pic.Height / 2), QBColor(i Mod TotColrs) 'Box
    End If
    If cStyle(64) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height / 2, 0), QBColor(i Mod TotColrs)  'Box
    End If
    If cStyle(65) = 1 Then
        pic.Line (pic.Height / 2, pic.Height)-(i, pic.Height / 2), QBColor(i Mod TotColrs) 'Box
    End If
    If cStyle(66) = 1 Then
        pic.Line (pic.Height - i, 0)-(i, pic.Height / 2), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(67) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height - i, pic.Height), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(68) = 1 Then
        pic.Line (0, i)-(pic.Height / 2, pic.Height - i), QBColor(i Mod TotColrs)  '3D
    End If
    If cStyle(69) = 1 Then
        pic.Line (pic.Height, i)-(pic.Height / 2, pic.Height - i), QBColor(i Mod TotColrs) '3D
    End If
    If cStyle(70) = 1 Then
        pic.Line (i, pic.Height / 2)-(0, 0), QBColor(i Mod TotColrs)  'trianglestyle
    End If
    If cStyle(71) = 1 Then
        pic.Line (0, pic.Height)-(i, pic.Height / 2), QBColor(i Mod TotColrs)  'trianglestyle
    End If
    If cStyle(72) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height, 0), QBColor(i Mod TotColrs)  'trianglestyle
    End If
    If cStyle(73) = 1 Then
        pic.Line (pic.Height, pic.Height)-(i, pic.Height / 2), QBColor(i Mod TotColrs)  'trianglestyle
    End If
    If cStyle(74) = 1 Then
        pic.Line (0, pic.Height / 2)-(i, pic.Height), QBColor(i Mod TotColrs)  'trianglestyle
    End If
    If cStyle(75) = 1 Then
        pic.Line (i, 0)-(0, pic.Height / 2), QBColor(i Mod TotColrs)  'trianglestyle
    End If
    If cStyle(76) = 1 Then
        pic.Line (pic.Height, pic.Height / 2)-(i, pic.Height), QBColor(i Mod TotColrs)  'trianglestyle
    End If
    If cStyle(77) = 1 Then
        pic.Line (i, 0)-(pic.Height, pic.Height / 2), QBColor(i Mod TotColrs)  'trianglestyle
    End If
    If cStyle(78) = 1 Then
        pic.Line (pic.Height / 2, i)-(pic.Height - i / 2, pic.Height), QBColor(i Mod TotColrs) 'Isocelesstyle
    End If
    If cStyle(79) = 1 Then
        pic.Line (pic.Height / 2, pic.Height - i)-(pic.Height / 2 - i / 2, pic.Height), QBColor(i Mod TotColrs) 'Isocelesstyle
    End If
    If cStyle(80) = 1 Then
        pic.Line (pic.Height - i / 2, 0)-(pic.Height / 2, pic.Height - i), QBColor(i Mod TotColrs) 'Isocelesstyle
    End If
    If cStyle(81) = 1 Then
        pic.Line (pic.Height / 2 - i / 2, 0)-(pic.Height / 2, i), QBColor(i Mod TotColrs)  'Isocelesstyle
    End If
    If cStyle(82) = 1 Then
        pic.Line (0, pic.Height / 2 - i / 2)-(i, pic.Height / 2), QBColor(i Mod TotColrs) 'Isocelesstyle
    End If
    If cStyle(83) = 1 Then
        pic.Line (0, pic.Height / 2 + i / 2)-(i, pic.Height / 2), QBColor(i Mod TotColrs) 'Isocelesstyle
    End If
    If cStyle(84) = 1 Then
        pic.Line (pic.Height - i, pic.Height / 2)-(pic.Height, pic.Height / 2 - i / 2), QBColor(i Mod TotColrs) 'Isocelesstyle
    End If
    If cStyle(85) = 1 Then
        pic.Line (i, pic.Height / 2)-(pic.Height, pic.Height - i / 2), QBColor(i Mod TotColrs) 'Isocelesstyle
    End If
    If cStyle(86) = 1 Then
        pic.Line (0, pic.Height / 2)-(pic.Height, i), QBColor(i Mod TotColrs) 'comet
    End If
    If cStyle(87) = 1 Then
        pic.Line (0, i)-(pic.Height, pic.Height / 2), QBColor(i Mod TotColrs) 'comet
    End If
    If cStyle(88) = 1 Then
        pic.Line (i, pic.Height)-(pic.Height / 2, 0), QBColor(i Mod TotColrs) 'comet
    End If
    If cStyle(89) = 1 Then
        pic.Line (i, 0)-(pic.Height / 2, pic.Height), QBColor(i Mod TotColrs) 'comet
    End If
    If cStyle(90) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 - i), i, QBColor(i Mod TotColrs) 'circles
    End If
    If cStyle(91) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 - i), i, QBColor(i Mod TotColrs) 'circles
    End If
    If cStyle(92) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles
    End If
    If cStyle(93) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles
    End If
    If cStyle(94) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles
    End If
    If cStyle(95) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles
    End If
    If cStyle(96) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 + i / 2, pic.ScaleWidth / 2 + i), i, QBColor(i Mod TotColrs) 'circles
    End If
    If cStyle(97) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 - i / 2, pic.ScaleWidth / 2 + i), i, QBColor(i Mod TotColrs) 'circles
    End If
    If cStyle(98) = 1 Then
        pic.Circle (pic.ScaleWidth / (i + 1), pic.ScaleHeight / (i + 1)), i, QBColor(i Mod TotColrs) 'Corner circles
    End If
    If cStyle(99) = 1 Then
        pic.Circle (pic.ScaleWidth, 0), i, QBColor(i Mod TotColrs) 'Corner circles
    End If
    If cStyle(100) = 1 Then
        pic.Circle (0, pic.ScaleHeight), i, QBColor(i Mod TotColrs) 'Corner circles
    End If
    If cStyle(101) = 1 Then
        pic.Circle (pic.ScaleWidth, pic.ScaleWidth), i, QBColor(i Mod TotColrs) 'Corner circles
    End If
    If cStyle(102) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - (i + 1), pic.ScaleHeight / 2 - (i + 1)), i, QBColor(i Mod TotColrs) 'circles w/Box style
    End If
    If cStyle(103) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + (i + 1), pic.ScaleHeight / 2 - (i + 1)), i, QBColor(i Mod TotColrs) 'circles w/Box style
    End If
    If cStyle(104) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 + i), i, QBColor(i Mod TotColrs) 'circles w/Box style
    End If
    If cStyle(105) = 1 Then
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 + i), i, QBColor(i Mod TotColrs) 'circles w/Box style
    End If
    If cStyle(106) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles lunar style
    End If
    If cStyle(107) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles lunar style
    End If
    If cStyle(108) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles lunar style
    End If
    If cStyle(109) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles lunar style
    End If
    If cStyle(110) = 1 Then
        pic.Circle (pic.ScaleWidth / 2, pic.ScaleHeight), i, QBColor(i Mod TotColrs) 'circles center Border style
    End If
    If cStyle(111) = 1 Then
        pic.Circle (pic.ScaleHeight, pic.ScaleWidth / 2), i, QBColor(i Mod TotColrs) 'circles center Border style
    End If
    If cStyle(112) = 1 Then
        pic.Circle (pic.ScaleWidth / 2, 0), i, QBColor(i Mod TotColrs) 'circles center Border style
    End If
    If cStyle(113) = 1 Then
        pic.Circle (0, pic.ScaleWidth / 2), i, QBColor(i Mod TotColrs) 'circles center Border style
    End If
    If cStyle(114) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2), i, QBColor(i Mod TotColrs) 'circles offset center
    End If
    If cStyle(115) = 1 Then
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2), i, QBColor(i Mod TotColrs) 'circles offset center
    End If
    If cStyle(116) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles offset center
    End If
    If cStyle(117) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles offset center
    End If
    If cStyle(118) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2), i, QBColor(i Mod TotColrs) 'circle center
    End If
    If cStyle(119) = 1 Then
        pic.Circle (pic.ScaleHeight, i), i, QBColor(i Mod TotColrs) 'circles corner to corner
    End If
    If cStyle(120) = 1 Then
        pic.Circle (pic.ScaleHeight - i, 0), i, QBColor(i Mod TotColrs) 'circles corner to corner
    End If
    If cStyle(121) = 1 Then
        pic.Circle (0, pic.ScaleHeight - i), i, QBColor(i Mod TotColrs) 'circles corner to corner
    End If
    If cStyle(122) = 1 Then
        pic.Circle (i, pic.ScaleHeight), i, QBColor(i Mod TotColrs) 'circles corner to corner
    End If
    If cStyle(123) = 1 Then
        pic.Circle (i, 0), i, QBColor(i Mod TotColrs) 'circles corner to corner
    End If
    If cStyle(124) = 1 Then
        pic.Circle (0, i), i, QBColor(i Mod TotColrs) 'circles corner to corner
    End If
    If cStyle(125) = 1 Then
        pic.Circle (pic.ScaleHeight, pic.ScaleHeight - i), i, QBColor(i Mod TotColrs) 'circles corner to corner
    End If
    If cStyle(126) = 1 Then
        pic.Circle (pic.ScaleHeight - i, pic.ScaleHeight), i, QBColor(i Mod TotColrs) 'circles corner to corner
    End If
    If cStyle(127) = 1 Then
        pic.Circle (i, pic.ScaleHeight / 2), i, QBColor(i Mod TotColrs) 'Circles Sides
    End If
    If cStyle(128) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, i), i, QBColor(i Mod TotColrs) 'Circles Sides
    End If
    If cStyle(129) = 1 Then
        pic.Circle (pic.ScaleHeight - i, pic.ScaleHeight / 2), i, QBColor(i Mod TotColrs) 'Circles Sides
    End If
    If cStyle(130) = 1 Then
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleHeight - i), i, QBColor(i Mod TotColrs) 'Circles Sides
    End If
    If cStyle(131) = 1 Then '*******************new ones
        pic.Line (pic.Width / 10, i + pic.Width / 10)-(0, i), QBColor(i Mod TotColrs) 'border style 1a
    End If
    If cStyle(132) = 1 Then
        pic.Line (0, i + pic.Width / 10)-(pic.Width / 10, i), QBColor(i Mod TotColrs) 'border style 1b
    End If
    If cStyle(133) = 1 Then
        pic.Line (i, pic.Width)-(i + pic.Width / 10, pic.Width - pic.Width / 10), QBColor(i Mod TotColrs)  'border side 2a
    End If
    If cStyle(134) = 1 Then
        pic.Line (i + pic.Width / 10, pic.Width)-(i, pic.Width - pic.Width / 10), QBColor(i Mod TotColrs)  'border side 2b
    End If
    If cStyle(135) = 1 Then
        pic.Line (pic.Width, i + pic.Width / 10)-(pic.Width - pic.Width / 10, i), QBColor(i Mod TotColrs) 'border side 3a
    End If
    If cStyle(136) = 1 Then
        pic.Line (pic.Width - pic.Width / 10, i + pic.Width / 10)-(pic.Width, i), QBColor(i Mod TotColrs) 'border side 3b
    End If
    If cStyle(137) = 1 Then
        pic.Line (i + pic.Width / 10, 0)-(i, pic.Width / 10), QBColor(i Mod TotColrs) 'border side 4a
    End If
    If cStyle(138) = 1 Then
        pic.Line (i, 0)-(i + pic.Width / 10, pic.Width / 10), QBColor(i Mod TotColrs) 'border side 4b
    End If
    If cStyle(139) = 1 Then '*********border straight style  new
        pic.Line (i, 0)-(i, pic.Width / 10), QBColor(i Mod TotColrs) 'border straight 1
    End If
    If cStyle(140) = 1 Then
        pic.Line (0, i)-(pic.Width / 10, i), QBColor(i Mod TotColrs) 'border straight 2
    End If
    If cStyle(141) = 1 Then
        pic.Line (i, pic.Width)-(i, pic.Width - pic.Width / 10), QBColor(i Mod TotColrs) 'border straight 4
    End If
    If cStyle(142) = 1 Then
        pic.Line (pic.Width, i)-(pic.Width - pic.Width / 10, i), QBColor(i Mod TotColrs) 'border straight 4
    End If
    If cStyle(143) = 1 Then '******** Slant centered style  new
        pic.Line (0, i)-(pic.Width / 2, i + pic.Width / 10), QBColor(i Mod TotColrs) 'Slant centered 1
    End If
    If cStyle(144) = 1 Then
        pic.Line (pic.Width / 2, i + pic.Width / 10)-(pic.Width, i), QBColor(i Mod TotColrs) 'Slant centered 2
    End If
    If cStyle(145) = 1 Then
        pic.Line (pic.Width / 2, i)-(pic.Width, i + pic.Width / 10), QBColor(i Mod TotColrs)  'Slant centered 3
    End If
    If cStyle(146) = 1 Then
        pic.Line (0, i + pic.Width / 10)-(pic.Width / 2, i), QBColor(i Mod TotColrs) 'Slant centered 4
    End If
    If cStyle(147) = 1 Then
        pic.Line (i, 0)-(i + pic.Width / 10, pic.Width / 2), QBColor(i Mod TotColrs) 'Slant centered 5
    End If
    If cStyle(148) = 1 Then
        pic.Line (i + pic.Width / 10, pic.Width / 2)-(i, pic.Width), QBColor(i Mod TotColrs)  'Slant centered 6
    End If
    If cStyle(149) = 1 Then
        pic.Line (i, pic.Width / 2)-(i + pic.Width / 10, pic.Width), QBColor(i Mod TotColrs)  'Slant centered 7
    End If
    If cStyle(150) = 1 Then
        pic.Line (i + pic.Width / 10, 0)-(i, pic.Width / 2), QBColor(i Mod TotColrs) 'Slant centered 8
    End If
    'DoEvents
     pic.Refresh
     Update_Progress ((i * 100) / pic.Width), "Processing..."
    Next i
    
    pic.ScaleMode = pScaleMode
    pic.Picture = pic.Image
Exit Sub

HandleErr:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

'preview of styles
Sub Preview_Grid(pic As PictureBox, ByVal numSteps, curStyle As Integer)
Dim pScaleMode, i

On Error Resume Next

  If numSteps = "" Then Exit Sub
   If numSteps <= 0 Then Exit Sub
    pScaleMode = pic.ScaleMode
    pic.ScaleMode = vbTwips
    pic.Picture = LoadPicture()
    
    For i = 0 To pic.Width Step pic.Width / numSteps
    
    Select Case curStyle
    Case 0
        pic.Line (0, i)-(i, pic.Height), QBColor(i Mod TotColrs) '3D effect
    Case 1
        pic.Line (i, 0)-(pic.Height, i), QBColor(i Mod TotColrs) '3D effect
    Case 2
        pic.Line (0, pic.Height - i)-(i, 0), QBColor(i Mod TotColrs) '3D effect
    Case 3
        pic.Line (pic.Height, pic.Height - i)-(i, pic.Height), QBColor(i Mod TotColrs)  '3D effect
    Case 4
         pic.Line (i, pic.Height)-(0, 0), QBColor(i Mod TotColrs) 'topleft
    Case 5
         pic.Line (0, 0)-(pic.Height, i), QBColor(i Mod TotColrs) 'topleft
    Case 6
         pic.Line (pic.Height, pic.Height)-(i, 0), QBColor(i Mod TotColrs) 'bottomleft
    Case 7
         pic.Line (pic.Height, pic.Height)-(0, i), QBColor(i Mod TotColrs) 'bottomleft
    Case 8
         pic.Line (0, i)-(pic.Height, 0), QBColor(i Mod TotColrs) 'topright
    Case 9
         pic.Line (i, pic.Height)-(pic.Height, 0), QBColor(i Mod TotColrs) 'topright
    Case 10
         pic.Line (i, 0)-(0, pic.Height), QBColor(i Mod TotColrs) 'bottomright
    Case 11
         pic.Line (pic.Height, i)-(0, pic.Height), QBColor(i Mod TotColrs) 'bottomright
    Case 12
         pic.Line (0, i)-(pic.Height, i), QBColor(i Mod TotColrs) 'horz
    Case 13
        pic.Line (i, pic.Height)-(i, 0), QBColor(i Mod TotColrs) 'vert
    Case 14
       pic.Line (0, i)-(i, 0), QBColor(i Mod TotColrs) 'mesh1
    Case 15
        pic.Line (i, pic.Height)-(pic.Height, i), QBColor(i Mod TotColrs) 'mesh2
    Case 16
        pic.Line (pic.Height - i, 0)-(pic.Height, i), QBColor(i Mod TotColrs) 'mesh3
    Case 17
        pic.Line (i, pic.Height)-(0, pic.Height - i), QBColor(i Mod TotColrs)  'mesh4
    Case 18
        pic.Line (0, i)-(pic.Height - i, pic.Height - i), QBColor(i Mod TotColrs) '3D effect 1
    Case 19
        pic.Line (pic.Height - i, pic.Height - i)-(i, 0), QBColor(i Mod TotColrs) '3D effect 2
    Case 20
        pic.Line (pic.Height - i, 0)-(i, pic.Height - i), QBColor(i Mod TotColrs) '3D effect 3
    Case 21
        pic.Line (i, pic.Height - i)-(pic.Height, i), QBColor(i Mod TotColrs) '3D effect 4
    Case 22
        pic.Line (pic.Height - i, i)-(0, pic.Height - i), QBColor(i Mod TotColrs) '3D effect 5
    Case 23
        pic.Line (i, pic.Height)-(pic.Height - i, i), QBColor(i Mod TotColrs) '3D effect 6
    Case 24
        pic.Line (pic.Height - i, pic.Height - i)-(i, pic.Height), QBColor(i Mod TotColrs) '3D effect 7
    Case 25
        pic.Line (pic.Height, i)-(pic.Height - i, pic.Height - i), QBColor(i Mod TotColrs) '3D effect 8
    Case 26
        pic.Line (pic.Height - i, pic.Height - i)-(i, pic.Height - i), QBColor(i Mod TotColrs) 'effect1
    Case 27
        pic.Line (pic.Height - i, i)-(pic.Height - i, pic.Height - i), QBColor(i Mod TotColrs) 'effect2
    Case 28
        pic.Line (pic.Height, 0)-(pic.Height - i, pic.Height - i), QBColor(i Mod TotColrs) 'box effect
    Case 29
        pic.Line (pic.Height - i, pic.Height - i)-(0, pic.Height), QBColor(i Mod TotColrs) 'box effect
    Case 30
        pic.Line (0, 0)-(i, pic.Height - i), QBColor(i Mod TotColrs) 'box effect
    Case 31
        pic.Line (pic.Height - i, i)-(pic.Height, pic.Height), QBColor(i Mod TotColrs) 'box effect
    Case 32
        pic.Line (i, pic.Height / 2)-(pic.Height - i, i), QBColor(i Mod TotColrs) '3D
    Case 33
         pic.Line (i, i)-(pic.Height / 2, pic.Height - i), QBColor(i Mod TotColrs) '3D
    Case 34
        pic.Line (pic.Height, i)-(0, pic.Height - i), QBColor(i Mod TotColrs) 'effect5
    Case 35
        pic.Line (pic.Height - i, 0)-(i, pic.Height), QBColor(i Mod TotColrs)  'effect6
    Case 36
        pic.Line (i, pic.Height)-(i, i), QBColor(i Mod TotColrs)   'line^1
    Case 37
        pic.Line (i, pic.Height)-(i, pic.Height - i), QBColor(i Mod TotColrs)  'line^2
    Case 38
        pic.Line (i, 0)-(i, pic.Height - i), QBColor(i Mod TotColrs) 'lineV1
    Case 39
        pic.Line (i, 0)-(i, i), QBColor(i Mod TotColrs)   'lineV2
    Case 40
        pic.Line (pic.Height, i)-(i, i), QBColor(i Mod TotColrs)  'line<1
    Case 41
        pic.Line (pic.Height, i)-(pic.Height - i, i), QBColor(i Mod TotColrs)  'line<2
    Case 42
        pic.Line (0, i)-(pic.Height - i, i), QBColor(i Mod TotColrs) 'line>1
    Case 43
        pic.Line (0, i)-(i, i), QBColor(i Mod TotColrs) 'line>2
    Case 44
        pic.Line (pic.Height - i / 2, i)-(i, pic.Height - i), QBColor(i Mod TotColrs) '3D
    Case 45
        pic.Line (i, (pic.Height / 2) - i / 2)-(pic.Height - i, i), QBColor(i Mod TotColrs) '3D
    Case 46
        pic.Line (pic.Height - i, i)-(i, pic.Height - i / 2), QBColor(i Mod TotColrs) '3D
    Case 47
        pic.Line (i, pic.Height - i)-((pic.Height / 2) - i / 2, i), QBColor(i Mod TotColrs) '3D
    Case 48
        pic.Line (pic.Height - i, pic.Height - i / 2)-(i, i), QBColor(i Mod TotColrs) '3D
    Case 49
        pic.Line (i, i)-(pic.Height - i / 2, pic.Height - i), QBColor(i Mod TotColrs) '3D
    Case 50
    pic.Line ((pic.Height / 2) - i / 2, pic.Height - i)-(i, i), QBColor(i Mod TotColrs) '3D
    Case 51
        pic.Line (i, i)-(pic.Height - i, (pic.Height / 2) - i / 2), QBColor(i Mod TotColrs) '3D
    Case 52
        pic.Line (pic.Height / 2, i)-(i, pic.Height / 2), QBColor(i Mod TotColrs) 'SlantBox
    Case 53
        pic.Line (pic.Height / 2, i)-(pic.Height - i, pic.Height / 2), QBColor(i Mod TotColrs) 'SlantBox
    Case 54
        pic.Line (pic.Height - i, i)-(0, pic.Height - i / 2), QBColor(i Mod TotColrs) '3D
    Case 55
        pic.Line (pic.Height / 2 - i / 2, pic.Height)-(i, pic.Height - i), QBColor(i Mod TotColrs) '3D
    Case 56
        pic.Line (pic.Height - i / 2, 0)-(i, pic.Height - i), QBColor(i Mod TotColrs) '3D
    Case 57
        pic.Line (pic.Height, pic.Height / 2 - i / 2)-(pic.Height - i, i), QBColor(i Mod TotColrs)  '3D
    Case 58
        pic.Line (pic.Height, pic.Height - i / 2)-(i, i), QBColor(i Mod TotColrs) '3D effect 8
    Case 59
        pic.Line (pic.Height - i / 2, pic.Height)-(i, i), QBColor(i Mod TotColrs) '3D effect 8
    Case 60
        pic.Line (0, pic.Height / 2 - i / 2)-(i, i), QBColor(i Mod TotColrs) '3D effect 8
    Case 61
        pic.Line (pic.Height / 2 - i / 2, 0)-(i, i), QBColor(i Mod TotColrs) '3D effect 8
    Case 62
        pic.Line (pic.Height / 2, i)-(0, pic.Height / 2), QBColor(i Mod TotColrs) 'Box
    Case 63
        pic.Line (pic.Height / 2, pic.Height - i)-(pic.Height, pic.Height / 2), QBColor(i Mod TotColrs) 'Box
    Case 64
        pic.Line (pic.Height / 2, 0)-(i, pic.Height / 2), QBColor(i Mod TotColrs) 'Box
    Case 65
        pic.Line (pic.Height / 2, pic.Height)-(i, pic.Height / 2), QBColor(i Mod TotColrs) 'Box
    Case 66
        pic.Line (pic.Height - i, 0)-(i, pic.Height / 2), QBColor(i Mod TotColrs) '3D
    Case 67
        pic.Line (i, pic.Height / 2)-(pic.Height - i, pic.Height), QBColor(i Mod TotColrs) '3D
    Case 68
        pic.Line (0, i)-(pic.Height / 2, pic.Height - i), QBColor(i Mod TotColrs)  '3D
    Case 69
        pic.Line (pic.Height, i)-(pic.Height / 2, pic.Height - i), QBColor(i Mod TotColrs) '3D
    Case 70
        pic.Line (0, 0)-(i, pic.Height / 2), QBColor(i Mod TotColrs)  'trianglestyle
    Case 71
        pic.Line (0, pic.Height)-(i, pic.Height / 2), QBColor(i Mod TotColrs)  'trianglestyle
    Case 72
        pic.Line (pic.Height, 0)-(i, pic.Height / 2), QBColor(i Mod TotColrs)  'trianglestyle
    Case 73
        pic.Line (pic.Height, pic.Height)-(i, pic.Height / 2), QBColor(i Mod TotColrs)  'trianglestyle
    Case 74
        pic.Line (0, pic.Height / 2)-(i, pic.Height), QBColor(i Mod TotColrs)  'trianglestyle
    Case 75
        pic.Line (0, pic.Height / 2)-(i, 0), QBColor(i Mod TotColrs)  'trianglestyle
    Case 76
        pic.Line (pic.Height, pic.Height / 2)-(i, pic.Height), QBColor(i Mod TotColrs)  'trianglestyle
    Case 77
        pic.Line (pic.Height, pic.Height / 2)-(i, 0), QBColor(i Mod TotColrs)  'trianglestyle
    Case 78
        pic.Line (pic.Height / 2, i)-(pic.Height - i / 2, pic.Height), QBColor(i Mod TotColrs) 'Isocelesstyle
    Case 79
        pic.Line (pic.Height / 2, pic.Height - i)-(pic.Height / 2 - i / 2, pic.Height), QBColor(i Mod TotColrs) 'Isocelesstyle
    Case 80
        pic.Line (pic.Height / 2, pic.Height - i)-(pic.Height - i / 2, 0), QBColor(i Mod TotColrs) 'Isocelesstyle
    Case 81
        pic.Line (pic.Height / 2, i)-(pic.Height / 2 - i / 2, 0), QBColor(i Mod TotColrs)  'Isocelesstyle
    Case 82
        pic.Line (i, pic.Height / 2)-(0, pic.Height / 2 - i / 2), QBColor(i Mod TotColrs) 'Isocelesstyle
    Case 83
        pic.Line (i, pic.Height / 2)-(0, pic.Height / 2 + i / 2), QBColor(i Mod TotColrs) 'Isocelesstyle
    Case 84
        pic.Line (pic.Height - i, pic.Height / 2)-(pic.Height, pic.Height / 2 - i / 2), QBColor(i Mod TotColrs) 'Isocelesstyle
    Case 85
        pic.Line (i, pic.Height / 2)-(pic.Height, pic.Height - i / 2), QBColor(i Mod TotColrs) 'Isocelesstyle
    Case 86
        pic.Line (0, pic.Height / 2)-(pic.Height, i), QBColor(i Mod TotColrs) 'comet
    Case 87
        pic.Line (0, i)-(pic.Height, pic.Height / 2), QBColor(i Mod TotColrs) 'comet
    Case 88
        pic.Line (i, pic.Height)-(pic.Height / 2, 0), QBColor(i Mod TotColrs) 'comet
    Case 89
        pic.Line (i, 0)-(pic.Height / 2, pic.Height), QBColor(i Mod TotColrs) 'comet
    Case 90
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 - i), i, QBColor(i Mod TotColrs) 'circles
    Case 91
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 - i), i, QBColor(i Mod TotColrs) 'circles
    Case 92
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles
    Case 93
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles
    Case 94
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles
    Case 95
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles
    Case 96
        pic.Circle (pic.ScaleHeight / 2 + i / 2, pic.ScaleWidth / 2 + i), i, QBColor(i Mod TotColrs) 'circles
    Case 97
        pic.Circle (pic.ScaleHeight / 2 - i / 2, pic.ScaleWidth / 2 + i), i, QBColor(i Mod TotColrs) 'circles
    Case 98
        pic.Circle (pic.ScaleWidth / (i + 1), pic.ScaleHeight / (i + 1)), i, QBColor(i Mod TotColrs) 'Corner circles
    Case 99
        pic.Circle (pic.ScaleWidth, 0), i, QBColor(i Mod TotColrs) 'Corner circles
    Case 100
        pic.Circle (0, pic.ScaleHeight), i, QBColor(i Mod TotColrs) 'Corner circles
    Case 101
        pic.Circle (pic.ScaleWidth, pic.ScaleWidth), i, QBColor(i Mod TotColrs) 'Corner circles
    Case 102
        pic.Circle (pic.ScaleWidth / 2 - (i + 1), pic.ScaleHeight / 2 - (i + 1)), i, QBColor(i Mod TotColrs) 'circles w/Box style
    Case 103
        pic.Circle (pic.ScaleWidth / 2 + (i + 1), pic.ScaleHeight / 2 - (i + 1)), i, QBColor(i Mod TotColrs) 'circles w/Box style
    Case 104
        pic.Circle (pic.ScaleHeight / 2 - i, pic.ScaleWidth / 2 + i), i, QBColor(i Mod TotColrs) 'circles w/Box style
    Case 105
        pic.Circle (pic.ScaleHeight / 2 + i, pic.ScaleWidth / 2 + i), i, QBColor(i Mod TotColrs) 'circles w/Box style
    Case 106
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles lunar style
    Case 107
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles lunar style
    Case 108
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles lunar style
    Case 109
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles lunar style
    Case 110
        pic.Circle (pic.ScaleWidth / 2, pic.ScaleHeight), i, QBColor(i Mod TotColrs) 'circles center Border style
    Case 111
        pic.Circle (pic.ScaleHeight, pic.ScaleWidth / 2), i, QBColor(i Mod TotColrs) 'circles center Border style
    Case 112
        pic.Circle (pic.ScaleWidth / 2, 0), i, QBColor(i Mod TotColrs) 'circles center Border style
    Case 113
        pic.Circle (0, pic.ScaleWidth / 2), i, QBColor(i Mod TotColrs) 'circles center Border style
    Case 114
        pic.Circle (pic.ScaleWidth / 2 + i / 2, pic.ScaleHeight / 2), i, QBColor(i Mod TotColrs) 'circles offset center
    Case 115
        pic.Circle (pic.ScaleWidth / 2 - i / 2, pic.ScaleHeight / 2), i, QBColor(i Mod TotColrs) 'circles offset center
    Case 116
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2 - i / 2), i, QBColor(i Mod TotColrs) 'circles offset center
    Case 117
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2 + i / 2), i, QBColor(i Mod TotColrs) 'circles offset center
    Case 118
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleWidth / 2), i, QBColor(i Mod TotColrs) 'circle center
    Case 119
        pic.Circle (pic.ScaleHeight, i), i, QBColor(i Mod TotColrs) 'circles corner to corner
    Case 120
        pic.Circle (pic.ScaleHeight - i, 0), i, QBColor(i Mod TotColrs) 'circles corner to corner
    Case 121
        pic.Circle (0, pic.ScaleHeight - i), i, QBColor(i Mod TotColrs) 'circles corner to corner
    Case 122
        pic.Circle (i, pic.ScaleHeight), i, QBColor(i Mod TotColrs) 'circles corner to corner
    Case 123
        pic.Circle (i, 0), i, QBColor(i Mod TotColrs) 'circles corner to corner
    Case 124
        pic.Circle (0, i), i, QBColor(i Mod TotColrs) 'circles corner to corner
    Case 125
        pic.Circle (pic.ScaleHeight, pic.ScaleHeight - i), i, QBColor(i Mod TotColrs) 'circles corner to corner
    Case 126
        pic.Circle (pic.ScaleHeight - i, pic.ScaleHeight), i, QBColor(i Mod TotColrs) 'circles corner to corner
    Case 127
        pic.Circle (i, pic.ScaleHeight / 2), i, QBColor(i Mod TotColrs) 'Circles Sides
    Case 128
        pic.Circle (pic.ScaleHeight / 2, i), i, QBColor(i Mod TotColrs) 'Circles Sides
    Case 129
        pic.Circle (pic.ScaleHeight - i, pic.ScaleHeight / 2), i, QBColor(i Mod TotColrs) 'Circles Sides
    Case 130
        pic.Circle (pic.ScaleHeight / 2, pic.ScaleHeight - i), i, QBColor(i Mod TotColrs) 'Circles Sides
    Case 131 '*******************new ones
        pic.Line (pic.Width / 10, i + pic.Width / 10)-(0, i), QBColor(i Mod TotColrs) 'border style 1a
    Case 132
        pic.Line (0, i + pic.Width / 10)-(pic.Width / 10, i), QBColor(i Mod TotColrs) 'border style 1b
    Case 133
        pic.Line (i, pic.Width)-(i + pic.Width / 10, pic.Width - pic.Width / 10), QBColor(i Mod TotColrs)  'border side 2a
    Case 134
        pic.Line (i + pic.Width / 10, pic.Width)-(i, pic.Width - pic.Width / 10), QBColor(i Mod TotColrs)  'border side 2b
    Case 135
        pic.Line (pic.Width, i + pic.Width / 10)-(pic.Width - pic.Width / 10, i), QBColor(i Mod TotColrs) 'border side 3a
    Case 136
        pic.Line (pic.Width - pic.Width / 10, i + pic.Width / 10)-(pic.Width, i), QBColor(i Mod TotColrs) 'border side 3b
    Case 137
        pic.Line (i, pic.Width / 10)-(i + pic.Width / 10, 0), QBColor(i Mod TotColrs) 'border side 4a
    Case 138
        pic.Line (i + pic.Width / 10, pic.Width / 10)-(i, 0), QBColor(i Mod TotColrs) 'border side 4b
    Case 139 '*********border straight style  new
        pic.Line (i, 0)-(i, pic.Width / 10), QBColor(i Mod TotColrs) 'border straight 1
    Case 140
        pic.Line (0, i)-(pic.Width / 10, i), QBColor(i Mod TotColrs) 'border straight 2
    Case 141
        pic.Line (i, pic.Width)-(i, pic.Width - pic.Width / 10), QBColor(i Mod TotColrs) 'border straight 4
    Case 142
        pic.Line (pic.Width, i)-(pic.Width - pic.Width / 10, i), QBColor(i Mod TotColrs) 'border straight 4
    Case 143 '******** Slant centered style  new
        pic.Line (0, i)-(pic.Width / 2, i + pic.Width / 10), QBColor(i Mod TotColrs) 'Slant centered 1
    Case 144
        pic.Line (pic.Width / 2, i + pic.Width / 10)-(pic.Width, i), QBColor(i Mod TotColrs) 'Slant centered 2
    Case 145
        pic.Line (pic.Width / 2, i)-(pic.Width, i + pic.Width / 10), QBColor(i Mod TotColrs)  'Slant centered 3
    Case 146
        pic.Line (0, i + pic.Width / 10)-(pic.Width / 2, i), QBColor(i Mod TotColrs) 'Slant centered 4
    Case 147
        pic.Line (i, 0)-(i + pic.Width / 10, pic.Width / 2), QBColor(i Mod TotColrs) 'Slant centered 5
    Case 148
        pic.Line (i + pic.Width / 10, pic.Width / 2)-(i, pic.Width), QBColor(i Mod TotColrs)  'Slant centered 6
    Case 149
        pic.Line (i, pic.Width / 2)-(i + pic.Width / 10, pic.Width), QBColor(i Mod TotColrs)  'Slant centered 7
    Case 150
        pic.Line (i + pic.Width / 10, 0)-(i, pic.Width / 2), QBColor(i Mod TotColrs) 'Slant centered 8
    End Select
      
      
      
      DoEvents
     pic.Refresh
    Next i
    pic.Picture = pic.Image
Exit Sub
End Sub

'draws the selected item
Public Sub DrawItem(pic As PictureBox, curItem As String, ByVal cClr As Long, _
ByVal StartX As Single, ByVal StartY As Single, ByVal EndX As Single, ByVal EndY As Single)
On Error Resume Next

Select Case curItem
    Case "line"
    pic.Line (StartX, StartY)-(EndX, EndY), cClr
    Case "box"
    Call Rectangle(pic.hDC, StartX, StartY, EndX, EndY)
    Case "circle"
    Call Ellipse(pic.hDC, StartX, StartY, EndX, EndY)
    Case "chord"
    Call Chord(pic.hDC, StartX, StartY, EndX, EndY, StartX + 5, StartY + 5, EndX + 5, EndY + 5)
    Case "arc"
    Call Arc(pic.hDC, StartX, StartY, EndX, EndY, StartX + 5, StartY + 5, EndX + 5, EndY + 5)
End Select
Exit Sub
End Sub

'draws a text on the picture box
Public Sub DrawText(DestDC As Long, ByVal tX As Long, ByVal tY As Long, ByVal oText As String)
On Error Resume Next
Call TextOut(DestDC, tX, tY - 6, oText, Len(oText))
Exit Sub
End Sub

'prints the tile or back
Public Sub Print_Tile(picDest As PictureBox, ByVal printMode As Integer, ByVal NumCopies As Integer)
Dim CCopies As Integer

'On Error GoTo HandlePrintERR
frmPrint.lblmsg.Caption = "Please Wait..."

Printer.Scale (-1, -1.5)-(7.5, 12)

For CCopies = 1 To NumCopies
frmPrint.lblmsg.Caption = "Sending " & NumCopies & " page[s] to printer..."
If printMode = 0 Then
  Printer.PaintPicture picDest.Picture, -0.5, -0.5, 7.25, 12, -0.5, -0.5, 7.25, 12, vbSrcCopy 'normal
Else
Printer.PaintPicture picDest.Picture, -0.5, -0.5 'stretched
End If
    If NumCopies > 1 Then
        If CCopies <= (NumCopies - 1) Then
          Printer.NewPage
        End If
    End If
Next
frmPrint.lblmsg.Caption = "Done."
Printer.EndDoc 'start print
Exit Sub

HandlePrintERR:
MsgBox Err.Description, vbCritical
frmMain.stbar.Text = "Error in printing!"
Printer.EndDoc
Exit Sub
End Sub


'This is a simple procedure I made. It offdsets a picture
'from one picturebox to another given the styles:
' 0 for Horizontal offset
' 1 for Vertical offset
' 2 for both
Sub Offset_Image(mainpic As PictureBox, destpic As PictureBox, ByVal OffsetStyle As Integer)
Dim Wid, Hgt As Single
Dim sMode1
Dim sMode2

On Error GoTo Offhandler

sMode1 = mainpic.ScaleMode
sMode2 = destpic.ScaleMode

Wid = mainpic.ScaleWidth
Hgt = mainpic.ScaleHeight

mainpic.ScaleMode = vbPixels
destpic.ScaleMode = vbPixels

destpic.Picture = LoadPicture()

Select Case OffsetStyle
Case 0 'horizontal
destpic.PaintPicture mainpic, 0, Wid / 2, Wid, Hgt / 2, 0, 0, Wid, Hgt / 2, vbSrcCopy   'horz half 1
destpic.PaintPicture mainpic, 0, 0, Wid, Hgt / 2, 0, Wid / 2, Wid, Hgt / 2, vbSrcCopy   'horz half 2
Case 1 'vertical
destpic.PaintPicture mainpic, Wid / 2, 0, Wid / 2, Hgt, 0, 0, Wid / 2, Hgt, vbSrcCopy   'vert half 1
destpic.PaintPicture mainpic, 0, 0, Wid / 2, Hgt, Wid / 2, 0, Wid / 2, Hgt, vbSrcCopy    'vert half 2
Case 2 'both
destpic.PaintPicture mainpic, Wid / 2, Hgt / 2, Wid / 2, Hgt / 2, 0, 0, Wid / 2, Hgt / 2, vbSrcCopy   'half 1
destpic.PaintPicture mainpic, 0, 0, Wid / 2, Hgt / 2, Wid / 2, Hgt / 2, Wid / 2, Hgt / 2, vbSrcCopy   'half 2
destpic.PaintPicture mainpic, Wid / 2, 0, Wid, Hgt, 0, Hgt / 2, Wid, Hgt, vbSrcCopy    'half 3
destpic.PaintPicture mainpic, 0, Hgt / 2, Wid, Hgt, Wid / 2, 0, Wid, Hgt, vbSrcCopy    'half 4
End Select

mainpic.Refresh
destpic.Refresh
destpic.Picture = destpic.Image
mainpic.ScaleMode = sMode1
destpic.ScaleMode = sMode2
Exit Sub

Offhandler:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

'this procedure fills an area of a picture box with a given
'color
Public Sub Fill_Area(Pic_Work As PictureBox, X As Single, Y As Single, ByVal withColr As Long)
Dim fR
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    Pic_Work.FillStyle = vbSolid
    Pic_Work.FillColor = withColr
    fR = ExtFloodFill(Pic_Work.hDC, X, Y, Pic_Work.Point(X, Y), FLOODFILLSURFACE)
    Pic_Work.Refresh
    Pic_Work.Picture = Pic_Work.Image
    Screen.MousePointer = vbDefault
    
    Exit Sub
End Sub


Sub Draw_Preview(picParent As PictureBox, picPrev As PictureBox)
Dim old_ScaleMode, Wid, Hgt
On Error Resume Next
If picParent.Picture = 0 Then Exit Sub
Wid = picPrev.ScaleWidth
Hgt = picPrev.ScaleHeight
old_ScaleMode = picParent.ScaleMode
picParent.ScaleMode = 3
picPrev.Picture = LoadPicture()
picPrev.PaintPicture picParent.Picture, _
        0, 0, Wid, Hgt
picParent.ScaleMode = old_ScaleMode
picPrev.Refresh
picPrev.Picture = picPrev.Image
Exit Sub
End Sub


Sub MakeIt3D(Ctrl As Control, nBevel%, nSpace%, bInset%)
'Makes the control appear on a 3D platform 3D.
''Parameters:
' Ctrl = apply 3D look to control name
' nBevel% = bevel width (pixels)
' nSpace% = surround distance from control (pixels)
' bInset% = True is 3D inset border' False is 3D outset border

PixX% = Screen.TwipsPerPixelX
PixY% = Screen.TwipsPerPixelY
CTop% = Ctrl.Top - PixX%
CLft% = Ctrl.Left - PixY%
CRgt% = Ctrl.Left + Ctrl.Width
CBtm% = Ctrl.Top + Ctrl.Height
If bInset% Then 'recessed border
For i% = nSpace% To (nBevel% + nSpace% - 1)
AddX% = i% * PixX%
AddY% = i% * PixY%
Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + AddX%, CTop% - AddY%), &HFFFFFF
Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - AddX%, CBtm% + AddY%), &HFFFFFF
Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + AddX% + PixX%, CBtm% + AddY%), &H808080
Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + AddX%, CBtm% + AddY%), &H808080
Next
Else 'raised border
For i% = nSpace% To (nBevel% + nSpace% - 5)
AddX% = i% * PixX%
AddY% = i% * PixY%
Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), &HFFFFFF
Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), &HFFFFFF
Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), &H808080
Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), &H808080
Next
End If
End Sub

'
'Special filters Thanks to the  Authors of
'Visual Basic Black Book. It is a great book
'it taught me how to use graphics effects

'To apply a colour Lens effect to an image
Public Sub ColorLens_Image(picSource As PictureBox, picDest As PictureBox, _
ByVal RVal, GVal, BVal As Long)
Dim Wid As Single, Hgt As Single
Dim X, Y As Single
Dim start
Dim r, g, b As Long

Wid = picSource.ScaleWidth
Hgt = picSource.ScaleHeight

picDest.Width = picSource.Width
picDest.Height = picSource.Height

For X = 0 To Wid
    For Y = 0 To Hgt
         r = (picSource.Point(X, Y) And RVal)
         g = picSource.Point(X, Y) And GVal
         b = picSource.Point(X, Y) And BVal
        picDest.PSet (X, Y), RGB(r, g, b)
    Next
    Update_Progress ((X * 100) / Wid), "Processing..."
Next
picDest.Refresh
End Sub

'this procedure mosaics an image with a given pixel size
Public Sub Mosaic_Image(picSource As PictureBox, picDest As PictureBox, ByVal mRange As Variant)
    Dim Wid As Single, Hgt As Single
    Dim X, Y As Single
    Dim bytRed, bytGreen, bytBlue As Byte
    Dim pCenter As Single
    Dim rRangeI, rRangeJ As Integer
    Dim pC, pR As Single
    Dim cLimit, rLimit
    Dim i, j As Single
    
    Wid = picSource.ScaleWidth
    Hgt = picSource.ScaleHeight
    
    picDest.Width = picSource.Width
    picDest.Height = picSource.Height

For X = 0 To Wid Step (mRange + 1)
    For Y = 0 To Hgt Step (mRange + 1)

        'Work out the distance between the square division grid, and the pixel to get data from.
            pCenter = (mRange) \ 2
            
        'Pixel size to copy over
            rRangeI = (mRange)
            rRangeJ = (mRange)
            
            'Check if it's running out of range
            If X + mRange > Wid Then rRangeI = Wid - X
            If Y + mRange > Hgt Then rRangeJ = Hgt - Y
            
            'Work out where to get the data from
            pC = X + pCenter
            pR = Y + pCenter
            
            If pC > Wid Then pC = X
            If pR > Hgt Then pR = Y
            
            'get the colors from point
            bytRed = ((picSource.Point(X, Y) And &HFF) + (picSource.Point(X, Y) And &HFF)) / 2
            bytGreen = (((picSource.Point(X, Y) And &HFF00) / &H100) Mod &H100 + ((picSource.Point(X, Y) And &HFF00) / &H100) Mod &H100) / 2
            bytBlue = (((picSource.Point(X, Y) And &HFF0000) / &H10000) Mod &H100 + ((picSource.Point(X, Y) And &HFF0000) / &H10000) Mod &H100) / 2
            
            If bytRed < 0 Then bytRed = 0
            If bytGreen < 0 Then bytGreen = 0
            If bytBlue < 0 Then bytBlue = 0
            
            
            If X = 0 Then cLimit = -pCenter
            If Y = 0 Then rLimit = -pCenter
            
            'Copy the palette entry number over the region's pixels
            For i = cLimit To (rRangeI)
                For j = rLimit To (rRangeJ)
                    picDest.PSet (X + i, Y + j), RGB(bytRed, bytGreen, bytBlue)
                Next j
            Next i
    Next Y
'    Update_Progress ((X * 100) / Wid),"Processing..."
Next X
picDest.Refresh
End Sub

'provide a progress bar
Sub Update_Progress(ByVal cProgI As Single, ByVal StatusText As String)
frmMain.ProgBar.Line (0, 0)-(cProgI, 10), vbBlue, BF
frmMain.ProgBar.CurrentX = 2
frmMain.ProgBar.CurrentY = 0
frmMain.ProgBar.ForeColor = vbWhite
frmMain.ProgBar.Print StatusText
End Sub


'Using Array and is three times faster
'This is a more flexible procedure. It would probably
'Emboss or Engrave and  applies Motion Blur effect
'requires a source picturebox and the destination
'picture box along with the filter type
'This procedure is quite user friendly
Public Sub Process_Image(picSource As PictureBox, picDest As PictureBox, ByVal sFilter As SpecialFilters)
    Dim Wid As Single, Hgt As Single
    Dim MinX, MinY, maxX, maxY As Single
    Dim OffsetX, OffsetY As Integer
    Dim SkipX1, SkipY1, SkipX2, SkipY2 As Integer
    Dim Flow As Integer
    Dim X, Y As Integer
    Dim bytRed, bytGreen, bytBlue, bytAverage As Integer
    Dim pixels() As Long
    
  
    'set the initial values
    Wid = picSource.ScaleWidth 'maxX
    Hgt = picSource.ScaleHeight 'maxY
    picDest.Width = picSource.Width
    picDest.Height = picSource.Height
    picDest.Picture = LoadPicture()
    
    'get filter
    Select Case sFilter
    Case fEmboss 'emboss
        MinX = Wid
        MinY = Hgt
        maxX = 0
        maxY = 0
        OffsetX = -1
        OffsetY = -1
        Flow = -1
        SkipX1 = 0
        SkipY1 = 0
        SkipX2 = 0
        SkipY2 = 0
    Case fEngrave 'engrave
        MinX = -1
        MinY = -1
        maxX = Wid
        maxY = Hgt
        OffsetX = 1
        OffsetY = 1
        Flow = 1
        SkipX1 = 0
        SkipY1 = 0
        SkipX2 = -1
        SkipY2 = -1
    Case fMotionBlur
        MinX = Wid
        MinY = Hgt
        maxX = -1
        maxY = -1
        OffsetX = 1
        OffsetY = 1
        Flow = -1
        SkipX1 = -2
        SkipY1 = -2
        SkipX2 = 0
        SkipY2 = 0
    End Select
    
    'Redimension array
    ReDim pixels(-1 To Wid, -1 To Hgt) As Long
    
    'Read pixels
    For X = -1 To Wid
        For Y = -1 To Hgt
            pixels(X, Y) = picSource.Point(X, Y)
        Next Y
        Update_Progress ((X * 100) / Wid), "Extracting Pixels..."
    Next X
    frmMain.ProgBar.Cls
     
    'determine colors
    For X = MinX + SkipX1 To maxX + SkipX2 Step Flow
        For Y = MinY + SkipY1 To maxY + SkipY2 Step Flow
            
            If sFilter = fMotionBlur Then
            bytRed = ((pixels(X + OffsetX, Y + OffsetY) And &HFF) + (pixels(X, Y) And &HFF)) / 2
            bytGreen = (((pixels(X + OffsetX, Y + OffsetY) And &HFF00) / &H100) Mod &H100 + ((pixels(X, Y) And &HFF00) / &H100) Mod &H100) / 2
            bytBlue = (((pixels(X + OffsetX, Y + OffsetY) And &HFF0000) / &H10000) Mod &H100 + ((pixels(X, Y) And &HFF0000) / &H10000) Mod &H100) / 2
            Else
            bytRed = ((pixels(X + OffsetX, Y + OffsetY) And &HFF) - (pixels(X, Y) And &HFF)) + 128
            bytGreen = (((pixels(X + OffsetX, Y + OffsetY) And &HFF00) / &H100) Mod &H100 - ((pixels(X, Y) And &HFF00) / &H100) Mod &H100) + 128
            bytBlue = (((pixels(X + OffsetX, Y + OffsetY) And &HFF0000) / &H1000) Mod &H100 - ((pixels(X, Y) And &HFF0000) / &H10000) Mod &H100) + 128
            End If
            
            If bytRed < 0 Then bytRed = 0
            If bytGreen < 0 Then bytGreen = 0
            If bytBlue < 0 Then bytBlue = 0
            
            If bytRed > 255 Then bytRed = 255
            If bytGreen > 255 Then bytGreen = 255
            If bytBlue > 255 Then bytBlue = 255
            
            bytAverage = (bytRed + bytGreen + bytBlue) / 3
            If sFilter = fMotionBlur Then
            pixels(X, Y) = RGB(bytRed, bytGreen, bytBlue)
            Else
            pixels(X, Y) = RGB(bytAverage, bytAverage, bytAverage)
            End If
         Next Y
         If MinX <= 0 Then
         Update_Progress ((X * 100) / Wid), "Processing Colors..."
         Else
         Update_Progress (((MinX - X) * 100) / Wid), "Processing Colors..."
         End If
    Next X
    frmMain.ProgBar.Cls
     
    'replace pixels
    For X = -1 To Wid
        For Y = -1 To Hgt
            picDest.PSet (X, Y), pixels(X, Y)
        Next Y
         Update_Progress ((X * 100) / Wid), "Creating Image..."
    Next X
    picDest.Refresh
End Sub

'provides a disabled effect to an image
Public Sub Disabled_Effect(picSource As PictureBox, picDest As PictureBox)
    Dim X, Y As Integer
    Dim bytRed, bytGreen, bytBlue, bytAverage As Integer
    Dim Wid, Hgt As Single
    Dim pixels() As Long
    
        
    'set the initial values
    Wid = picSource.ScaleWidth 'maxX
    Hgt = picSource.ScaleHeight 'maxY
    picDest.Width = picSource.Width
    picDest.Height = picSource.Height
    picDest.Picture = LoadPicture()
    
    'Redimension array
    ReDim pixels(-1 To Wid, -1 To Hgt) As Long
    
    'Read pixels
    For X = -1 To Wid
        For Y = -1 To Hgt
            pixels(X, Y) = picSource.Point(X, Y)
        Next Y
         Update_Progress ((X * 100) / Wid), "Extracting Pixels..."
    Next X
    frmMain.ProgBar.Cls
        
    For X = -1 To Wid - 1
        For Y = -1 To Hgt - 1
            bytRed = ((pixels(X + 1, Y + 1) And &HFF) - (pixels(X, Y) And &HFF)) + 195
            bytGreen = (((pixels(X + 1, Y + 1) And &HFF00) / &H100) Mod &H100 - ((pixels(X, Y) And &HFF00) / &H100) Mod &H100) + 195
            bytBlue = (((pixels(X + 1, Y + 1) And &HFF0000) / &H10000) Mod &H100 - ((pixels(X, Y) And &HFF0000) / &H10000) Mod &H100) + 195
            
            If bytRed < 0 Then bytRed = 128
            If bytGreen < 0 Then bytGreen = 128
            If bytBlue < 0 Then bytBlue = 128
            bytAverage = (bytRed + bytGreen + bytBlue) / 3
            
           pixels(X, Y) = RGB(bytAverage, bytAverage, bytAverage)
         Next Y
         Update_Progress ((X * 100) / Wid), "Processing Colors..." 'a progress bar
    Next X

    frmMain.ProgBar.Cls
    
    'Replacing...
    For X = -1 To Wid
        For Y = -1 To Hgt
            picDest.PSet (X, Y), pixels(X, Y)
        Next Y
         Update_Progress ((X * 100) / Wid), "Creating Image..."
    Next X
    picDest.Refresh
    picDest.Picture = picDest.Image
End Sub

'applies a cloth effect to an image
Public Sub Cloth_Effect(picSource As PictureBox, picDest As PictureBox, _
Optional stX As Integer = 1, Optional stY As Integer = 1, _
Optional RVal As Integer = 0, Optional GVal As Integer = 0, Optional BVal As Integer = 0, _
Optional XRaise As Integer = 1, Optional YRaise As Integer = 1, _
Optional InColor As Boolean = True)

    Dim Wid As Single, Hgt As Single
    Dim X, Y As Integer
    Dim bytRed, bytGreen, bytBlue, bytAverage As Integer
    Dim pixels() As Long
    
        
    'set the initial values
    Wid = picSource.ScaleWidth 'maxX
    Hgt = picSource.ScaleHeight 'maxY
    picDest.Width = picSource.Width
    picDest.Height = picSource.Height
    picDest.Picture = LoadPicture()
        
    'Redimension array
    ReDim pixels(-1 To Wid, -1 To Hgt) As Long
    'Read pixels
    For X = -1 To Wid
        For Y = -1 To Hgt
            pixels(X, Y) = picSource.Point(X, Y)
        Next Y
         Update_Progress ((X * 100) / Wid), "Extracting Pixels..."
    Next X
    frmMain.ProgBar.Cls
    
    'begin the loop for calculations
    For X = -1 To Wid Step stX
       For Y = -1 To Hgt Step stY
            bytRed = pixels(X, Y) And pixels(X, Y) Mod &HFF + RVal
            bytGreen = pixels(X, Y) And pixels(X, Y) Mod &HFF00 + GVal
            bytBlue = pixels(X, Y) And pixels(X, Y) Mod &HFF0000 + BVal
            'determine the range
            If bytRed < 0 Then bytRed = 0
            If bytGreen < 0 Then bytGreen = 0
            If bytBlue < 0 Then bytBlue = 0
            If bytRed > 255 Then bytRed = 255
            If bytGreen > 255 Then bytGreen = 255
            If bytBlue > 255 Then bytBlue = 255
            
            'restore new pixels
            If InColor Then
                pixels(X, Y) = RGB(bytRed, bytGreen, bytBlue)
            Else
                bytAverage = (bytRed + bytGreen + bytBlue) / 3
                pixels(X, Y) = RGB(bytAverage, bytAverage, bytAverage)
            End If
            
         Next Y
         Update_Progress ((X * 100) / Wid), "Processing Colors..." 'a progress bar
    Next X
    frmMain.ProgBar.Cls
    
    'Replacing...
    For X = -1 To Wid
        For Y = -1 To Hgt
            picDest.PSet (X - Sin(Y ^ XRaise), Y - Sin(X ^ YRaise)), pixels(X, Y)
        Next Y
         Update_Progress ((X * 100) / Wid), "Creating Image..."
    Next X
    picDest.Refresh
    picDest.Picture = picDest.Image
End Sub

'This procedure replaces a specified color of an image
'with a specified color
'call Replace_Color picture1,picture2,matchcolor,withcolor
Public Sub Replace_Color(picSource As PictureBox, picDest As PictureBox, _
ByVal MatchColor As Long, ByVal WithColor As Long)
    Dim Wid As Single, Hgt As Single
    Dim X, Y As Single
    Dim bytRed, bytGreen, bytBlue, bytAverage As Integer
    Dim pixels() As Long
    
  
    'set the initial values
    Wid = picSource.ScaleWidth  'maxX
    Hgt = picSource.ScaleHeight  'maxY
    picDest.Width = picSource.Width
    picDest.Height = picSource.Height
    picDest.Picture = LoadPicture()
    
    'Redimension array
    ReDim pixels(-1 To Wid, -1 To Hgt) As Long
    
    'Read pixels
    For X = -1 To Wid
        For Y = -1 To Hgt
                pixels(X, Y) = picSource.Point(X, Y)
        Next Y
        Update_Progress ((X * 100) / Wid), "Extracting Pixels..."
    Next X
    frmMain.ProgBar.Cls

    For X = -1 To Wid
        For Y = -1 To Hgt
            bytRed = (pixels(X, Y) And &HFF) ' - (Pixels(X, Y) And &HFF))
            bytGreen = ((pixels(X, Y) And &HFF00) / &H100) Mod &H100 ' - ((Pixels(X, Y) And &HFF00) / &H100) Mod &H100)
            bytBlue = ((pixels(X, Y) And &HFF0000) / &H10000) Mod &H100 ' - ((Pixels(X, Y) And &HFF0000) / &H10000) Mod &H100)
            
            If bytRed < 0 Then bytRed = 0
            If bytGreen < 0 Then bytGreen = 0
            If bytBlue < 0 Then bytBlue = 0
            
            If bytRed > 255 Then bytRed = 255
            If bytGreen > 255 Then bytGreen = 255
            If bytBlue > 255 Then bytBlue = 255
            
            ScanColor = RGB(bytRed, bytGreen, bytBlue)
            If ScanColor = MatchColor Then
                pixels(X, Y) = WithColor
               
            Else
                pixels(X, Y) = RGB(bytRed, bytGreen, bytBlue)
            End If
         
         Next Y
         Update_Progress ((X * 100) / Wid), "Processing Colors..." 'a progress bar
    Next X
    frmMain.ProgBar.Cls
    
    'Replacing...
    For X = -1 To Wid
        For Y = -1 To Hgt
            picDest.PSet (X, Y), pixels(X, Y)
        Next Y
         Update_Progress ((X * 100) / Wid), "Creating Image..."
    Next X
    frmMain.ProgBar.Cls
    picDest.Refresh
picDest.Picture = picDest.Image
picSource.Picture = picDest.Image
End Sub

' Draws a wave effect of the image
Sub DrawWaves(picSource As PictureBox, picDest As PictureBox, _
ByVal Amp As Integer, ByVal WaveLen As Integer, _
Optional Horizontal As Integer = 0, Optional Vertical As Integer = 0)
    Dim Wid As Single, Hgt As Single
    Dim X, Y As Single
    Dim bytRed, bytGreen, bytBlue, bytAverage As Integer
    Dim pixels() As Long
    Dim old_smode As Integer
    Dim WaveLength As Single
    Const pi = 3.14159
    
    If Horizontal = 0 And Vertical = 0 Then Exit Sub
    
    WaveLength = WaveLen * pi
 
    'set the initial values
    Wid = picSource.ScaleWidth  'maxX
    Hgt = picSource.ScaleHeight  'maxY
    picDest.Width = picSource.Width
    picDest.Height = picSource.Height
    picDest.Picture = LoadPicture()
    
    
    'Redimension array
    ReDim pixels(-1 To Wid, -1 To Hgt) As Long
    
    'Read pixels
    For X = -1 To Wid
        For Y = -1 To Hgt
                pixels(X, Y) = picSource.Point(X, Y)
        Next Y
        Update_Progress ((X * 100) / Wid), "Extracting Pixels..."
    Next X
    frmMain.ProgBar.Cls
    
    
    old_smode = picDest.ScaleMode
    picDest.ScaleMode = 3   ' Pixel.
    
    picDest.Picture = LoadPicture()   ' Clear the picture box.
    For X = -1 To Wid
        For Y = -1 To Hgt
            If Horizontal = 1 Then
               picDest.PSet (X, Y + Amp * Sin(X / WaveLength)), pixels(X, Y) 'horizontal
            End If
            If Vertical = 1 Then
               picDest.PSet (X + Amp * Sin(Y / WaveLength), Y), pixels(X, Y) 'vertical
            End If
        Next Y
        Update_Progress ((X * 100) / Wid), "Replacing Pixels..."
    Next X
    frmMain.ProgBar.Cls
    picDest.Picture = picDest.Image
    picDest.ScaleMode = old_smode
End Sub

'rotates an image with a clockwise or anticlockwise rotation
Public Sub Rotate_Image(picSource As PictureBox, picDest As PictureBox, _
Optional Clockwise90 As Boolean = True)
    Dim X, Y As Integer
    Dim Wid As Single, Hgt As Single
    
    Wid = picSource.ScaleWidth 'maxX
    Hgt = picSource.ScaleHeight 'maxY
    picDest.Width = picSource.Width
    picDest.Height = picSource.Height
    picDest.Picture = LoadPicture()

    'Read pixels and set them
    For X = -1 To Wid
        For Y = -1 To Hgt
        If Clockwise90 Then
            picDest.PSet ((Hgt - Y - 1), X), picSource.Point(X, Y)
        Else
            picDest.PSet (Y, (Wid - X - 1)), picSource.Point(X, Y)
        End If
        Next Y
         Update_Progress ((X * 100) / Wid), "Creating Image..."
    Next X
        picDest.Refresh
    picDest.Picture = picDest.Image
End Sub


