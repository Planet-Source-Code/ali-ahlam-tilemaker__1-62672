VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImageProcessDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Enum EFilterTypes
    [_Min]
    eBlur
    eBlurMore
    eSoften
    eSoftenMore
    eSharpen
    eSharpenMore
    eUnSharp
    eEmboss
    eMedian
    eMinimum
    eMaximum
    eCount
    eCustom
    [_Max]
End Enum

Public Enum eFilterError
    eeFilterErrorBase = vbObjectError Or 1048 Or &H500
End Enum

Public Event InitProgress(ByVal lMax As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)

Private m_iSize As Long
Private m_iOffset As Long
Private m_iFilt() As Long
Private m_iWeight As Long

Private m_eFilterType As EFilterTypes

Public RSelection As Long
Public GSelection As Long
Public BSelection As Long

Public Property Let FilterType(ByVal eType As EFilterTypes)
    If (eType > EFilterTypes.[_Min] And eType < EFilterTypes.[_Max]) Then
        m_eFilterType = eType
        If (m_eFilterType <> eCustom) Then
            pBuildFilterArray
        End If
    Else
        Err.Raise eeFilterErrorBase + 2, App.EXEName & ".cImageProcess", "Invalid filter types."
    End If
End Property

Public Property Get FilterArraySize() As Long
    FilterArraySize = m_iSize
End Property
Public Property Let FilterArraySize(ByVal LSize As Long)
    If (LSize Mod 2) = 0 Then
        Err.Raise eeFilterErrorBase + 1, App.EXEName & ".cImageProcess", "Size must be an odd number"
    Else
        If (LSize < 0) Or (LSize > 9) Then
            Err.Raise eeFilterErrorBase + 2, App.EXEName & ".cImageProcess", "Invalid size.  Size should be an odd number from 3 to 9"
        Else
            m_iSize = LSize
            m_iOffset = m_iSize \ 2
            ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset) As Long
        End If
    End If
End Property
Public Property Get FilterValue(ByVal iX As Long, ByVal iY As Long) As Long
    FilterValue = m_iFilt(iX, iY)
End Property
Public Property Let FilterValue(ByVal iX As Long, ByVal iY As Long, ByVal lValue As Long)
    m_iFilt(iX, iY) = lValue
End Property
Public Property Get FilterWeight() As Long
    FilterWeight = m_iWeight
End Property
Public Property Let FilterWeight(lWeight As Long)
    m_iWeight = lWeight
End Property

Private Sub pBuildFilterArray()
Dim i As Long, j As Long
Dim iX As Long, iY As Long, iLM As Long

    m_iWeight = 0
    
    Select Case m_eFilterType
    Case eBlur, eBlurMore
        If (m_eFilterType = eBlur) Then
            FilterArraySize = 3
        Else
            FilterArraySize = 5
        End If
        For i = -m_iOffset To m_iOffset
            For j = -m_iOffset To m_iOffset
                m_iFilt(i, j) = 1
                m_iWeight = m_iWeight + m_iFilt(i, j)
            Next j
        Next i
        
    Case eSoften, eSoftenMore
        If (m_eFilterType = eSoften) Then
            FilterArraySize = 3
        Else
            FilterArraySize = 5
        End If
        For i = -m_iOffset To m_iOffset
            For j = -m_iOffset To m_iOffset
                
                iX = Abs(i)
                iY = Abs(j)
                If (iX > iY) Then
                    iLM = iX
                Else
                    iLM = iY
                End If
                
                If (iLM = 0) Then
                    m_iFilt(i, j) = (m_iSize * (m_iSize / 2#))
                Else
                    m_iFilt(i, j) = m_iOffset - iLM + 1
                End If
                Debug.Print m_iFilt(i, j); ",";
                m_iWeight = m_iWeight + m_iFilt(i, j)
            Next j
            Debug.Print
        Next i
        Debug.Print m_iWeight
        
    Case eSharpen, eSharpenMore
        FilterArraySize = 3
        If (m_eFilterType = eSharpen) Then
            m_iFilt(-1, -1) = -1: m_iFilt(-1, 0) = -1: m_iFilt(-1, 1) = -1
            m_iFilt(0, -1) = -1: m_iFilt(0, 0) = 15: m_iFilt(0, 1) = -1
            m_iFilt(1, -1) = -1: m_iFilt(1, 0) = -1: m_iFilt(1, 1) = -1
        Else
            m_iFilt(-1, -1) = 0: m_iFilt(-1, 0) = -1: m_iFilt(-1, 1) = 0
            m_iFilt(0, -1) = -1: m_iFilt(0, 0) = 5: m_iFilt(0, 1) = -1
            m_iFilt(1, -1) = 0: m_iFilt(1, 0) = -1: m_iFilt(1, 1) = 0
        End If
        For i = -m_iOffset To m_iOffset
            For j = -m_iOffset To m_iOffset
                m_iWeight = m_iWeight + m_iFilt(i, j)
            Next j
        Next i
        
    Case eEmboss
        FilterArraySize = 3
        m_iFilt(-1, -1) = -1: m_iFilt(1, 1) = 1
        m_iWeight = 1
        
    End Select
End Sub

Private Function pbRankFilter( _
        ByRef cImage As cDIBSection, _
        ByRef cBuffer As cDIBSection _
    ) As Boolean
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim X As Long, Y As Long
Dim rgbOffset As Long, xOffset As Long
Dim r As Long, g As Long, b As Long
Dim i As Long, j As Long, yMax As Long, xMax As Long
Dim lTIme As Long
Dim rR As Long, rB As Long, rG As Long
Dim iOffset As Long, iWeight As Long
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cImage.Height 'bmp.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cImage.BytesPerScanLine 'bmp.bmWidthBytes
        .pvData = cImage.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cBuffer.Height 'bmp2.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cBuffer.BytesPerScanLine
        .pvData = cBuffer.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    
    ' Do filter on pict into pict2
    lTIme = timeGetTime
    
    iOffset = 1
    iWeight = 9
    rgbOffset = iOffset * 3
    yMax = cImage.Height - 1 - iOffset
    xMax = (cImage.Width - 1) * 3 - rgbOffset
    RaiseEvent InitProgress(xMax)
    
    For X = rgbOffset To xMax Step 3
        For Y = iOffset To yMax
            'Debug.Print X, Y
            'Debug.Print pict(x + i, y + j), pict(x + 1 + i, y + j), pict(x + 2 + i, y + j)
            
            If m_eFilterType = eMinimum Then
                rR = 255: rG = 255: rB = 255
            Else
                rR = 0: rG = 0: rB = 0
            End If
            
            For i = -iOffset To iOffset
                xOffset = i * 3
                For j = -iOffset To iOffset
                    Select Case m_eFilterType
                    Case eMinimum
                        If pict(X + xOffset, Y + j) < rB Then
                            rB = pict(X + xOffset, Y + j)
                        End If
                        If pict(X + 1 + xOffset, Y + j) < rG Then
                            rG = pict(X + 1 + xOffset, Y + j)
                        End If
                        If pict(X + 2 + xOffset, Y + j) < rR Then
                            rR = pict(X + 2 + xOffset, Y + j)
                        End If
                    Case eMedian
                        rB = rB + pict(X + xOffset, Y + j)
                        rG = rG + pict(X + 1 + xOffset, Y + j)
                        rR = rR + pict(X + 2 + xOffset, Y + j)
                    Case eMaximum
                        If pict(X + xOffset, Y + j) > rB Then
                            rB = pict(X + xOffset, Y + j)
                        End If
                        If pict(X + 1 + xOffset, Y + j) > rG Then
                            rG = pict(X + 1 + xOffset, Y + j)
                        End If
                        If pict(X + 2 + xOffset, Y + j) > rR Then
                            rR = pict(X + 2 + xOffset, Y + j)
                        End If
                    End Select
                Next j
            Next i
            If (m_eFilterType = eMedian) Then
                rR = rR \ iWeight: rG = rG \ iWeight: rB = rB \ iWeight
            End If
            
            If (rR < 0) Then rR = 0
            If (rG < 0) Then rG = 0
            If (rB < 0) Then rB = 0
            If (rR > 255) Then rR = 255
            If (rG > 255) Then rG = 255
            If (rB > 255) Then rB = 255
            
            'Debug.Print rR, rG, rB
            
            pict2(X, Y) = rB: pict2(X + 1, Y) = rG: pict2(X + 2, Y) = rR
            
        Next Y
        RaiseEvent Progress(X)
    Next X
         
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
                
    RaiseEvent Complete(timeGetTime - lTIme)
    
    pbRankFilter = True
        
End Function
Public Function ProcessImage( _
        ByRef cImage As cDIBSection, _
        ByRef cBuffer As cDIBSection _
    ) As Boolean
    Select Case m_eFilterType
    Case eMaximum, eMedian, eMinimum
        cBuffer.LoadPictureBlt cImage.hdc
        ProcessImage = pbRankFilter(cImage, cBuffer)
        cImage.LoadPictureBlt cBuffer.hdc
    Case eCount
    Case eBlur, eBlurMore, eCustom, eSharpen, eSharpenMore, eSoften, eSoftenMore
        cBuffer.LoadPictureBlt cImage.hdc
        ProcessImage = pbStandardFilter(cImage, cBuffer)
        cImage.LoadPictureBlt cBuffer.hdc

    Case eUnSharp
        ' Subtract a blurred version of the image from twice the
        ' original bitmap's value:
        FilterType = eBlur
        cBuffer.LoadPictureBlt cImage.hdc
        pbStandardFilter cImage, cBuffer
        AddImages cBuffer, cImage, -1, 0, 0, 0, 2, 0, 0, 0
        FilterType = eUnSharp
        
    Case eEmboss
        ' Perform emboss filter as normal, then add 127 to the R,G,B
        ' values to give a gray background
        cBuffer.LoadPictureBlt cImage.hdc
        ProcessImage = pbStandardFilter(cImage, cBuffer)
        AddImages cBuffer, cImage, 1, 127, 127, 127, 0, 150, 150, 150
    End Select
End Function
Private Function pbStandardFilter( _
        ByRef cImage As cDIBSection, _
        ByRef cBuffer As cDIBSection _
    ) As Boolean
' these are used to address the pixel using matrices
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim X As Long, Y As Long
Dim rgbOffset As Long, xOffset As Long
Dim r As Long, g As Long, b As Long
Dim i As Long, j As Long, yMax As Long, xMax As Long
Dim lTIme As Long
Dim rR As Long, rB As Long, rG As Long
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cImage.Height 'bmp.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cImage.BytesPerScanLine 'bmp.bmWidthBytes
        .pvData = cImage.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cBuffer.Height 'bmp2.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cBuffer.BytesPerScanLine
        .pvData = cBuffer.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    
    ' Do filter on pict into pict2
    lTIme = timeGetTime
    
    rgbOffset = m_iOffset * 3
    yMax = cImage.Height - 1 - m_iOffset
    xMax = (cImage.Width - 1) * 3 - rgbOffset
    RaiseEvent InitProgress(xMax)
    
    For X = rgbOffset To xMax Step 3
        For Y = m_iOffset To yMax
            'Debug.Print X, Y
            'Debug.Print pict(X + i, Y + j), pict(X + 1 + i, Y + j), pict(X + 2 + i, Y + j)
            r = 0: g = 0: b = 0
            For i = -m_iOffset To m_iOffset
                xOffset = i * 3
                For j = -m_iOffset To m_iOffset
                    r = r + m_iFilt(i, j) * pict(X + xOffset, Y + j)
                    g = g + m_iFilt(i, j) * pict(X + 1 + xOffset, Y + j)
                    b = b + m_iFilt(i, j) * pict(X + 2 + xOffset, Y + j)
                Next j
            Next i
            rR = r \ m_iWeight: rG = g \ m_iWeight: rB = b \ m_iWeight
                If (rR < 0) Then rR = 0
                If (rG < 0) Then rG = 0
                If (rB < 0) Then rB = 0
                If (rR > 255) Then rR = 255
                If (rG > 255) Then rG = 255
                If (rB > 255) Then rB = 255
            'Debug.Print rR, rG, rB, vbCrLf
            pict2(X, Y) = rR: pict2(X + 1, Y) = rG: pict2(X + 2, Y) = rB
        Next Y
        RaiseEvent Progress(X)
    Next X
         
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
        
    RaiseEvent Complete(timeGetTime - lTIme)
    
    pbStandardFilter = True
    
End Function
Public Function AddLightest( _
      ByRef cFrom As cDIBSection, _
      ByRef cTo As cDIBSection _
   )
' these are used to address the pixel using matrices
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim X As Long, Y As Long
Dim i As Long, j As Long, yMax As Long, lTIme As Long
Dim lGray1 As Long, lGray2 As Long

    lTIme = timeGetTime()
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Pict now stores the To buffer
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cFrom.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cFrom.BytesPerScanLine
        .pvData = cFrom.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    ' Pict2 now stores the From buffer

    yMax = cTo.Height - 1
    
    For X = 0 To (cTo.Width - 1) * 3 Step 3
        For Y = 0 To yMax
            lGray1 = (222& * pict(X + 1, Y) + 707& * pict(X + 1, Y) + 71& * pict(X, Y))
            lGray2 = (222& * pict2(X + 1, Y) + 707& * pict2(X + 1, Y) + 71& * pict2(X, Y))
            If (lGray2 < lGray1) Then
               pict(X, Y) = pict2(X, Y)
               pict(X + 1, Y) = pict2(X + 1, Y)
               pict(X + 2, Y) = pict2(X + 2, Y)
            End If
        Next Y
        'prgMain.Value = x
    Next X
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
   
End Function

Public Function AddDarkest( _
      ByRef cFrom As cDIBSection, _
      ByRef cTo As cDIBSection _
   )
' these are used to address the pixel using matrices
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim X As Long, Y As Long
Dim i As Long, j As Long, yMax As Long, lTIme As Long
Dim lGray1 As Long, lGray2 As Long

    lTIme = timeGetTime()
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Pict now stores the To buffer
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cFrom.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cFrom.BytesPerScanLine
        .pvData = cFrom.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    ' Pict2 now stores the From buffer

    yMax = cTo.Height - 1
    
    For X = 0 To (cTo.Width - 1) * 3 Step 3
        For Y = 0 To yMax
            lGray1 = (222& * pict(X + 1, Y) + 707& * pict(X + 1, Y) + 71& * pict(X, Y))
            lGray2 = (222& * pict2(X + 1, Y) + 707& * pict2(X + 1, Y) + 71& * pict2(X, Y))
            If (lGray1 < lGray2) Then
               pict(X, Y) = pict2(X, Y)
               pict(X + 1, Y) = pict2(X + 1, Y)
               pict(X + 2, Y) = pict2(X + 2, Y)
            End If
        Next Y
        'prgMain.Value = x
    Next X
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
   
End Function
Public Function AddImages( _
        ByRef cFrom As cDIBSection, _
        ByRef cTo As cDIBSection, _
        ByVal lFromMultiplier As Long, _
        ByVal lFromOffsetR As Long, ByVal lFromOffsetG As Long, ByVal lFromOffsetB As Long, _
        ByVal lToMultiplier As Long, _
        ByVal lToOffsetR As Long, ByVal lToOffsetG As Long, ByVal lToOffsetB As Long _
    ) As Boolean
' these are used to address the pixel using matrices
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim X As Long, Y As Long
Dim i As Long, j As Long, yMax As Long, lTIme As Long
Dim rR As Long, rG As Long, rB As Long

    lTIme = timeGetTime()
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Pict now stores the To buffer
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cFrom.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cFrom.BytesPerScanLine
        .pvData = cFrom.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    ' Pict2 now stores the From buffer

    yMax = cTo.Height - 1
    
    For X = 0 To (cTo.Width - 1) * 3 Step 3
        For Y = 0 To yMax
            rR = (pict(X, Y) + lToOffsetR) * lToMultiplier + (pict2(X, Y) + lFromOffsetR) * lFromMultiplier
            rG = (pict(X + 1, Y) + lToOffsetG) * lToMultiplier + (pict2(X + 1, Y) + lFromOffsetG) * lFromMultiplier
            rB = (pict(X + 2, Y) + lToOffsetB) * lToMultiplier + (pict2(X + 2, Y) + lFromOffsetG) * lFromMultiplier
                If (rR < 0) Then rR = 0
                If (rG < 0) Then rG = 0
                If (rB < 0) Then rB = 0
                If (rR > 255) Then rR = 255
                If (rG > 255) Then rG = 255
                If (rB > 255) Then rB = 255
            
            pict(X, Y) = rR
            pict(X + 1, Y) = rG
            pict(X + 2, Y) = rB
        Next Y
        'prgMain.Value = x
    Next X
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
        
End Function

Public Sub BlackAndWhite( _
        ByRef cFrom As cDIBSection, _
        ByRef cTo As cDIBSection _
    )
' Converts to Black and WHite using Floyd-Steinberg error diffusion
' process.
' see http://www.dcs.ed.ac.uk/~mxr/gfx/faqs/colourspace.faq for details.
'
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim X As Long, Y As Long
Dim i As Long, iCoeff As Long
Dim lTIme As Long
Dim xMax As Long, yMax As Long
Dim lError As Long
Dim lNew As Long
Dim iC As Long, iC2 As Long

    lTIme = timeGetTime()
       
    GrayScale cFrom
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Pict now stores the To buffer
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cFrom.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cFrom.BytesPerScanLine
        .pvData = cFrom.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    ' Pict2 now stores the From buffer


    yMax = cTo.Height - 1
    xMax = (cTo.Width - 1) * 3
    
    RaiseEvent InitProgress(xMax)
    For X = 0 To xMax Step 3
        For Y = 0 To yMax
            ' Apply a simple threshold:
            If (pict2(X, Y) > 128) Then
                iC = iC + 1
                pict(X, Y) = 255
                pict(X + 1, Y) = 255
                pict(X + 2, Y) = 255
                lError = (255 - pict2(X, Y)) - 128
            Else
                iC2 = iC2 + 1
                pict(X, Y) = 0
                pict(X + 1, Y) = 0
                pict(X + 2, Y) = 0
                ' Black tolerance:
                If (pict2(X, Y) > 16) Then
                    lError = pict2(X, Y)
                Else
                    lError = 0
                End If
            End If
            
            ' Diffuse the error:
            If (X < xMax - 3) Then
                lNew = pict2(X + 3, Y) + (lError * 7) \ 16
                If (lNew > 255) Then lNew = 255
                If (lNew < 0) Then lNew = 0
                pict2(X + 3, Y) = lNew
                pict2(X + 4, Y) = lNew
                pict2(X + 5, Y) = lNew
            End If
            If (Y < yMax) Then
                For i = -3 To 3 Step 3
                    If (X + i) > 0 And (X + i) < xMax Then
                        Select Case i
                        Case -3
                            iCoeff = 3
                        Case 0
                            iCoeff = 5
                        Case 3
                            iCoeff = 1
                        End Select
                        lNew = pict2(X + i, Y + 1) + (lError * iCoeff) \ 16
                        If (lNew > 255) Then lNew = 255
                        If (lNew < 0) Then lNew = 0
                        pict2(X + i, Y + 1) = lNew
                        pict2(X + i + 1, Y + 1) = lNew
                        pict2(X + i + 2, Y + 1) = lNew
                    End If
                Next i
            End If
        Next Y
        RaiseEvent Progress(X)
    Next X
    
    Debug.Print iC, iC2
    cFrom.LoadPictureBlt cTo.hdc
    RaiseEvent Complete(timeGetTime - lTIme)
    
    
End Sub

Public Sub ApplyPalette( _
      ByRef cFrom As cDIBSection, _
      ByRef cTo As cDIBSection, _
      ByRef cPal As cPalette _
   )
'
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim X As Long, Y As Long
Dim i As Long, iCoeff As Long
Dim lTIme As Long
Dim xMax As Long, yMax As Long
Dim lErrorRed As Long, lErrorBlue As Long, lErrorGreen As Long
Dim lNewRed As Long, lNewBlue As Long, lNewGreen As Long
Dim lIndex As Long
Dim iC As Long, iC2 As Long

    lTIme = timeGetTime()
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Pict now stores the To buffer
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cFrom.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cFrom.BytesPerScanLine
        .pvData = cFrom.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    ' Pict2 now stores the From buffer
   

    yMax = cTo.Height - 1
    xMax = (cTo.Width - 1) * 3
    
    RaiseEvent InitProgress(xMax)
    For X = 0 To xMax Step 3
        For Y = 0 To yMax
            ' Get nearest colour:
            lIndex = cPal.ClosestIndex(pict2(X + 2, Y), pict2(X + 1, Y), pict2(X, Y))
                        
            pict(X + 2, Y) = cPal.Red(lIndex)
            pict(X + 1, Y) = cPal.Green(lIndex)
            pict(X, Y) = cPal.Blue(lIndex)
                        
            lErrorRed = -1 * (CLng(pict(X + 2, Y)) - pict2(X + 2, Y))
            lErrorGreen = -1 * (CLng(pict(X + 1, Y)) - pict2(X + 1, Y))
            lErrorBlue = -1 * (CLng(pict(X, Y)) - pict2(X, Y))
            
            ' Diffuse the error:
            'Debug.Print lErrorRed, lErrorGreen, lErrorBlue
            If Abs(lErrorRed) + Abs(lErrorGreen) + Abs(lErrorBlue) > 3 Then
            If (X < xMax - 3) Then
                lNewBlue = pict2(X + 3, Y) + (lErrorBlue * 7) \ 16
                lNewGreen = pict2(X + 4, Y) + (lErrorGreen * 7) \ 16
                lNewRed = pict2(X + 5, Y) + (lErrorRed * 7) \ 16
                Range lNewBlue, 0, 255
                Range lNewGreen, 0, 255
                Range lNewRed, 0, 255
                pict2(X + 3, Y) = lNewBlue
                pict2(X + 4, Y) = lNewGreen
                pict2(X + 5, Y) = lNewRed
            End If
            If (Y < yMax) Then
                For i = -3 To 3 Step 3
                    If (X + i) > 0 And (X + i) < xMax Then
                        Select Case i
                        Case -3
                            iCoeff = 0
                        Case 0
                            iCoeff = 4
                        Case 3
                            iCoeff = 0
                        End Select
                        lNewBlue = pict2(X + i, Y + 1) + (lErrorBlue * iCoeff) \ 16
                        lNewGreen = pict2(X + i + 1, Y + 1) + (lErrorGreen * iCoeff) \ 16
                        lNewRed = pict2(X + i + 2, Y + 1) + (lErrorRed * iCoeff) \ 16
                        Range lNewBlue, 0, 255
                        Range lNewGreen, 0, 255
                        Range lNewRed, 0, 255
                        pict2(X + i, Y + 1) = lNewBlue
                        pict2(X + i + 1, Y + 1) = lNewGreen
                        pict2(X + i + 2, Y + 1) = lNewRed
                    End If
                Next i
            End If
            End If
        Next Y
        RaiseEvent Progress(X)
    Next X
    
    Debug.Print iC, iC2
    cFrom.LoadPictureBlt cTo.hdc
    RaiseEvent Complete(timeGetTime - lTIme)
    
    
End Sub
   
Private Sub Range( _
      ByRef lIn As Long, _
      ByVal lMin As Long, _
      ByVal lMax As Long _
   )
   If (lIn < lMin) Then
      lIn = lMin
   ElseIf (lIn > lMax) Then
      lIn = lMax
   End If
End Sub

Public Sub GrayScale( _
        ByRef cTo As cDIBSection _
    )
' Gray scale using standard intensity components.
' see http://www.dcs.ed.ac.uk/~mxr/gfx/faqs/colourspace.faq for details.
'
Dim bDib() As Byte
Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long
Dim bContinue As Boolean
Dim lB As Long, lG As Long, lR As Long
Dim lGray As Long
Dim lTIme As Long
Dim tSA As SAFEARRAY2D

    lTIme = timeGetTime()
        
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
    yMax = cTo.Height - 1
    xMax = cTo.Width - 1
    
    RaiseEvent InitProgress(xMax)
    For X = 0 To (xMax * 3) Step 3
        For Y = 0 To yMax
            lB = bDib(X, Y)
            lG = bDib(X + 1, Y)
            lR = bDib(X + 2, Y)
                
            'But now all people *should* use the most accurate, it means ITU standard:
            lGray = (222 * lR + 707 * lG + 71 * lB) / 1000
            
            bDib(X, Y) = lGray
            bDib(X + 1, Y) = lGray
            bDib(X + 2, Y) = lGray
        Next Y
        RaiseEvent Progress(X)
    Next X
    
    RaiseEvent Complete(timeGetTime - lTIme)
    
End Sub

Public Sub AddNoise( _
        ByRef cTo As cDIBSection, _
        ByVal lPercent As Long, _
        Optional ByVal bRandom = False _
    )
Dim bDib() As Byte
Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long
Dim bContinue As Boolean
Dim lB As Long, lG As Long, lR As Long
Dim lA As Long, lA2 As Long
Dim lTIme As Long
Dim tSA As SAFEARRAY2D

    lTIme = timeGetTime()
    
    lA = 128 * lPercent \ 100
    lA2 = lA \ 2
       
    
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
    yMax = cTo.Height - 1
    xMax = cTo.Width - 1
    
    RaiseEvent InitProgress(xMax)
    For X = 0 To (xMax * 3) Step 3
        For Y = 0 To yMax
            If (bRandom) Then
                bContinue = False
                If (Rnd * 100 > lPercent) Then
                    bContinue = True
                End If
            End If
            If Not (bRandom) Or bContinue Then 'random colors
                lB = bDib(X, Y)
                lG = bDib(X + 1, Y)
                lR = bDib(X + 2, Y)
                
                lB = lB - lA2 + (Rnd * lA)
                lG = lG - lA2 + (Rnd * lA)
                lR = lR - lA2 + (Rnd * lA)
                
                If (lB < 0) Then lB = 0
                If (lG < 0) Then lG = 0
                If (lR < 0) Then lR = 0
                If (lR > 255) Then lR = 255
                If (lG > 255) Then lG = 255
                If (lB > 255) Then lB = 255
                
                bDib(X, Y) = lB
                bDib(X + 1, Y) = lG
                bDib(X + 2, Y) = lR
                Else '*************with selected colors
                lB = bDib(X, Y)
                lG = bDib(X + 1, Y)
                lR = bDib(X + 2, Y)
                
                lB = lB - lA2 + (RSelection * lA)
                lG = lG - lA2 + (GSelection * lA)
                lR = lR - lA2 + (BSelection * lA)
                
                If (lB < 0) Then lB = 0
                If (lG < 0) Then lG = 0
                If (lR < 0) Then lR = 0
                If (lR > 255) Then lR = 255
                If (lG > 255) Then lG = 255
                If (lB > 255) Then lB = 255
                
                bDib(X, Y) = lB
                bDib(X + 1, Y) = lG
                bDib(X + 2, Y) = lR

            End If
        Next Y
        RaiseEvent Progress(X)
    Next X
    RaiseEvent Complete(timeGetTime - lTIme)
        
End Sub

Public Sub Fade( _
      ByRef cTo As cDIBSection, _
      ByVal lAmount As Long _
   )
Dim bDib() As Byte
Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long
Dim bContinue As Boolean
Dim lB As Long, lG As Long, lR As Long
Dim lA As Long, lA2 As Long
Dim lTIme As Long
Dim tSA As SAFEARRAY2D
    
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
    yMax = cTo.Height - 1
    xMax = cTo.Width - 1
    
    RaiseEvent InitProgress(xMax)
    For X = 0 To (xMax * 3) Step 3
        For Y = 0 To yMax
            lB = lAmount * bDib(X, Y) \ 255
            lG = lAmount * bDib(X + 1, Y) \ 255
            lR = lAmount * bDib(X + 2, Y) \ 255
            bDib(X, Y) = lB
            bDib(X + 1, Y) = lG
            bDib(X + 2, Y) = lR
        Next Y
        RaiseEvent Progress(X)
    Next X
    RaiseEvent Complete(timeGetTime - lTIme)
    
End Sub

Public Sub Lighten( _
      ByRef cTo As cDIBSection, _
      ByVal lAmount As Long _
   )
Dim bDib() As Byte
Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long
Dim bContinue As Boolean
Dim lB As Long, lG As Long, lR As Long
Dim h As Single, s As Single, l As Single
Dim lTIme As Long
Dim tSA As SAFEARRAY2D
    
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
    yMax = cTo.Height - 1
    xMax = cTo.Width - 1
    
    RaiseEvent InitProgress(xMax)
    For X = 0 To (xMax * 3) Step 3
        For Y = 0 To yMax
            RGBToHLS bDib(X + 2, Y), bDib(X + 1, Y), bDib(X, Y), h, s, l
            l = l * (1 + (lAmount / 100))
            If (l > 1) Then l = 1
            HLSToRGB h, s, l, lR, lG, lB
            bDib(X, Y) = lB
            bDib(X + 1, Y) = lG
            bDib(X + 2, Y) = lR
        Next Y
        RaiseEvent Progress(X)
    Next X
    RaiseEvent Complete(timeGetTime - lTIme)
    
End Sub
Public Sub Colourise( _
      ByRef cTo As cDIBSection, _
      ByVal fHue As Single, _
      ByVal fSaturation As Single _
   )
' Saturation only applies to grey scale images.  Otherwise saturation
' is taken from the colour.
Dim bDib() As Byte
Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long
Dim bContinue As Boolean
Dim lB As Long, lG As Long, lR As Long
Dim h As Single, s As Single, l As Single
Dim lTIme As Long
Dim tSA As SAFEARRAY2D
    
    ' fHue runs from -1 to 5...
    
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
    yMax = cTo.Height - 1
    xMax = cTo.Width - 1
    
    RaiseEvent InitProgress(xMax)
    For X = 0 To (xMax * 3) Step 3
        For Y = 0 To yMax
            RGBToHLS bDib(X + 2, Y), bDib(X + 1, Y), bDib(X, Y), h, s, l
            If (h = 0) Then
               ' Set saturation (should allow user to choose...)
               s = 0.5
               h = fHue
            Else
               h = fHue
            End If
            HLSToRGB h, s, l, lR, lG, lB
            bDib(X, Y) = lB
            bDib(X + 1, Y) = lG
            bDib(X + 2, Y) = lR
        Next Y
        RaiseEvent Progress(X)
    Next X
    RaiseEvent Complete(timeGetTime - lTIme)
    

End Sub

