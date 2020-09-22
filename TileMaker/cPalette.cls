VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RGB
   Red As Byte
   Green As Byte
   Blue As Byte
End Type
Private m_tPal() As RGB
Private m_iPalette As Long

Public Property Get ClosestIndex( _
      ByVal Red As Long, _
      ByVal Green As Long, _
      ByVal Blue As Long _
   ) As Long
Dim i As Long
Dim lER As Single, lEB As Single, lEG As Single
Dim lMinER As Single, lMinEB As Single, lMinEG As Single
Dim lMinIndex As Long
Dim hO As Single, sO As Single, lO As Single
Dim hN As Single, sN As Single, lN As Single
Dim hErr As Single, sErr As Single, lErr As Single
   
   lMinER = 255: lMinEB = 255: lMinEG = 255
   'hErr = 10: sErr = 10: lErr = 10
   'RGBToHLS Red, Green, Blue, hO, sO, lO
   'hO = hO + 1 ' hue runs -1 to 5
  '
   'For i = 1 To m_iPalette
   '   With m_tPal(i)
   '      If (Red = .Red) And (Blue = .Blue) And (Green = .Green) Then
   '         ClosestIndex = i
   '         Exit Property
   '      Else
   '         RGBToHLS .Red, .Green, .Blue, hN, sN, lN
   '         hN = hN + 1
   '         If (Abs(hN - hO) < hErr) And (Abs(sN - sO) < sErr) And (Abs(lN - lO) < lErr) Then
   '            hErr = Abs(hN - hO)
   '            sErr = Abs(sN - sO)
   '            lErr = Abs(lN - lO)
   '            lMinIndex = i
   '         End If
   '      End If
   '   End With
   'Next i
   
   For i = 1 To m_iPalette
      With m_tPal(i)
         If (Red = .Red) And (Blue = .Blue) And (Green = .Green) Then
            ClosestIndex = i
            Exit Property
         Else
            lER = Abs(Red - .Red)
            lEB = Abs(Blue - .Blue)
            lEG = Abs(Green - .Green)
            If (lER + lEB + lEG < lMinER + lMinEB + lMinEG) Then
               lMinER = lER
               lMinEB = lEB
               lMinEG = lEG
               lMinIndex = i
               'Debug.Print i, lER, lEB, lEG
            End If
         End If
      End With
   Next i
   
   
   ClosestIndex = lMinIndex
   
   
End Property

Public Property Get Count() As Long
   Count = m_iPalette
End Property
Public Property Get Red(ByVal iIndex) As Byte
   Red = m_tPal(iIndex).Red
End Property
Public Property Get Green(ByVal iIndex) As Byte
   Green = m_tPal(iIndex).Green
End Property
Public Property Get Blue(ByVal iIndex) As Byte
   Blue = m_tPal(iIndex).Blue
End Property

Public Function LoadFromFile(ByVal sFileName As String) As Boolean
Dim iFile As Long
Dim sBuf As String
Dim iPos As Long
Dim iNextPos As Long
Dim sLines() As String, iLineCount As Long, iLine As Long
Dim sParts() As String, iPartCount As Long

   m_iPalette = 0
   Erase m_tPal

   On Error GoTo LoadFromFileError
   ' Load it:
   iFile = FreeFile
   Open sFileName For Binary Access Read As #iFile
   sBuf = String$(LOF(iFile), 32)
   Get #iFile, , sBuf
   Close #iFile
   iFile = 0
   
   SplitDelimitedString sBuf, vbCrLf, sLines(), iLineCount
   For iLine = 1 To iLineCount
      SplitDelimitedString sLines(iLine), " ", sParts(), iPartCount
      If (iPartCount = 3) Then
         m_iPalette = m_iPalette + 1
         ReDim Preserve m_tPal(1 To m_iPalette) As RGB
         With m_tPal(m_iPalette)
            .Red = CByte(sParts(1))
            .Green = CByte(sParts(2))
            .Blue = CByte(sParts(3))
         End With
      End If
   Next iLine
   
   LoadFromFile = True
   
   Exit Function

LoadFromFileError:
   Err.Raise Err.Number, App.EXEName & ".cPalette", Err.Description
   If (iFile <> 0) Then
      Close #iFile
      m_iPalette = 0
      Erase m_tPal
   End If
   Exit Function

End Function

Private Sub SplitDelimitedString( _
        ByVal sString As String, _
        ByVal sDelim As String, _
        ByRef sValues() As String, _
        ByRef iCount As Long _
    )
' ==================================================================
' Splits sString into an array of parts which are
' delimited in the string by sDelim.  The array is
' indexed 1-iCount where iCount is the number of
' items.  If no items found iCount=1 and the array has
' one element, the original string.
'   sString : String to split
'   sDelim  : Delimiter
'   sValues : Return array of values
'   iCount  : Number of items returned in sValues()
' ==================================================================
Dim iPos As Long
Dim iNextPos As Long
Dim iDelimLen As Long

    iCount = 0
    Erase sValues
    iDelimLen = Len(sDelim)
    iPos = 1
    iNextPos = InStr(sString, sDelim)
    Do While iNextPos > 0
        iCount = iCount + 1
        ReDim Preserve sValues(1 To iCount) As String
        sValues(iCount) = Mid$(sString, iPos, (iNextPos - iPos))
        iPos = iNextPos + iDelimLen
        iNextPos = InStr(iPos, sString, sDelim)
    Loop
    iCount = iCount + 1
    ReDim Preserve sValues(1 To iCount) As String
    sValues(iCount) = Mid$(sString, iPos)
End Sub



