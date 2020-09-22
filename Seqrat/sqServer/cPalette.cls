VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal lpvDest As Long, lpvSource As Any, ByVal cbCopy As Long)

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

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias _
   "VarPtr" (Ptr() As Any) As Long
Private Type tNode                              'Octree node struct
    bIsLeaf As Boolean                          'Leaf flag
    bAddedReduce As Boolean                     'Linked list flag
    vRed As Long                                'Red Value
    vGreen As Long                              'Green Value
    vBlue As Long                               'Blue Value
    cClrs As Long                               'Pixel count
    iChildren(0 To 1, 0 To 1, 0 To 1) As Long   'Child pointers
    iNext As Long                               'Next reducable node
End Type

Private aNodes() As tNode, cNodes As Long
Private nDepth As Byte, TopGarbage As Long
Private cClr As Long, aReduce(1 To 8) As Long
Private cOps As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

' make private storage in compatible
' with windows storage in memory/on disk
Private Type RGBQUAD
   Red As Byte
   Green As Byte
   Blue As Byte
   Reserved As Byte
End Type
Private m_tPal() As RGBQUAD
Private m_iPalette As Long

Public Event InitProgress(ByVal Max As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)

Public Sub Add(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
    m_iPalette = m_iPalette + 1
    ReDim Preserve m_tPal(1 To m_iPalette) As RGBQUAD
    m_tPal(m_iPalette).Red = Red
    m_tPal(m_iPalette).Green = Green
    m_tPal(m_iPalette).Blue = Blue
End Sub

Public Property Get ClosestIndex( _
      ByVal Red As Long, _
      ByVal Green As Long, _
      ByVal Blue As Long _
   ) As Long
Dim i As Long
Dim lER As Long, lEB As Long, lEG As Long
Dim lMinER As Long, lMinEB As Long, lMinEG As Long
Dim lMinIndex As Long
   
   lMinER = 255: lMinEB = 255: lMinEG = 255
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
            End If
         End If
      End With
   Next i
   
   ClosestIndex = lMinIndex
      
End Property

Public Property Get Count() As Long
' Number of colours in this palette:
   Count = m_iPalette
End Property
Public Sub CreateStatic(ColorDepth As Long)

  Select Case ColorDepth
  Case 2
    Call CreateMono
  Case 16
    Call Create16Colour
  Case 256
    Call CreateHalfTone
  Case Else
    Call Err.Raise(vbObjectError, , "clsPalette: Invalid ColorDepth for color palette")
  End Select
  
End Sub
Public Property Get Red(ByVal iIndex) As Byte
' Red value for palette index
   Red = m_tPal(iIndex).Red
End Property
Public Property Get Green(ByVal iIndex) As Byte
' Green value for palette index
   Green = m_tPal(iIndex).Green
End Property
Public Property Get Blue(ByVal iIndex) As Byte
' Blue value for palette index
   Blue = m_tPal(iIndex).Blue
End Property

Public Sub CreateHalfTone()
Dim lIndex As Long
Dim r As Long, g As Long, b As Long
Dim rA As Long, gA As Long, bA As Long
Dim l As Long, i As Long

   ' Halftone 256 colour palette
   m_iPalette = 256
   ReDim m_tPal(1 To 256) As RGBQUAD
   For b = 0 To &H100 Step &H40
      If b = &H100 Then
         bA = b - 1
      Else
         bA = b
      End If
      For g = 0 To &H100 Step &H40
         If g = &H100 Then
            gA = g - 1
         Else
            gA = g
         End If
         For r = 0 To &H100 Step &H40
            If r = &H100 Then
               rA = r - 1
            Else
               rA = r
            End If
            lIndex = lIndex + 1
            With m_tPal(lIndex)
               .Red = rA: .Green = gA: .Blue = bA
            End With
         Next r
      Next g
   Next b
   
End Sub
Public Sub CreateMono()
   ' Monochrome palette
   m_iPalette = 2
   ReDim m_tPal(1 To 2) As RGBQUAD
   With m_tPal(2)
      .Blue = 255
      .Green = 255
      .Red = 255
   End With
End Sub
Private Sub p16ColourLow8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 1
   With m_tPal(lStartIndex + 2)
      .Red = &H80: .Green = 0: .Blue = 0
   End With
   With m_tPal(lStartIndex + 3)
      .Red = 0: .Green = &H80: .Blue = 0
   End With
   With m_tPal(lStartIndex + 4)
      .Red = &H80: .Green = &H80: .Blue = 0
   End With
   With m_tPal(lStartIndex + 5)
      .Red = 0: .Green = 0: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 6)
      .Red = &H80: .Green = 0: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 7)
      .Red = 0: .Green = &H80: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 8)
      .Red = &HC0: .Green = &HC0: .Blue = &HC0
   End With

End Sub
Private Sub p16ColourHigh8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 9
   With m_tPal(lStartIndex + 9)
      .Red = &H80: .Green = &H80: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 10)
      .Red = &HFF: .Green = 0: .Blue = 0
   End With
   With m_tPal(lStartIndex + 11)
      .Red = 0: .Green = &HFF: .Blue = 0
   End With
   With m_tPal(lStartIndex + 12)
      .Red = &HFF: .Green = &HFF: .Blue = 0
   End With
   With m_tPal(lStartIndex + 13)
      .Red = 0: .Green = 0: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 14)
      .Red = &HFF: .Green = 0: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 15)
      .Red = 0: .Green = &HFF: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 16)
      .Red = &HFF: .Green = &HFF: .Blue = &HFF
   End With
End Sub
Public Sub Create16Colour()
   ' Standard EGA style 16 colour palette:
   m_iPalette = 16
   ReDim m_tPal(1 To 16) As RGBQUAD
   p16ColourLow8 1
   p16ColourHigh8 9
End Sub

Public Sub CreateOptimal( _
      ByRef cDIB As cDIBSection, _
      Optional ByVal PalSize As Long = 236, _
      Optional ByVal nLevels As Long = 6 _
   )
Dim tSA As SAFEARRAY2D, aDib() As Byte
Dim x As Long, y As Long
Dim xMax As Long, yMax As Long
Dim sTime As Long, pPal As Double
Dim Inc As Double

   'Creates a VB array without copying data
   With tSA
       .cbElements = 1
       .cDims = 2
       .Bounds(0).lLbound = 0
       .Bounds(0).cElements = cDIB.Height
       .Bounds(1).lLbound = 0
       .Bounds(1).cElements = cDIB.BytesPerScanLine
       .pvData = cDIB.DIBSectionBitsPtr
   End With
   CopyMemory VarPtrArray(aDib), VarPtr(tSA), 4
   
   xMax = (cDIB.Width - 1) * 3
   Debug.Print "hptr = " & Hex(cDIB.DIBSectionBitsPtr)
   Debug.Print "xMax = " & xMax
   Debug.Print "byte = " & cDIB.BytesPerScanLine
   Debug.Print "ubound2 = " & UBound(aDib, 2)
   Debug.Print "ubound1 = " & UBound(aDib, 1)
   
   yMax = cDIB.Height - 1
   
   nDepth = nLevels
   
   sTime = timeGetTime
   
   ReDim aNodes(1 To 50) 'Allocates initial storage
   TopGarbage = 0
   cNodes = 1
   cClr = 0
   
   RaiseEvent InitProgress(yMax + PalSize)
   
   For y = 0 To yMax
      For x = 0 To xMax
         'Adds the current pixel to the color octree
          AddClr 1, 1, 0, 255, 0, 255, 0, 255, _
            aDib(x + 2, y), aDib(x + 1, y), aDib(x, y)
         
         pPal = -1
         Do While cClr > PalSize
            'Combine the levels to get down to desired palette size
            pPal = CombineNodes
            If pPal = False Then Exit Do
         Loop
         If pPal = False Then
            If MsgBox("Error in reducing nodes", vbYesNo) = vbNo _
               Then Stop
         End If
      Next x
      RaiseEvent Progress(y)
   Next y
   
   Clear
   Debug.Print "Colors: " & cClr
   pPal = 0
   Inc = pPal / cNodes
   'Go through octree and extract colors
   For y = 1 To UBound(aNodes)
      If aNodes(y).bIsLeaf Then
         With aNodes(y)
            pPal = pPal + 1
            Add .vRed / .cClrs, .vGreen / .cClrs, .vBlue / .cClrs
            RaiseEvent Progress(y)
         End With
      End If
   Next y
   
   Debug.Print pPal
   
   RaiseEvent Complete(ByVal (timeGetTime - sTime))

End Sub

Private Sub AddClr(ByVal iBranch As Long, ByVal nLevel As Long, _
      ByVal vMinR As Byte, ByVal vMaxR As Byte, _
      ByVal vMinG As Byte, ByVal vMaxG As Byte, _
      ByVal vMinB As Byte, ByVal vMaxB As Byte, _
      ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
        
Dim iR As Byte, iG As Byte, iB As Byte
Dim vMid As Long, iIndex As Long

   
   'Find mid values for colors and decide which path to take
   'Also update max and min values for later call to self
   vMid = vMinR / 2 + vMaxR / 2
   If r > vMid Then
      iR = 1
      vMinR = vMid
   Else
      iR = 0
      vMaxR = vMid
   End If
   
   vMid = vMinG / 2 + vMaxG / 2
   If g > vMid Then
      iG = 1
      vMinG = vMid
   Else
      iG = 0
      vMaxG = vMid
   End If
   
   vMid = vMinB / 2 + vMaxB / 2
   If b > vMid Then
      iB = 1
      vMinB = vMid
   Else
      iB = 0
      vMaxB = vMid
   End If
   
   If aNodes(iBranch).iChildren(iR, iG, iB) = 0 Then  'If no child here then...
      iIndex = GetFreeNode 'Get a new node index
      aNodes(iBranch).iChildren(iR, iG, iB) = iIndex
      aNodes(iBranch).cClrs = aNodes(iBranch).cClrs + 1
      With aNodes(iIndex)   'Clear/set data
         .bIsLeaf = (nLevel = nDepth): .vGreen = 0
         .iNext = 0: .cClrs = 0: .vBlue = 0: .vRed = 0
      End With
   Else  'Has a child here
      iIndex = aNodes(iBranch).iChildren(iR, iG, iB)
   End If
   
   
   If aNodes(iIndex).bIsLeaf Then  'If it is a leaf
      With aNodes(iIndex)
         If .cClrs = 0 Then cClr = cClr + 1
         .cClrs = .cClrs + 1
         .vRed = .vRed + r
         .vGreen = .vGreen + g
         .vBlue = .vBlue + b
      End With
   Else  'Otherwise
      With aNodes(iIndex)
         If .bAddedReduce = False Then    'If 2 or more colors, add to reducable aNodes list
            .iNext = aReduce(nLevel)
            aReduce(nLevel) = iIndex
            .bAddedReduce = True
            'Debug.Print "Reducable Added: " & iIndex, "Level = " & nLevel, "Top was = " & .iNext
         End If
      End With
      
      'Search a level deeper
      AddClr iIndex, nLevel + 1, vMinR, vMaxR, vMinG, vMaxG, _
         vMinB, vMaxB, r, g, b
   End If
End Sub

Private Function CombineNodes() As Boolean
Dim i As Long, iIndex As Long
Dim iR As Byte, iG As Byte, iB As Byte
Dim nR As Long, nG As Long, nB As Long, nPixel As Long
   
   
   
   For i = nDepth To 1 Step -1   'Find deepest reducable level
      If aReduce(i) <> 0 Then Exit For
   Next i
   
   'Debug.Print "aReduce Index = " & i
   
   If i = 0 Then Exit Function
   iIndex = aReduce(i)
   aReduce(i) = aNodes(iIndex).iNext
   
   For i = 0 To 7
      'iR = IIf((i And 1) = 1, 1, 0)
      If (i And 1) = 1 Then
         iR = 1
      Else
         iR = 0
      End If
      'iG = IIf((i And 2) = 2, 1, 0)
      If (i And 2) = 2 Then
         iG = 1
      Else
         iG = 0
      End If
      'iB = IIf((i And 4) = 4, 1, 0)
      If (i And 4) = 4 Then
         iB = 1
      Else
         iB = 0
      End If
      
      If aNodes(iIndex).iChildren(iR, iG, iB) <> 0 Then 'If there is a child
         With aNodes(aNodes(iIndex).iChildren(iR, iG, iB))
            'Add red, green, blue, and pixel count to running total
            nR = nR + .vRed
            nG = nG + .vGreen
            nB = nB + .vBlue
            nPixel = nPixel + .cClrs
            FreeNode aNodes(iIndex).iChildren(iR, iG, iB) 'Free the node
            cClr = cClr - 1
         End With
         aNodes(iIndex).iChildren(iR, iG, iB) = 0 'Clear the link
      End If
   Next i
   
   cClr = cClr + 1
   
   'Set the new node data
   With aNodes(iIndex)
      .cClrs = nPixel
      .bIsLeaf = True
      .vRed = nR
      .vBlue = nB
      .vGreen = nG
   End With
   CombineNodes = True
   
End Function

Public Sub EraseNodes()
  Erase aNodes
  TopGarbage = 0
End Sub

Private Sub FreeNode(ByVal iNode As Long)
   aNodes(iNode).iNext = TopGarbage
   TopGarbage = iNode
   aNodes(iNode).bIsLeaf = False 'Necessary for final loop through
   aNodes(iNode).bAddedReduce = False
   cNodes = cNodes - 1
End Sub

Private Function GetFreeNode() As Long
Dim i As Long, iR As Byte, iG As Byte, iB As Byte
   cNodes = cNodes + 1
   If TopGarbage = 0 Then
      If cNodes > UBound(aNodes) Then
         i = cNodes * 1.1
         ReDim Preserve aNodes(1 To i)
      End If
      GetFreeNode = cNodes
   Else
      GetFreeNode = TopGarbage
      TopGarbage = aNodes(TopGarbage).iNext
      For i = 0 To 7
         'iR = IIf((i And 1) = 1, 1, 0)
         If (i And 1) = 1 Then
            iR = 1
         Else
            iR = 0
         End If
         'iG = IIf((i And 2) = 2, 1, 0)
         If (i And 2) = 2 Then
            iG = 1
         Else
            iG = 0
         End If
         'iB = IIf((i And 4) = 4, 1, 0)
         If (i And 4) = 4 Then
            iB = 1
         Else
            iB = 0
         End If
         aNodes(GetFreeNode).iChildren(iR, iG, iB) = 0
      Next i
   End If
End Function

Public Sub Clear()
   Erase m_tPal
   m_iPalette = 0
End Sub

Private Sub SplitDelimitedString( _
        ByVal sString As String, _
        ByVal sDelim As String, _
        ByRef sValues() As String, _
        ByRef iCount As Long _
    )
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





