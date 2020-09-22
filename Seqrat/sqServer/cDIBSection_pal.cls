VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_LONG = 2147483647

Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCPAINT = &HEE0086

Private m_Colors As Long

Private Type POINTAPI
  x As Long
  y As Long
End Type
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
'Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
'Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest

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
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
  RGBBlue As Byte
  RGBGreen As Byte
  RGBRed As Byte
  RGBReserved As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors() As RGBQUAD
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'Handle to the current DIBSection
Private m_hDIb As Long

'Handle to the old bitmap in the DC, for clear up
Private m_hBmpOld As Long

'Handle to the Device context holding the DIBSection
Private m_HDC As Long

'Address of memory pointing to the DIBSection's bits
Private m_lPtr As Long

'Type containing the Bitmap information
Private m_TBI As BITMAPINFO

Private Type BITMAPINFO2
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 1) As RGBQUAD
End Type
Private Type BITMAPINFO16
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 15) As RGBQUAD
End Type
Private Type BITMAPINFO256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type
Private Declare Function CreateDIBSection2 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO2, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection16 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO16, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection256 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO256, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long

Private Function BitCount() As Long

  Select Case m_Colors
  Case 2
    BitCount = 1
  Case 16
    BitCount = 4
  Case 256
    BitCount = 8
  Case 2 ^ 24
    BitCount = 24
  End Select
  
End Function

Friend Function SafeArray() As SAFEARRAY2D

  'Get the bits in the from DIB section
  With SafeArray
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = m_TBI.bmiHeader.biHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = BytesPerScanLine()
    .pvData = m_lPtr
  End With

End Function

Public Sub Clear()

  Dim x As Long
  Dim y As Long
  Dim Size As Long
  Dim bDibOrig() As Byte
  Dim tSAOrig As SAFEARRAY2D
  
  'Get the bits in the DIB section
  tSAOrig = SafeArray()
  Call CopyMemory(ByVal VarPtrArray(bDibOrig()), VarPtr(tSAOrig), 4)

  Size = tSAOrig.Bounds(0).cElements * tSAOrig.Bounds(1).cElements
  Call CopyMem(bDibOrig(0, 0), ByVal String$(Size, 0), Size)

  'Clear the temporary array descriptor
  'This is necessary under NT4.
  CopyMemory ByVal VarPtrArray(bDibOrig), 0&, 4
  
End Sub

Public Property Get Colors() As Long
  Colors = m_Colors
End Property

Public Property Let Colors(New_Value As Long)
  m_Colors = New_Value
End Property

Public Function CreateDIB(ByVal lHDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByRef hDib As Long) As Boolean

  Dim i As Long
  Dim cP As cPalette
  Dim tBI2 As BITMAPINFO2
  Dim tBI16 As BITMAPINFO16
  Dim tBI256 As BITMAPINFO256
  
  'Set up the header structure
  With m_TBI.bmiHeader
    .biSize = Len(m_TBI.bmiHeader)
    .biWidth = lWidth
    .biHeight = lHeight
    .biPlanes = 1
    .biBitCount = BitCount()
    .biCompression = BI_RGB
    .biSizeImage = BytesPerScanLine() * .biHeight
  End With
  
  'Create the color palette with the colordepth,
  'unless we use a TrueColor (2^24) where no
  'color palette is needed
  If (m_Colors <> 2 ^ 24) Then
    Set cP = New cPalette
    Call cP.CreateStatic(m_Colors)
  End If
    
  Select Case m_Colors
  Case 2
    'Set up the bitmap header
    LSet tBI2.bmiHeader = m_TBI.bmiHeader
    
    'Populate bitmap color table
    For i = 0 To 1
      With tBI2.bmiColors(i)
        .RGBBlue = cP.Blue(i + 1)
        .RGBGreen = cP.Green(i + 1)
        .RGBRed = cP.Red(i + 1)
      End With
    Next
    
    'Create the DIB
    hDib = CreateDIBSection2(lHDC, tBI2, DIB_RGB_COLORS, m_lPtr, 0, 0)
  Case 16
    'Set up the bmiHeader size
    LSet tBI16.bmiHeader = m_TBI.bmiHeader
    
    'Populate bitmap color table
    For i = 0 To 15
      With tBI16.bmiColors(i)
        .RGBBlue = cP.Blue(i + 1)
        .RGBGreen = cP.Green(i + 1)
        .RGBRed = cP.Red(i + 1)
      End With
    Next
    
    'Create the DIB
    hDib = CreateDIBSection16(lHDC, tBI16, DIB_RGB_COLORS, m_lPtr, 0, 0)
  Case 256
    'Copy the bitmap header
    LSet tBI256.bmiHeader = m_TBI.bmiHeader
    
    'Populate bitmap color table
    For i = 0 To 255
      With tBI256.bmiColors(i)
        .RGBBlue = cP.Blue(i + 1)
        .RGBGreen = cP.Green(i + 1)
        .RGBRed = cP.Red(i + 1)
      End With
    Next
    
    'Create the DIB
    hDib = CreateDIBSection256(lHDC, tBI256, DIB_RGB_COLORS, m_lPtr, 0, 0)
  Case Else
    'Create a DIB with TrueColor
    hDib = CreateDIBSection(lHDC, m_TBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
  End Select
  
  'Return TRUE if the DIB was created
  CreateDIB = (hDib <> 0)

End Function

Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
  Call ClearUp
  m_HDC = CreateCompatibleDC(0)
  If (m_HDC <> 0) Then
    If (CreateDIB(m_HDC, lWidth, lHeight, m_hDIb)) Then
      m_hBmpOld = SelectObject(m_HDC, m_hDIb)
      Create = True
    Else
      DeleteObject m_HDC
      m_HDC = 0
    End If
  End If

End Function

Public Property Get BytesPerScanLine() As Long
  
  BytesPerScanLine = (Width * BitCount)
  If (BytesPerScanLine Mod 32 > 0) Then
    BytesPerScanLine = BytesPerScanLine + 32 - (BytesPerScanLine Mod 32)
  End If
  BytesPerScanLine = BytesPerScanLine \ 8
  
'  'Scans must align on dword boundaries
'  Select Case m_Colors
'  Case 2
'    BytesPerScanLine = (Width + 7) \ 8
'  Case 16
'    BytesPerScanLine = (Width + 1) \ 2
'  Case 256
'    BytesPerScanLine = Width
'  Case Else
'    BytesPerScanLine = Width * 3
'  End Select
'
'  'The value must be dword-aligned
'  If (BytesPerScanLine Mod 4 > 0) Then
'    BytesPerScanLine = BytesPerScanLine + (4 - (BytesPerScanLine Mod 4))
'  End If

End Property

Public Function Size() As Long
  Size = m_TBI.bmiHeader.biSizeImage
End Function
Public Function ToByte(Value() As Byte)

  Dim y As Integer
  Dim bDibFrom() As Byte
  Dim tSAFrom As SAFEARRAY2D
  
  'Store the data (width - height - data)
  ReDim Value(Size - 1 + 8)
  Call CopyMem(Value(0), Width, 4)
  Call CopyMem(Value(4), Height, 4)
  If (Size > 0) Then
  'Get the bits in the DIB section
  With tSAFrom
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = m_TBI.bmiHeader.biHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = BytesPerScanLine()
    .pvData = m_lPtr
  End With
  CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4
  
    Call CopyMem(Value(8), bDibFrom(0, 0), Size)
  
  'Clear the temporary array descriptor
  'This is necessary under NT4.
  CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
  End If

End Function

Public Property Get Width() As Long
  Width = m_TBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
  Height = m_TBI.bmiHeader.biHeight
End Property

Public Property Get hdc() As Long
  hdc = m_HDC
End Property
Public Property Get hDib() As Long
  hDib = m_hDIb
End Property
Public Property Get DIBSectionBitsPtr() As Long
  DIBSectionBitsPtr = m_lPtr
End Property

Public Sub ClearUp()
  If (m_HDC <> 0) Then
    If (m_hDIb <> 0) Then
      SelectObject m_HDC, m_hBmpOld
      DeleteObject m_hDIb
    End If
    DeleteObject m_HDC
  End If
  m_HDC = 0
  m_hDIb = 0
  m_hBmpOld = 0
  m_lPtr = 0
End Sub

Private Sub Class_Initialize()
  m_Colors = 2 ^ 24
End Sub

Private Sub Class_Terminate()
  ClearUp
End Sub
