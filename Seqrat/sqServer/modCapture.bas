Attribute VB_Name = "modCapture"
Option Explicit
'### Blitten
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020            ' (DWORD) dest = source
Private Const WHITENESS = &HFF0062          ' (DWORD) dest = WHITE
'### Pixel ermitteln
'Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'### Desktop
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private DIB As New cDIBSection
Private RecDib As New cDIBSection
Private ZLib As New clsZLib

Public Sub DoCapture(ByVal index As Integer, ByVal afterDark As Single, CRATE As Byte, iCOLORS As Integer)
Dim xPos As Long                '### aktuelle Breiteninformation/ actual X-position
Dim yPos As Long                '### aktuelle Höhenposition/ actual Y-position
Dim Ret As Long                 '### Rückgabewert der API's/ returnvalue from API's
Dim DeskHwnd As Long            '### Handle des Fensters/windowhanlde
Dim DeskHdc As Long             '### Handle des Gerätekontextes/handle of devicecontext
Dim DeskRect As RECT            '### Abmessungen des Desktops/rect's of the desktop
Dim dibHdc As Long              '### i tested something
Dim ByteArray() As Byte
Dim sValue As String

Dim CS() As Long   '### Array für Prüfsummen/ if store all
                                        '### checksums of the last picture there
                                        '### so i know if the picture is different
                                        '### from the last and if i have to send it
                                        '### a second time or not
                                    
Dim CS_Tmp As Long              '### temp. Zwischensumme/checksum(bad translation)
Dim K As Long                   '### actual part of the desktop
Dim start As Single


'Set DIB = New cDIBSection
'Set RecDib = New cDIBSection
'Set ZLib = New clsZLib

    ReDim CS(CRATE * CRATE * CRATE)
    '### get desktophandle
    DeskHwnd = GetDesktopWindow()
    '### get devicecontext
    DeskHdc = GetDC(DeskHwnd)
    '### get windowrect's
    Ret = GetWindowRect(DeskHwnd, DeskRect)
    '### create 16 clolored DIB
    DIB.Colors = iCOLORS
    RecDib.Colors = iCOLORS
    Call DIB.Create(DeskRect.Right / CRATE, DeskRect.Bottom / CRATE)
    Call RecDib.Create(DeskRect.Right / CRATE, DeskRect.Bottom / CRATE)
    '### when ENDE is true then capturing shall end
    '### Arrayposition für Prüfsummen neu setzen
    '### begin with first part of the desktop
    K = 0
    '### set reponse to FALSE
    'modCapture.C_Response = False
    'modCapture.C_Set_Response = False
    Do While capScreen = 1
        For yPos = 0 To DeskRect.Bottom Step (DeskRect.Bottom / CRATE)
            For xPos = 0 To DeskRect.Right Step (DeskRect.Right / CRATE)
                    '### aktuellen Ausschnitt des Desktops in DIB blitten
                    '### blit actual part of the desktop into the dib
                    Ret = BitBlt(DIB.hdc, 0, 0, DeskRect.Right / CRATE, DeskRect.Bottom / CRATE, DeskHdc, xPos, yPos, SRCCOPY)
                    '### in Bytearray konvertieren
                    '### stire the dib in an array
                    Call DIB.ToByte(ByteArray)
                    '### Bytearray komprimieren
                    '### compress the array
                    Call ZLib.CompressByte(ByteArray)
                    '### Prüfsumme errechnen
                    '### save the checksum
                    CS_Tmp = UBound(ByteArray)
                    '### wenn anders als letzte, dann Daten senden
                    '### if the part is different to the last-> send the data
                    If CS_Tmp <> CS(K) Then
                        CS(K) = CS_Tmp
                        On Error GoTo NoConn
                        '### Positionsdaten des Ausschnitts senden
                        '### first send the actual position
                        scReply = 0
                        frmMain.ws(index).SendData "xy" & CStr(xPos) & ";" & CStr(yPos)
                        '### auf Antwort warten
                        '### wait for response
                        Do While scReply = 0 And capScreen = 1
                            DoEvents
                        Loop
                        'C_Set_Response = False
                        '### Ausschnitt senden
                        '### send data
                        scReply = 0
                        frmMain.wsScreen.SendData ByteArray
                        '### auf Antwort warten
                        '### wait for response
                        Do While scReply = 0 And capScreen = 1
                            DoEvents
                        Loop
                        'C_Response = False
                        On Error GoTo 0
                    End If
                    '### Arrayposition für Prüfsummen neu setzen
                    '### next part of the desktop...
                    K = K + 1
                    DoEvents
            Next xPos
        Next yPos
        '### Ausschnittposition neu setzen
        '### begin at pos (0,0)
        xPos = 0
        yPos = 0
        
        K = 0
        '### einen FPS dazuzählen
        '### one frame made
        'Q = Q + 1
        start = Timer
        Do While Timer < start + afterDark
            DoEvents
        Loop
    Loop
    'Exit Sub
NoConn:
    '### an error occured and because of that: close all and go online again
    Set DIB = Nothing
    Set RecDib = Nothing
    Set ZLib = Nothing
End Sub
