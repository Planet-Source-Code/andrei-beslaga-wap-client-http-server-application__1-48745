VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmScreen 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   3075
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   ScaleHeight     =   3075
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1560
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock wsMouseK 
      Left            =   960
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsScreen 
      Left            =   360
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox pVideo 
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020            ' (DWORD) dest = source

Dim RecDib As New cDIBSection
Dim ZLib As New clsZLib

Dim mouseK As Byte
Dim keyK As Byte

Dim CRate As Byte
Dim iCOLORS As Integer


Private Sub form_load()
On Error Resume Next
Dim tmpParm() As String

    capScreen = 1
    sdx = -10000
    sdy = -10000
    sc = 0
    
    'strOptSend = x;y;z;m
    'timetowait ; crate(3-10) ; colors(2,16,256) ; mousekbon
    
    frmMain.ws.SendData "sb" & strOptSend
    tmpParm = Split(strOptSend, ";")
    
    CRate = CByte(tmpParm(1))
    iCOLORS = CInt(tmpParm(2))
    
    Do While sc = 0
        DoEvents
    Loop
    
    Do While sdx = -10000 And sdy = -10000
        DoEvents
    Loop
    Me.Top = 0
    Me.Left = 0
    Me.Width = Me.ScaleX(sdx, vbPixels, vbTwips)
    Me.Height = Me.ScaleY(sdy, vbPixels, vbTwips)
    pVideo.Width = Me.Width
    pVideo.Height = Me.Height

    wsScreen.RemoteHost = frmMain.Text3.Text
    wsScreen.RemotePort = 17090
    wsScreen.Connect
    Do While wsScreen.State <> sckConnected
        DoEvents
    Loop
       
    RecDib.Colors = iCOLORS
    Call RecDib.Create(sdx / CRate, sdy / CRate)
    'i = 0

    If tmpParm(3) <> "n" Then
        wsMouseK.RemoteHost = frmMain.Text3.Text
        wsMouseK.RemotePort = 17091
        wsMouseK.Connect
        Do While wsMouseK.State <> sckConnected
            DoEvents
        Loop
        mouseK = 1
    End If

    Timer1.Interval = 100
    Timer1.Enabled = True
    
End Sub
Private Sub form_unload(cancel As Integer)
On Error Resume Next
    sc = 0
    frmMain.ws.SendData "se"
    Do While sc = 0
        DoEvents
    Loop
    Timer1.Enabled = False
    If wsScreen.State <> sckClosed Then wsScreen.Close
    If wsMouseK.State <> sckClosed Then wsMouseK.Close
    Unload Me
End Sub


Private Sub timer1_timer()
    If wsMouseK.State = sckConnected And mouseK = 1 Then
        mouseK = 0
        wsMouseK.SendData "sp" & Str(GetX) & Chr(0) & Str(GetY)
    End If
End Sub

Private Sub pVideo_MouseDown(Button As Integer, _
      Shift As Integer, x As Single, y As Single)
On Error Resume Next

Dim mKstr
If wsMouseK.State = sckConnected And mouseK = 1 Then
    mouseK = 0
    
    If Button = 1 Then
        Select Case Shift
            Case 6: mKstr = "ld"
            Case 7: mKstr = "lu"
            Case Else: mKstr = "lc"
        End Select
    End If
    If Button = 2 Then
        Select Case Shift
            Case 6: mKstr = "rd"
            Case 7: mKstr = "ru"
            Case Else: mKstr = "rc"
        End Select
    End If
    If Button = 4 Then
        Select Case Shift
            Case 6: mKstr = "md"
            Case 7: mKstr = "mu"
            Case Else: mKstr = "mc"
        End Select
    End If

    wsMouseK.SendData mKstr

End If
End Sub

Private Sub form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim ShiftDown, AltDown, CtrlDown
Dim sts As String
Dim sk

'If mouseK = 0 Or wsMouseK = sckClosed Then Exit Sub

mouseK = 0

ShiftDown = (Shift And vbShiftMask) > 0
AltDown = (Shift And vbAltMask) > 0
CtrlDown = (Shift And vbCtrlMask) > 0

sts = sts & IIf(ShiftDown, "+", " ")
sts = sts & IIf(AltDown, "%", " ")
sts = sts & IIf(ControlDown, "^", " ")

    Select Case KeyCode
       Case vbKeyF10:
            If CtrlDown And AltDown Then
                   capScreen = 0
                   Unload Me
            Else: GoTo keyile
            End If
       Case vbKeyF9:
            If CtrlDown And AltDown Then
                wsMouseK.SendData "dc"
            Else: GoTo keyile
            End If
       Case vbKeyF1:
            If CtrlDown And AltDown Then
                wsMouseK.Close
            Else: GoTo keyile
            End If
       Case vbKeyF2:
            If CtrlDown And AltDown Then
                wsMouseK.Connect
            Else: GoTo keyile
            End If
        
        Case Else:
keyile:                 Select Case KeyCode
                        Case 8: sk = "{BKSP}"
                        Case 13: sk = "{ENTER}"
                        Case 27: sk = "{ESC}"
                        Case 38: sk = "{UP}"
                        Case 40: sk = "{DOWN}"
                        Case 37: sk = "{LEFT}"
                        Case 39: sk = "{RIGHT}"
                        Case 19: sk = "{BREAK}"
                        Case 35: sk = "{END}"
                        Case 9: sk = "{TAB}"
                        Case 145: sk = "{SCROLLLOCK}"
                        Case 144: sk = "{NUMLOCK}"
                        Case 33: sk = "{PGUP}"
                        Case 34: sk = "{PGDN}"
                        Case 45: sk = "{INSERT}"
                        Case 36: sk = "{HOME}"
                        Case 46: sk = "{DEL}"
                        Case 20: sk = "{CAPSLOCK}"
                        Case 112: sk = "{F1}"
                        Case 113: sk = "{F2}"
                        Case 114: sk = "{F3}"
                        Case 115: sk = "{F4}"
                        Case 116: sk = "{F5}"
                        Case 117: sk = "{F6}"
                        Case 118: sk = "{F7}"
                        Case 119: sk = "{F8}"
                        Case 120: sk = "{F9}"
                        Case 121: sk = "{F10}"
                        Case 122: sk = "{F11}"
                        Case 123: sk = "{F12}"
                        Case 32: sk = " "
                        Case 48: sk = IIf(ShiftDown, ")", "0")
                        Case 49: sk = IIf(ShiftDown, "!", "1")
                        Case 50: sk = IIf(ShiftDown, "@", "2")
                        Case 51: sk = IIf(ShiftDown, "#", "3")
                        Case 52: sk = IIf(ShiftDown, "$", "4")
                        Case 53: sk = IIf(ShiftDown, "{%}", "5")
                        Case 54: sk = IIf(ShiftDown, "{^}", "6")
                        Case 55: sk = IIf(ShiftDown, "&", "7")
                        Case 56: sk = IIf(ShiftDown, "*", "8")
                        Case 57: sk = IIf(ShiftDown, "(", "9")
                        Case &H60 To &H69: sk = CStr(KeyCode - &H60) 'Numpad 0-9
                        Case 65 To 90: sk = IIf(ShiftDown, UCase$(Chr$(KeyCode)), LCase$(Chr$(KeyCode))) 'a-z
                        Case 186: sk = IIf(ShiftDown, ":", ";") ';
                        Case 187: sk = IIf(ShiftDown, "{+}", "=") '=
                        Case 188: sk = IIf(ShiftDown, "<", ",") ',
                        Case 189: sk = IIf(ShiftDown, "_", "-") '-
                        Case 190: sk = IIf(ShiftDown, ">", ".") '.
                        Case 191: sk = IIf(ShiftDown, "?", "/") '/
                        Case 192: sk = IIf(ShiftDown, "{~}", "`") '`
                        Case 219: sk = IIf(ShiftDown, "{{}", "{[}") '[
                        Case 220: sk = IIf(ShiftDown, "|", "\") '\
                        Case 221: sk = IIf(ShiftDown, "{}}", "{]}") ']
                        Case 222: sk = IIf(ShiftDown, Chr$(34), "'") ''
                    End Select
                    sts = sts & sk
                    wsMouseK.SendData "lk" & sts
    End Select
End Sub


Private Sub wsScreen_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errore:
Dim Ret As Long
Dim ByteArray() As Byte
'### okay, get the compressed bytearray
    wsScreen.GetData ByteArray, vbByte
'### decompress it
    Call ZLib.DecompressByte(ByteArray)
    DoEvents
'### make a dib from the bytearray
    Call RecDib.ParseByte(ByteArray)
'### and finally blit it to the actual position
    Ret = BitBlt(pVideo.hdc, scrPos(0), scrPos(1), sdx / CRate, sdy / CRate, RecDib.hdc, 0, 0, SRCCOPY)
'### send some reply - it doesn't matter what you send
'### i think it would be better to send one byte
    If capScreen = 1 Then
        frmMain.ws.SendData Chr(0)
    Else
        frmMain.ws.SendData "se"
    End If
errore:
End Sub


Private Sub wsMouseK_dataArrival(ByVal bytes As Long)
    mouseK = 1
End Sub

