VERSION 5.00
Begin VB.Form frmChat 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmChat.frx":0000
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   240
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'copyright (c)2002 - Joakoman Encina
Private Sub form_load()
On Error Resume Next
    x = SystemParametersInfo(97, True, CStr(1), 0)
    Move 0, 0, Screen.Width * Screen.TwipsPerPixelX, Screen.Height * Screen.TwipsPerPixelY
    Text1.Left = 300
    Text1.Top = 300
    Text1.Width = Me.Width - 700
    Text1.Height = Me.Height - 400
    Text1.SelStart = Len(Text1.Text)
    Image1.Left = 100
    Image1.Top = 100
    Image1.Width = Me.Width - 300
    Image1.Height = Me.Height - 200
    frmMain.FormStayOnTop Me.hwnd, True
    frmMain.HideTaskBar
End Sub

Public Sub text1_KeyDown(keycode As Integer, shift As Integer)
    If keycode = vbKeyReturn Then cmdSend
End Sub

Public Sub cmdSend()
On Error Resume Next
data = lastreturn
For soknr = 1 To 5
If frmMain.ws(soknr).State = 7 Then
    frmMain.ws(soknr).SendData "26" & encdec(sessKey(soknr), "<" + frmMain.ws(soknr).LocalHostName + "> " + data)
    sc(soknr) = 0
    Do While sc(soknr) = 0
    DoEvents
    Loop
End If
Next soknr
Text1.SelStart = Len(Text1.Text)
End Sub

Public Function lastreturn() As String
On Error Resume Next
    i = Len(Text1.Text) - 1
    Do While Mid(Text1.Text, i, 1) <> Chr(13) And i > 0
        i = i - 1
    Loop
    lastreturn = Mid(Text1.Text, i + 2, Len(Text1.Text))
End Function

Private Sub form_unload(i As Integer)
On Error Resume Next
x = SystemParametersInfo(97, False, CStr(1), 0)
Call frmMain.ShowTaskBar
End Sub
