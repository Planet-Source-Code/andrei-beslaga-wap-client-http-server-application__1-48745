VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seqrat Controls Installer"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   3615
   End
   Begin VB.CommandButton btnInstall 
      Caption         =   "Install Libraries"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================
' this is the SEQRAT installer program
' it installs the needed controls and libraries from the resource file
' zlib.dll; flatbtn2.ocx; msscript.ocx
'
' portions copyright by Encina Joaquin
'
'=============================================================

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub btnInstall_Click()
If btnInstall.Caption = "Install Libraries" Then
    btnInstall.Enabled = False
    extract "zlib.dll", 101
    DoEvents
    extract "flatbtn2.ocx", 102
    DoEvents
    extract "msscript.ocx", 103
    DoEvents
    outPut "Done !"
    btnInstall.Caption = "Exit"
    btnInstall.Enabled = True
Else
    Unload Me
End If
End Sub

Public Function extract(ByVal filename As String, ByVal resnr As Integer)
Dim Ret As Boolean
Dim cfile As String
    cfile = SysDir() & "\" & filename
    If Dir(cfile) <> filename Then
        outPut "extracting " & filename & " ..."
        DoEvents
        Ret = ExtractLibs(resnr, cfile)
        If Ret = False Then Exit Function
        If LCase(Right(filename, 4)) = ".ocx" Then
            DoEvents
            outPut "registering " & filename & " ..."
            Shell "regsvr32 /s " & cfile
        End If
    Else
        outPut filename & " already installed."
    End If
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

Public Function Strip(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        Strip = Left$(strString, intZeroPos - 1)
    Else
        Strip = strString
    End If
End Function

Public Function ExtractLibs(intResNr As Integer, strPath As String) As Boolean
' ==================================================================================
' Name:         ExtractLibs
' Usage:        extracts needed libraries from a resourcefile
' Arguments:    intResNr(Integer) - identifies the resource to extract
'               strPath (String)  - destination of extracted file
' Returns:      False - function fails
'               True  - success
' Filename:     modExtractLibs.bas
' Author:       Daniel Pramel
' Date:         15 May 2002
' ==================================================================================
Dim intFileNumber As Integer
Dim bLibBuffer() As Byte
    On Error GoTo Errhandler
    bLibBuffer = LoadResData(intResNr, "CUSTOM")
    intFileNumber = FreeFile
    Open strPath For Binary Access Write As #intFileNumber
        Put #intFileNumber, , bLibBuffer
    Close #intFileNumber
    On Error GoTo 0
    ExtractLibs = True
    Exit Function
Errhandler:
    ExtractLibs = False
    outPut "error extracting library " & Str(intResNr) & " - " & strPath & " !"
End Function


Private Sub outPut(ByVal what As String)
    Text1.Text = Text1.Text & what & vbCrLf & vbCrLf
    Text1.SelStart = Len(Text1.Text)
End Sub
