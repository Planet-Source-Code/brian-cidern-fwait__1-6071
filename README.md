<div align="center">

## fWait


</div>

### Description

Shells an app, then waits for that app to close before it continues processing.
 
### More Info
 
None --

Pseudo code:

Uses API to get the OS dir (for 95/98/NT4/2000 compatability) and appends result with Notepad.exe. Shells Notepad, returning process id. fWait gets the app hdl and issues a Do Events until the exit code of the app <> STILL_ACTIVE&. When app is closed, a cheezy MsgBox displays.

Create a Std EXE. Add a command button, and use the default name (Command1).

Shelled app exit code


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Cidern](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-cidern.md)
**Level**          |Advanced
**User Rating**    |4.5 (50 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-cidern-fwait__1-6071/archive/master.zip)

### API Declarations

```
Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&
Const INFINITE& = &HFFFF
Private Declare Function GetWindowsDirectory _
  Lib "kernel32" _
  Alias "GetWindowsDirectoryA" ( _
  ByVal lpBuffer As String, _
  ByVal nSize As Long _
  ) As Long
Private Declare Function OpenProcess _
  Lib "kernel32" ( _
  ByVal dwDesiredAccess As Long, _
  ByVal bInheritHandle As Long, _
  ByVal dwProcessId As Long _
  ) As Long
Private Declare Function WaitForSingleObject _
  Lib "kernel32" ( _
  ByVal hHandle As Long, _
  ByVal dwMilliseconds As Long _
  ) As Long
Private Declare Function GetExitCodeProcess _
  Lib "kernel32" ( _
  ByVal hProcess As Long, _
  lpExitCode As Long _
  ) As Long
Private Declare Function CloseHandle _
  Lib "kernel32" ( _
  ByVal hObject As Long _
  ) As Long
```


### Source Code

```
Private Sub Command1_Click()
  Dim sCmdLine As String
  Dim idProg As Long, iExit As Long
  sCmdLine = fGetWinDir & "\notepad.exe"
  idProg = Shell(sCmdLine)
  iExit = fWait(idProg)
  If iExit Then
    MsgBox "Something very, very bad just happened."
  Else
    MsgBox "Finished processing Notepad."
  End If
End Sub
Function fWait(ByVal lProgID As Long) As Long
  ' Wait until proggie exit code <> STILL_ACTIVE&
  Dim lExitCode As Long, hdlProg As Long
  ' Get proggie handle
  hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, lProgID)
  ' Get current proggie exit code
  GetExitCodeProcess hdlProg, lExitCode
  Do While lExitCode = STILL_ACTIVE&
    DoEvents
    GetExitCodeProcess hdlProg, lExitCode
  Loop
  CloseHandle hdlProg
  fWait = lExitCode
End Function
Private Function fGetWinDir() As String
  ' Wrapper to return OS Path
  Dim lRet As Long, lSize As Long, sBuf As String * 512
  lSize = 512
  lRet = GetWindowsDirectory(sBuf, lSize)
  fGetWinDir = Left(sBuf, InStr(1, sBuf, Chr(0)) - 1)
End Function
```

