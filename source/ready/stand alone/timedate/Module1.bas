Attribute VB_Name = "Module1"
Public AlarmHr As Integer  'hold railway time
Public AlarmMin As Integer
Public AlarmSec As Integer
Public Alarm As Boolean 'true when alarm sett
Public playFile As String
Public exeFile As String
Public runExe As Boolean
Public alarmMsg As String
Public isplaySound As Boolean 'play alarm if true else only show message
'''''''' window top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_SHOWWINDOW = &H40
'shell
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
