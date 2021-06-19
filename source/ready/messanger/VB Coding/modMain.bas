Attribute VB_Name = "modMain"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes _
'    As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress _
'    As Long, lpParameter As Any, ByVal dwCreationFlags As Long, _
'    lpThreadId As Long) As Long
'
'Public Type SECURITY_ATTRIBUTES
'        nLength As Long
'        lpSecurityDescriptor As Long
'        bInheritHandle As Long
'End Type
Global accept As Integer
Public Prompt As Boolean
Public DisplayCorr As Boolean
Public PlaySnd As Boolean
Public strack As String
Public Const SND_ASYNC = &H1

Function ReadINI(Section As String, KeyName As String, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Function PlayWav(File As String)
Dim Play
File = App.Path & "\" & File
Play = sndPlaySound(File, SND_ASYNC)
End Function

Sub Main()
If Command = "-prompt" Then
Prompt = True
Else
Prompt = True 'False
End If

If ReadINI("DisplayCorrection", "On", App.Path & "\Options.ini") = "True" Then
    DisplayCorr = True
Else
    DisplayCorr = False
End If

If ReadINI("Sounds", "Value", App.Path & "\Options.ini") = "On" Then
    PlaySnd = True
Else
    PlaySnd = False
End If

frmSignon.Show
End Sub



