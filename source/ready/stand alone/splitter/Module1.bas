Attribute VB_Name = "Module1"
Option Explicit
Private Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function ShowHelp(strTopic As String, bIsLocal As Boolean) As Boolean
Dim strDir As String
If bIsLocal Then

' Get registry entry pointing to Help
strDir = App.Path + "\Help\"

End If

' Launch topic
Dim hinst As Long
hinst = ShellExecute(frmSplit.hwnd, vbNullString, strTopic, vbNullString, strDir, SW_SHOWNORMAL)

' Handle less than 32 indicates failure
ShowHelp = hinst > 32

End Function
