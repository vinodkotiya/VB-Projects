VERSION 5.00
Begin VB.Form frmShut 
   Caption         =   "ShutDown PC By Vinod Kotiya"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmShut.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shut"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmShut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''create cshortcut
'NOTE: In Visual Basic 5.0, change Stkit432.dll in the following
      'statement to Vb5stkit.dll.

      Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
       lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
       lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

''''''''''''''''
Private Type LUID
         UsedPart As Long
         IgnoredForNowHigh32BitPart As Long
      End Type

      Private Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        TheLuid As LUID
        Attributes As Long
      End Type

      Private Const EWX_SHUTDOWN As Long = 1
      Private Const EWX_FORCE As Long = 4
      Private Const EWX_REBOOT = 2

      Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
           dwOptions As Long, ByVal dwReserved As Long) As Long

      Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
      Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
         ProcessHandle As Long, _
         ByVal DesiredAccess As Long, TokenHandle As Long) As Long
      Private Declare Function LookupPrivilegeValue Lib "advapi32" _
         Alias "LookupPrivilegeValueA" _
         (ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
         As LUID) As Long
      Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
         (ByVal TokenHandle As Long, _
         ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
         , ByVal BufferLength As Long, _
      PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
      Private Sub AdjustToken()
         Const TOKEN_ADJUST_PRIVILEGES = &H20
         Const TOKEN_QUERY = &H8
         Const SE_PRIVILEGE_ENABLED = &H2
         Dim hdlProcessHandle As Long
         Dim hdlTokenHandle As Long
         Dim tmpLuid As LUID
         Dim tkp As TOKEN_PRIVILEGES
         Dim tkpNewButIgnored As TOKEN_PRIVILEGES
         Dim lBufferNeeded As Long

         hdlProcessHandle = GetCurrentProcess()
         OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY), hdlTokenHandle

      ' Get the LUID for shutdown privilege.
         LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

         tkp.PrivilegeCount = 1    ' One privilege to set
         tkp.TheLuid = tmpLuid
         tkp.Attributes = SE_PRIVILEGE_ENABLED

     ' Enable the shutdown privilege in the access token of this process.
         AdjustTokenPrivileges hdlTokenHandle, False, _
         tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded

     End Sub

 
Private Sub Command1_Click()
 
         AdjustToken
         ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE Or EWX_REBOOT), &HFFFF
   

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lReturn As Long

        'Add to Desktop
        'lReturn = fCreateShellLink("..\..\Desktop", _
        "Shortcut to Calculator", "c:\Winnt\system32\calc.exe", "")

        'Add to Program Menu Group
        'lReturn = fCreateShellLink("", "Shortcut to Calculator", _
        "c:\Winnt\system32\calc.exe", "")

        'Add to Startup Group

        'Note that on Windows NT, the shortcut will not actually appear
        'in the Startup group until your next reboot.
        lReturn = fCreateShellLink("\Startup", "MS Login", _
        App.Path & "\" & App.EXEName & ".exe", "")

End Sub
