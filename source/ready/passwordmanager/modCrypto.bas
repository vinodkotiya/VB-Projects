Attribute VB_Name = "modCrypto"
Option Explicit
Public KeyCode As Byte

Public Function Encrypt(inText As String, vinKey As Byte) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''
'This function is written by vinod kotiya
' 21-aug - 2003 9:15 PM
Dim i As Integer
Dim l As Long  'length
Dim outText As String
Dim c As String * 1
l = Len(inText)

For i = 1 To l
 c = Mid$(inText, i, 1)
 c = Chr(AscB(c) Xor vinKey)
 outText = outText & c
Next

Encrypt = outText
End Function

Public Function Decrypt(inText As String, vinKey As Byte) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''
'This function is written by vinod kotiya
' 21-aug - 2003 9:15 PM
Dim i As Integer
Dim l As Long  'length
Dim outText As String
Dim c As String * 1
l = Len(inText)

For i = 1 To l
 c = Mid$(inText, i, 1)
 c = Chr(AscB(c) Xor vinKey)
 outText = outText & c
Next

Decrypt = outText
End Function

