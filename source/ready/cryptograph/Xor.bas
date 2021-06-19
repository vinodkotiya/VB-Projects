Attribute VB_Name = "Module1"
Option Explicit
Public starttime As Date
Public isAbort As Boolean

Public Function XORED(infile, outfile, key)

Dim FileLength As Long, v As Long
Dim lk As Integer, i As Long

lk = Len(key)

'On Error Resume Next

getkey key

Dim c As String * 1


FileLength = FileLen(infile)
Dim fnumIn As Integer  ''file handle
Dim fnumOut As Integer
Dim mask As Variant
fnumIn = FreeFile

Open infile For Binary As #fnumIn
If Err Then MsgBox (Error(Err))
fnumOut = FreeFile
Open outfile For Binary As #fnumOut
If Err Then MsgBox (Error(Err))

mask = Int(Rnd * 256)


'main loop
For i = 1 To FileLength
''' abort
 If isAbort Then
   Close fnumIn
   If Err Then MsgBox (Error(Err))
   Close fnumOut
   If Err Then MsgBox (Error(Err))
   Exit Function
 End If
 ''''''''''''''
    Get fnumIn, , c
    c = Chr(Asc(c) Xor mask)
    Put fnumOut, , c

    If i Mod 100 = 0 Then
     v = (i / FileLength) * 100
     
     '''  foreign code '''''''''''''''''''
     frmCrypto.lnTop(0).X2 = frmCrypto.lnTop(0).X1 + Round(5760 * v / 100)      'progress bar
     frmCrypto.lnTop(1).X2 = frmCrypto.lnTop(1).X1 + Round(5760 * v / 100)         'progress bar
     frmCrypto.lblTime(0).Caption = DateDiff("s", starttime, Now) & " Sec"
     frmCrypto.lblTime(3).Caption = Round(i / 1024) & " KB"
     ''''''''''''''''''''''''''''''''''''''''
     
    End If
    
   
'rotate password

key = Right(key, 1) & Left(key, lk - 1)

' get new leftmost character ANSI value
Dim X, j, t As Long
X = Asc(Left(key, 1))

'throw away random numbers up to the value of the character
    
    For j = 1 To X
        t = Rnd
        DoEvents
    Next j


mask = Int(Rnd * 256)

DoEvents
Next i



Close fnumIn
If Err Then MsgBox (Error(Err))
Close fnumOut
If Err Then MsgBox (Error(Err))
'Dim cr
'cr = Chr(13) & Chr(10)

End Function

Public Sub getkey(key)

''Form1.Hide
''Form3.Show
DoEvents


Dim X, i, j As Long
Dim z As Single
Dim t As Single
Dim n As Long

For i = 1 To Len(key)
    X = Asc(Mid(key, i, 1))
''Form3.Caption = X 'display
    n = i * X
    For j = 1 To n
        t = Rnd
        DoEvents
    Next j
    DoEvents
Next i

End Sub


