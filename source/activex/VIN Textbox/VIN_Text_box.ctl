VERSION 5.00
Begin VB.UserControl VINText 
   BackStyle       =   0  'Transparent
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   ScaleHeight     =   795
   ScaleWidth      =   3105
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   360
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "VIN "
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "VINText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim switch As Boolean
Dim color1 As OLE_COLOR
Dim color2 As OLE_COLOR

Private Sub Timer1_Timer()
If switch = True Then
     Text1.BackColor = color1
     switch = False
 Else
    Text1.BackColor = color2
     switch = True
End If
End Sub

Private Sub UserControl_Initialize()
switch = True
Timer1.Interval = 0
color1 = &HFFFFC0
color2 = &HFFC0FF
End Sub


Public Property Get BlinkInterval() As Long
Attribute BlinkInterval.VB_Description = "Set the blinking time in mSeconds"
BlinkInterval = Timer1.Interval
End Property

Public Property Let BlinkInterval(ByVal Newinterval As Long)
Timer1.Interval = Newinterval
PropertyChanged BlinkInterval
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Text1.Text = PropBag.ReadProperty("VINText", "My Text")
End Sub

Public Property Get BlinkColor1() As OLE_COLOR
Attribute BlinkColor1.VB_Description = "Set the blink color"
BlinkColor1 = color1
End Property

Public Property Let BlinkColor1(ByVal vNewValue As OLE_COLOR)
 color1 = vNewValue
 PropertyChanged BlinkColor1
End Property
Public Property Get BlinkColor2() As OLE_COLOR
Attribute BlinkColor2.VB_Description = "Set the blink color"
BlinkColor2 = color2
End Property

Public Property Let BlinkColor2(ByVal vNewValue As OLE_COLOR)
 color2 = vNewValue
 PropertyChanged BlinkColor2
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("VINText", Text1.Text, "My Text")
End Sub
