VERSION 5.00
Begin VB.UserControl mycontrol 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "mycontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Command1_Click()
Text1.Text = "Vinod kotiya"
Call Me.MyMethod

End Sub

Public Property Get myproperty() As String
myproperty = Text1.Text
End Property

Public Property Let myproperty(ByVal vNewValue As String)
Text1.Text = vNewValue
PropertyChanged " MyProperty"

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Text1.Text = PropBag.ReadProperty("MyProperty", "Enter Text")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("MyProperty", Text1.Text, "Enter Text")

End Sub

Public Sub MyMethod()
MsgBox "my control is in action"
End Sub
