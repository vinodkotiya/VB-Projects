VERSION 5.00
Begin VB.UserControl VIN_Label 
   BackColor       =   &H00FFFFC0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ScaleHeight     =   270
   ScaleWidth      =   2055
End
Attribute VB_Name = "VIN_Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_Caption = "VIN Label"
Const m_def_Alignment = 0
Const m_def_Effect3D = 6
'Const m_def_TextAlignment = 4
'Const m_def_Caption = "VIN Label"
'Const m_def_Effect = 2
'Property Variables:
Dim m_Caption As String
Dim m_Alignment As Integer
Dim m_Effect3D As Variant
Dim shadowColor As OLE_COLOR
'Dim m_TextAlignment As Integer
'Dim m_Caption As Variant
'Dim m_Effect As Integer
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."

Enum Align
 [Left]
 [Center]
 [Right]
End Enum
Enum Effects
    None
    [Carved Light]
    Carved
    [Carved Heavy]
    [Raised Light]
    Raised
    [Raised Heavy]
End Enum







'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    If New_BackStyle = 0 Then Drawcaption
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PopupMenu
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
    UserControl.PopupMenu Menu, Flags, X, Y, DefaultMenu
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Appearance
'Public Property Get Appearance() As Integer
'    Appearance = UserControl.Appearance
'End Property
'
'Public Property Let Appearance(ByVal New_Appearance As Integer)
'    UserControl.Appearance = New_Appearance
'    PropertyChanged "Appearance"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=7,0,0,4
''Public Property Get TextAlignment() As Integer
''    TextAlignment = m_TextAlignment
''End Property
''
''Public Property Let TextAlignment(ByVal New_TextAlignment As Integer)
''    m_TextAlignment = New_TextAlignment
''    Drawcaption
''    PropertyChanged "TextAlignment"
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=14,0,0,VIN Label
''Public Property Get Caption() As Variant
''    Caption = m_Caption
''End Property
''
''Public Property Let Caption(ByVal New_Caption As Variant)
''    m_Caption = New_Caption
''    Drawcaption
''    PropertyChanged "Caption"
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=7,0,0,2
''Public Property Get Effect() As Integer
''    Effect = m_Effect
''End Property
''
''Public Property Let Effect(ByVal New_Effect As Integer)
''    m_Effect = New_Effect
''    Drawcaption
''    PropertyChanged "Effect"
''End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
'    m_TextAlignment = m_def_TextAlignment
'    m_Caption = m_def_Caption
'    m_Effect = m_def_Effect
    UserControl.BorderStyle = 1
    UserControl.BackStyle = 1
    m_Caption = m_def_Caption
    m_Alignment = m_def_Alignment
    m_Effect3D = m_def_Effect3D
End Sub

Private Sub UserControl_Paint()
Drawcaption
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
'    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
'    m_TextAlignment = PropBag.ReadProperty("TextAlignment", m_def_TextAlignment)
'    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
'    m_Effect = PropBag.ReadProperty("Effect", m_def_Effect)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_Effect3D = PropBag.ReadProperty("Effect3D", m_def_Effect3D)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
'    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
'    Call PropBag.WriteProperty("TextAlignment", m_TextAlignment, m_def_TextAlignment)
'    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
'    Call PropBag.WriteProperty("Effect", m_Effect, m_def_Effect)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("Effect3D", m_Effect3D, m_def_Effect3D)
End Sub

Private Sub Drawcaption()
Dim CaptionWidth As Long, CaptionHeight As Long
Dim CurrX As Long, CurrY As Long
Dim oldForeColor As OLE_COLOR

    CaptionHeight = TextHeight(m_Caption)
    CaptionWidth = TextWidth(m_Caption)
    Select Case m_Alignment
        Case 0:   'left
            CurrX = 30
            CurrY = (UserControl.Height - CaptionHeight) / 2
        Case 1:    'middle
            CurrX = (UserControl.Width - CaptionWidth) / 2
            CurrY = (UserControl.Height - CaptionHeight) / 2
        Case 2:
            CurrX = UserControl.Width - CaptionWidth - 30
            CurrY = (UserControl.Height - CaptionHeight) / 2
    End Select

oldForeColor = UserControl.ForeColor
    Select Case m_Effect3D
        Case 0:
            UserControl.Cls
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.Print m_Caption
        Case 1:
            UserControl.Cls
            UserControl.CurrentX = CurrX + 15
            UserControl.CurrentY = CurrY + 15
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 2:
            UserControl.Cls
            UserControl.CurrentX = CurrX + 30
            UserControl.CurrentY = CurrY + 30
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 3:
            UserControl.Cls
            UserControl.CurrentX = CurrX + 45
            UserControl.CurrentY = CurrY + 45
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX + 30
            UserControl.CurrentY = CurrY + 30
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX + 15
            UserControl.CurrentY = CurrY + 15
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 4:
            UserControl.Cls
            UserControl.CurrentX = CurrX - 15
            UserControl.CurrentY = CurrY - 15
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 5:
            UserControl.Cls
            UserControl.CurrentX = CurrX - 30
            UserControl.CurrentY = CurrY - 30
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        Case 6:
            UserControl.Cls
            UserControl.CurrentX = CurrX - 45
            UserControl.CurrentY = CurrY - 45
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX - 30
            UserControl.CurrentY = CurrY - 30
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            
            UserControl.CurrentX = CurrX - 15
            UserControl.CurrentY = CurrY - 15
            UserControl.ForeColor = shadowColor
            UserControl.Print m_Caption
            UserControl.CurrentX = CurrX
            UserControl.CurrentY = CurrY
            UserControl.ForeColor = oldForeColor
            UserControl.Print m_Caption
        
        End Select
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,VIN Label
Public Property Get Caption() As String
Attribute Caption.VB_Description = "show the text of label to ve displayed."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    Drawcaption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Alignment() As Align
Attribute Alignment.VB_Description = "Determine how the caption will be alligned on the control"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Align)
    m_Alignment = New_Alignment
    Drawcaption
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,6
Public Property Get Effect3D() As Effects
Attribute Effect3D.VB_Description = "determine the type of 3D effects you wanna"
    Effect3D = m_Effect3D
End Property

Public Property Let Effect3D(ByVal New_Effect3D As Effects)
    m_Effect3D = New_Effect3D
    Drawcaption
    PropertyChanged "Effect3D"
End Property

Public Property Get Colorshadow() As OLE_COLOR
    Colorshadow() = shadowColor
End Property

Public Property Let Colorshadow(ByVal New_ColorShadow As OLE_COLOR)
   shadowColor = New_ColorShadow
    PropertyChanged "ColorShadow"
End Property

