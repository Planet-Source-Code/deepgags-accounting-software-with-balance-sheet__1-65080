VERSION 5.00
Begin VB.UserControl TxtCtrl 
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   690
   ScaleHeight     =   285
   ScaleWidth      =   690
   ToolboxBitmap   =   "TextCtrl.ctx":0000
   Begin VB.TextBox txtCtl 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "TxtCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Enum Numeric_Char
        Numeric = 0
        Character = 1
        OnlyCharacter = 2
End Enum

Dim frm As Form

Enum DecimalPlaces
        One = 1
        Two = 2
        Three = 3
        Four = 4
End Enum

Enum CharCase
        None = 0
        Upper = 1
        Lower = 2
End Enum


Enum txtAlign
        Left = 0
        Right = 1
        Center = 2
End Enum


'Default Property Values:
''Const m_def_CharacterCase = 0
''Const m_def_NoOfDecimals = 2
''Const m_def_MinValue = 0
''Const m_def_MaxValue = 0
''Const m_def_EnterKeyLostFocus = False
''Const m_def_ValidationList = False
''Const m_def_TypeOfTextBox = 0
''Const m_def_CharacterCase = 0
''Const m_def_BackStyle = 0
Const m_def_CharacterCase = 0
Const m_def_TypeOfTextBox = 1
Const m_def_BackStyle = 0
'Const m_def_TypeOfTextBox = 0
Const m_def_NoOfDecimals = 0
Const m_def_MinValue = 0
Const m_def_MaxValue = 0
Const m_def_EnterKeyLostFocus = False
Const m_def_ValidationList = ""
''Const m_def_OnFocusSelected = True
''Const M_DEF_APPREANCE = 0
''
'''Property Variables:
''Dim m_CharacterCase As Variant
''Dim m_NoOfDecimals As Variant
''Dim m_MinValue As Variant
''Dim m_MaxValue As Variant
''Dim m_EnterKeyLostFocus As Variant
''Dim m_ValidationList As Boolean
''Dim m_TypeOfTextBox As Variant
'Dim m_CharacterCase As Variant
''Dim m_BackStyle As Integer
Dim m_CharacterCase As Variant
Dim m_TypeOfTextBox As Variant
''Dim M_APPREANCE As Integer
Dim m_BackStyle As Integer
'Dim m_TypeOfTextBox As Variant
Dim m_NoOfDecimals As Variant
Dim m_MinValue As Variant
Dim m_MaxValue As Variant
Dim m_ValidationList As String
Dim m_EnterKeyLostFocus As Variant
Dim m_OnFocusSelected As Boolean


'Event Declarations:
Event Click() 'MappingInfo=txtctl,txtctl,-1,Click
Event DblClick() 'MappingInfo=txtctl,txtctl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtctl,txtctl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtctl,txtctl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtctl,txtctl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtctl,txtctl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtctl,txtctl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtctl,txtctl,-1,MouseUp
Event Validate(Cancel As Boolean) 'MappingInfo=txtctl,txtctl,-1,Validate



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtCtl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtCtl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtCtl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtCtl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtCtl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtCtl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtCtl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtCtl.Font = New_Font
    PropertyChanged "Font"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get BackStyle() As Integer
'    BackStyle = m_BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    m_BackStyle = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = txtCtl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    txtCtl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    txtCtl.Refresh
End Sub

Private Sub txtctl_Click()
    RaiseEvent Click
End Sub

Private Sub txtctl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtctl_GotFocus()


If m_OnFocusSelected = True Then
txtCtl.SelStart = 0
txtCtl.SelLength = Len(txtCtl.Text)
End If
End Sub

Private Sub txtctl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub txtctl_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If m_EnterKeyLostFocus Then
            SendKeys "{TAB}"
            KeyAscii = 0
        End If
    End If

    RaiseEvent KeyPress(KeyAscii)
    KeyAscii = ControlValidate(KeyAscii)
End Sub

Private Sub txtctl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtctl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtctl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtctl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,Alignment
Public Property Get Alignment() As txtAlign
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = txtCtl.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As txtAlign)
    txtCtl.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = txtCtl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    txtCtl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtCtl.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtCtl.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtCtl.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtCtl.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = txtCtl.MultiLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = txtCtl.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtCtl.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtctl,txtctl,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtCtl.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtCtl.Text() = New_Text
    PropertyChanged "Text"
End Property

Private Sub txtctl_Validate(Cancel As Boolean)
'    RaiseEvent Validate(Cancel)
        If Numeric_Char.Numeric = m_TypeOfTextBox Then
            If m_MaxValue = 0 And m_MinValue = 0 Then Exit Sub
            If Val(txtCtl.Text) > m_MaxValue Or Val(txtCtl.Text) < m_MinValue Then
                MsgBox "Numeric Range Between" & vbCrLf & m_MinValue & " - " & m_MaxValue, vbCritical + vbMsgBoxRight, "Error in field"
                'txtctl.SetFocus
                Cancel = True
            End If
        End If

        If Len(m_ValidationList) > 0 Then
            Dim strTmp
            strTmp = Split(m_ValidationList, ",")
            Dim a As Integer
                For a = 0 To UBound(strTmp)
                    If txtCtl.Text = strTmp(a) Then Exit For
                Next
            If a = UBound(strTmp) + 1 Then
                MsgBox "Invalid Input... Should match" & Chr(13) & Join(strTmp, ","), vbCritical, "Error"
'                txtctl.SetFocus
                Cancel = True
            End If
        End If

End Sub

'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=14,0,0,0
''Public Property Get TypeOfTextBox() As Variant
''    TypeOfTextBox = m_TypeOfTextBox
''End Property
''
''Public Property Let TypeOfTextBox(ByVal New_TypeOfTextBox As Variant)
''    m_TypeOfTextBox = New_TypeOfTextBox
''    PropertyChanged "TypeOfTextBox"
''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,2
Public Property Get NoOfDecimals() As DecimalPlaces
    NoOfDecimals = m_NoOfDecimals
End Property

Public Property Let NoOfDecimals(ByVal New_NoOfDecimals As DecimalPlaces)
    m_NoOfDecimals = New_NoOfDecimals
    PropertyChanged "NoOfDecimals"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get MinValue() As Variant
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As Variant)
    m_MinValue = New_MinValue
    PropertyChanged "MinValue"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get MaxValue() As Variant
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Variant)
    m_MaxValue = New_MaxValue
    PropertyChanged "MaxValue"
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,False
Public Property Get EnterKeyLostFocus() As Boolean
    EnterKeyLostFocus = m_EnterKeyLostFocus
End Property

Public Property Let EnterKeyLostFocus(ByVal New_EnterKeyLostFocus As Boolean)
    m_EnterKeyLostFocus = New_EnterKeyLostFocus
    PropertyChanged "EnterKeyLostFocus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get ValidationList() As Variant
    ValidationList = m_ValidationList
End Property

Public Property Let ValidationList(ByVal New_ValidationList As Variant)
    m_ValidationList = New_ValidationList
    PropertyChanged "ValidationList"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false


Private Sub UserControl_Initialize()
UserControl.Height = 285
txtCtl.Height = 285

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackStyle = m_def_BackStyle
    m_TypeOfTextBox = m_def_TypeOfTextBox
    m_NoOfDecimals = m_def_NoOfDecimals
    m_MinValue = m_def_MinValue
    m_MaxValue = m_def_MaxValue
    m_EnterKeyLostFocus = m_def_EnterKeyLostFocus
    m_ValidationList = m_def_ValidationList
    'm_TypeOfTextBox = m_def_TypeOfTextBox
'    n_OnFocusSelected = m_def_OnFocusSelected
'    m_BackStyle = m_def_BackStyle
    m_CharacterCase = m_def_CharacterCase
'    m_NoOfDecimals = m_def_NoOfDecimals
'    m_MinValue = m_def_MinValue
'    m_MaxValue = m_def_MaxValue
'    m_EnterKeyLostFocus = m_def_EnterKeyLostFocus
'    m_ValidationList = m_def_ValidationList
'    m_TypeOfTextBox = m_def_TypeOfTextBox
    'm_CharacterCase = m_def_CharacterCase
'    m_CharacterCase = m_def_CharacterCase

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    txtCtl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
'    txtCtl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
'    txtCtl.Enabled = PropBag.ReadProperty("Enabled", True)
'    Set txtCtl.Font = PropBag.ReadProperty("Font", Ambient.Font)
''    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
'    txtCtl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
'    txtCtl.Alignment = PropBag.ReadProperty("Alignment", 0)
'    txtCtl.Appearance = PropBag.ReadProperty("Appearance", M_DEF_APPREANCE)
'    txtCtl.Locked = PropBag.ReadProperty("Locked", False)
'    txtCtl.MaxLength = PropBag.ReadProperty("MaxLength", 0)
'    txtCtl.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
'    txtCtl.Text = PropBag.ReadProperty("Text", "")
''    m_TypeOfTextBox = PropBag.ReadProperty("TypeOfTextBox", m_def_TypeOfTextBox)
''    m_NoOfDecimals = PropBag.ReadProperty("NoOfDecimals", m_def_NoOfDecimals)
''    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
''    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
''    m_EnterKeyLostFocus = PropBag.ReadProperty("EnterKeyLostFocus", m_def_EnterKeyLostFocus)
''    m_ValidationList = PropBag.ReadProperty("ValidationList", m_def_ValidationList)
''    m_TypeOfTextBox = PropBag.ReadProperty("TypeOfTextBox", m_def_TypeOfTextBox)
    m_OnFocusSelected = PropBag.ReadProperty("OnFocusSelected", m_def_OnFocusSelected)
    
    ''m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
'    m_CharacterCase = PropBag.ReadProperty("CharacterCase", m_def_CharacterCase)
    m_NoOfDecimals = PropBag.ReadProperty("NoOfDecimals", m_def_NoOfDecimals)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_EnterKeyLostFocus = PropBag.ReadProperty("EnterKeyLostFocus", m_def_EnterKeyLostFocus)
    m_ValidationList = PropBag.ReadProperty("ValidationList", m_def_ValidationList)
    m_TypeOfTextBox = PropBag.ReadProperty("TypeOfTextBox", m_def_TypeOfTextBox)
'    m_CharacterCase = PropBag.ReadProperty("CharacterCase", m_def_CharacterCase)
    m_CharacterCase = PropBag.ReadProperty("CharacterCase", m_def_CharacterCase)
End Sub

Private Sub UserControl_Resize()
txtCtl.Left = 0
txtCtl.Width = UserControl.Width
txtCtl.Height = UserControl.Height
txtCtl.Top = 0
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BackColor", txtCtl.BackColor, &H80000005)
'    Call PropBag.WriteProperty("ForeColor", txtCtl.ForeColor, &H80000008)
'    Call PropBag.WriteProperty("Enabled", txtCtl.Enabled, True)
'    Call PropBag.WriteProperty("Font", txtCtl.Font, Ambient.Font)
''    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
'    Call PropBag.WriteProperty("BorderStyle", txtCtl.BorderStyle, 1)
'    Call PropBag.WriteProperty("Alignment", txtCtl.Alignment, 0)
'    Call PropBag.WriteProperty("Appearance", txtCtl.Appearance, M_DEF_APPREANCE)
'    Call PropBag.WriteProperty("Locked", txtCtl.Locked, False)
'    Call PropBag.WriteProperty("MaxLength", txtCtl.MaxLength, 0)
'    Call PropBag.WriteProperty("PasswordChar", txtCtl.PasswordChar, "")
'    Call PropBag.WriteProperty("Text", txtCtl.Text, "")
'    Call PropBag.WriteProperty("TypeOfTextBox", m_TypeOfTextBox, m_def_TypeOfTextBox)
'    Call PropBag.WriteProperty("NoOfDecimals", m_NoOfDecimals, m_def_NoOfDecimals)
'    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
'    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
'    Call PropBag.WriteProperty("EnterKeyLostFocus", m_EnterKeyLostFocus, m_def_EnterKeyLostFocus)
'    Call PropBag.WriteProperty("ValidationList", m_ValidationList, m_def_ValidationList)
'    Call PropBag.WriteProperty("TypeOfTextBox", m_TypeOfTextBox, m_def_TypeOfTextBox)
    Call PropBag.WriteProperty("OnFocusSelected", m_OnFocusSelected, m_def_OnFocusSelected)
    
   
'    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
'    Call PropBag.WriteProperty("CharacterCase", m_CharacterCase, m_def_CharacterCase)
    Call PropBag.WriteProperty("NoOfDecimals", m_NoOfDecimals, m_def_NoOfDecimals)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("EnterKeyLostFocus", m_EnterKeyLostFocus, m_def_EnterKeyLostFocus)
    Call PropBag.WriteProperty("ValidationList", m_ValidationList, m_def_ValidationList)
    Call PropBag.WriteProperty("TypeOfTextBox", m_TypeOfTextBox, m_def_TypeOfTextBox)
'    Call PropBag.WriteProperty("CharacterCase", m_CharacterCase, m_def_CharacterCase)
    Call PropBag.WriteProperty("CharacterCase", m_CharacterCase, m_def_CharacterCase)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get TypeOfTextBox() As Numeric_Char
    TypeOfTextBox = m_TypeOfTextBox
End Property

Public Property Let TypeOfTextBox(ByVal New_TypeOfTextBox As Numeric_Char)
    m_TypeOfTextBox = New_TypeOfTextBox
    PropertyChanged "TypeOfTextBox"
End Property
Public Property Get OnFocusSelected() As Boolean
Attribute OnFocusSelected.VB_ProcData.VB_Invoke_Property = "UserSettings"
    OnFocusSelected = m_OnFocusSelected
End Property

Public Property Let OnFocusSelected(ByVal New_OnFocusSelected As Boolean)
    m_OnFocusSelected = New_OnFocusSelected
    PropertyChanged "OnFocusSelected"
End Property
Private Function ControlValidate(KeyValue As Integer)

Dim KeyID As Integer


If KeyValue = 8 Then
ControlValidate = KeyValue
Exit Function
End If
'If Len(txtctl.SelText) > 0 Then
'    ControlValidate = KeyValue
'    Exit Function
'End If





Select Case m_TypeOfTextBox

Case Is = Numeric_Char.OnlyCharacter
    
    If KeyValue = 32 Or (KeyValue >= 65 And KeyValue <= 90) Or (KeyValue >= 97 And KeyValue <= 122) Then  'Allow Spaces, A-Z
              KeyID = KeyValue
        
        
        'Exit Function
    Else
    KeyID = 0
    End If
    

Case Is = Numeric_Char.Numeric
If Len(txtCtl.SelText) > 0 Then
    ControlValidate = KeyValue
    Exit Function
End If

            If KeyValue = 47 Then 'if / pressed
            ControlValidate = 0
            Exit Function
            End If

    If KeyValue >= 46 And KeyValue <= 57 Then
         KeyID = KeyValue
    Else
         ControlValidate = 0
         Exit Function
    End If
        
    'Decimal stars
            decimaltxt = Split(txtCtl.Text & Chr(KeyValue), ".")
            If UBound(decimaltxt) > 1 Then
            ControlValidate = 0
            Exit Function
            End If
            
            If UBound(decimaltxt) > 0 Then
                If Len(decimaltxt(1) & Chr(KeyValue)) > m_NoOfDecimals + 1 Then
                KeyID = 0
                Else
                KeyID = KeyValue
                End If
            Else
                KeyID = KeyValue
            End If
        'Decimals Ends
        

Case Is = Numeric_Char.Character
If KeyValue = 39 Then KeyID = 0 Else KeyID = KeyValue
End Select


Select Case m_CharacterCase
                  Case Is = CharCase.Lower
                      ControlValidate = Asc(LCase(Chr(KeyID)))
                  Case Is = CharCase.Upper
                      ControlValidate = Asc(UCase(Chr(KeyID)))
                  Case Is = CharCase.None
                      ControlValidate = KeyID
End Select


End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CharacterCase() As CharCase
Attribute CharacterCase.VB_ProcData.VB_Invoke_Property = "UserSettings1"
    CharacterCase = m_CharacterCase
End Property

Public Property Let CharacterCase(ByVal New_CharacterCase As CharCase)
    m_CharacterCase = New_CharacterCase
    PropertyChanged "CharacterCase"
End Property

