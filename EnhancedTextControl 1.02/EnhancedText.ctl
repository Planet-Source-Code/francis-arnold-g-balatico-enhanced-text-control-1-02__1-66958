VERSION 5.00
Begin VB.UserControl EnhancedText 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ClipControls    =   0   'False
   DataBindingBehavior=   1  'vbSimpleBound
   ScaleHeight     =   1110
   ScaleWidth      =   4500
   ToolboxBitmap   =   "EnhancedText.ctx":0000
   Begin VB.TextBox txtText2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Top             =   420
      Width           =   4455
   End
   Begin VB.Shape shpShape 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   15
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "EnhancedText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'SPECIFICATIONS (Enhanced Property) :
'====================================
'
' InputType :
'------------
'  0 - inpNone
'  1 - inpAlphabetic
'  2 - inpNumber
'  3 - inpAlphaNumeric
'
' CharCase : [if Alphabetic + AlphaNumeric]:
'---------------------------------------
'  0 - cseNone
'  1 - cseUpper
'  2 - cseLower
'  3 - cseProper
'
' Alignment :
'------------
'  0 - Left Justify
'  1-  right justify
'  2 - Center justify
'
' Text :
'-------
' Any value
'
' OnFocusSelect :
'----------------
'   - True
'   - False
'
' OnFocusBgColor :
'-----------------
'   - Color selection dialog
'
' OnFocusFontColor :
'-------------------
'   - Color selection dialog
'
' OnFocusFont :
'-------------
'   - FontSelectionDialog
'
' EnterExitKey :
'---------------
'   - True
'   - False
'
'*********************************************************************************
Option Explicit

'Proper character flag
Dim blnSpaceFlag As Boolean

'Property Variable
Dim intCharAccept As Long
Dim intCase As Long
Dim intNormalBorderPattern As Long
Dim intFocusBorderPattern As Long
Dim intDisabledBorderPattern As Long

Dim blnFocusSelect As Boolean
Dim blnExitkey As Boolean
Dim blnAutoTab As Boolean

Dim oleNormalBackColor As OLE_COLOR
Dim oleNormalGrooveBackColor As OLE_COLOR
Dim oleFocusGrooveBackColor As OLE_COLOR
Dim oleFocusBackColor As OLE_COLOR
Dim oleNormalFontColor As OLE_COLOR
Dim oleFocusFontColor As OLE_COLOR
Dim oleNormalBorderColor As OLE_COLOR
Dim oleFocusBorderColor As OLE_COLOR

Dim oleDisabledGrooveBackColor As OLE_COLOR
Dim oleDisabledColor As OLE_COLOR
Dim oleDisabledBorderColor As OLE_COLOR

Dim fntNormal As StdFont
Dim fntFocus As StdFont

Dim strFormatString As String
Dim strSpecialChar As String

Dim m_Multi As Boolean

Enum enmCharAccept
    None
    Alphabetic
    Numeric
    AlphaNumeric
    Customize
End Enum
Enum BorderPattern
    Transparent
    Solid
    Dash
    Dot
    DashDot
    DashDotDot
    InsideSolid
End Enum
Enum enmCase
    None
    UpperCase
    LowerCase
    ProperCase
End Enum
Enum enmAlignment
    Left_Align
    Right_Align
    Center_Align
End Enum

'Events declaration
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event Click()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)

'*******************************************************************************
'   Control commmon Events
'*******************************************************************************

Private Sub txtText_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub txtText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub txtText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub txtText_Change()

    RaiseEvent Change
    If (blnAutoTab = True) Then
        If (CDbl(Len(txtText.Text)) >= CDbl(txtText.MaxLength)) Then 'autotab handling
            SendKeys "{Tab}"
            Exit Sub
        End If
    End If

End Sub

Private Sub txtText_Click()

    RaiseEvent Click

End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)

    On Error Resume Next
        RaiseEvent KeyPress(KeyAscii)

        If KeyAscii = 8 Then Exit Sub

        If (Len(txtText.Text) = 0) Then blnSpaceFlag = True

        If (blnExitkey = True) Then         'ENTER as TAB key
            If KeyAscii = 13 Then
                SendKeys "{Tab}"
                Exit Sub
            End If
        End If

        If (InStr(strSpecialChar, Chr(KeyAscii)) <> 0) Then 'special char handling
            Exit Sub
        End If

        Select Case intCharAccept
          Case 0          '  0 - inpNone

            KeyAscii = ModifyCase(KeyAscii)
            If KeyAscii = 32 Then blnSpaceFlag = True
            Exit Sub

          Case 1          '  1 - inpAlphabetic

            If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                KeyAscii = ModifyCase(KeyAscii)
              ElseIf (KeyAscii = 32) Then
                blnSpaceFlag = True
              Else
                KeyAscii = 0
                Beep
            End If
            Exit Sub

          Case 2          '  2 - inpDecimalNumber

            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 46)) Or isDupDot(1, KeyAscii) = True Then
                KeyAscii = 0
                Beep
            End If
            Exit Sub

          Case 3          '  4 - inpAlphaNumeric

            If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                KeyAscii = ModifyCase(KeyAscii)
              ElseIf ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 46)) Then
                If blnSpaceFlag = True Then blnSpaceFlag = False
              ElseIf (KeyAscii = 32) Then
                blnSpaceFlag = True
              Else
                KeyAscii = 0
                Beep
            End If
            Exit Sub
          Case 4
            If (InStr(strSpecialChar, Chr(KeyAscii)) = 0) Then 'special char handling
                KeyAscii = 0
            End If
            Exit Sub
        End Select

End Sub

Private Sub txtText_LostFocus()

    blnSpaceFlag = True
    UserControl.BackColor = oleNormalGrooveBackColor

    If txtText.Enabled = True Then
        txtText.BackColor = oleNormalBackColor
      Else
        txtText.BackColor = oleDisabledColor
    End If

    txtText.ForeColor = oleNormalFontColor
    shpShape.BorderColor = oleNormalBorderColor
    shpShape.BorderStyle = intNormalBorderPattern
    Set txtText.Font = fntNormal
    If (intCharAccept = 2) Then
        If Not IsNumeric(txtText.Text) Then txtText.Text = 0
    End If
    If Not IsNull(strFormatString) Then
        txtText.Text = Format$(txtText.Text, strFormatString)
    End If

End Sub

Private Sub txtText_GotFocus()

    If (blnFocusSelect = True) Then
        txtText.SelStart = 0
        txtText.SelLength = Len(txtText)
    End If

    UserControl.BackColor = oleFocusGrooveBackColor

    If txtText.Enabled = True Then
        txtText.BackColor = oleFocusBackColor
      Else
        txtText.BackColor = oleDisabledColor
    End If

    txtText.ForeColor = oleFocusFontColor
    shpShape.BorderColor = oleFocusBorderColor
    shpShape.BorderStyle = intFocusBorderPattern
    Set txtText.Font = fntFocus

End Sub

'*******************************************************************************
'   Control commmon procedure
'*******************************************************************************

Private Sub UserControl_Initialize()

    Set fntNormal = New StdFont
    Set fntFocus = New StdFont
    
    blnSpaceFlag = True

    intCharAccept = 0
    intCase = 0

    intNormalBorderPattern = 1
    intFocusBorderPattern = 1

    blnFocusSelect = True
    
    m_Multi = False
    
    oleNormalBackColor = &HFFFFFF
    oleFocusBackColor = &HC0FFFF
    oleNormalFontColor = &H0
    oleFocusFontColor = &HC00000
    oleNormalBorderColor = &H808080
    oleFocusBorderColor = &H80FF&
    oleNormalGrooveBackColor = &H8000000F
    oleFocusGrooveBackColor = &H8000000F
    oleDisabledColor = &H8000000F
    oleDisabledBorderColor = &H8000000F
    
    fntNormal.Name = "Arial"
    fntNormal.Size = 8

    fntFocus.Name = "Arial"
    fntFocus.Size = 8
    fntFocus.Bold = True

End Sub

Private Function ModifyCase(KeyAscii As Integer) As Integer

    Select Case intCase
      Case 0  'No case
        ModifyCase = KeyAscii
        Exit Function
      Case 1  'Upper
        If (KeyAscii >= 97 And KeyAscii <= 122) Then
            ModifyCase = KeyAscii - 32
          Else
            ModifyCase = KeyAscii
        End If
        Exit Function
      Case 2  'Lower
        If (KeyAscii >= 65 And KeyAscii <= 90) Then
            ModifyCase = KeyAscii + 32
          Else
            ModifyCase = KeyAscii
        End If
        Exit Function
      Case 3 'Proper
        If (blnSpaceFlag = True) Then
            If (KeyAscii >= 97 And KeyAscii <= 122) Then
                ModifyCase = KeyAscii - 32
              Else
                ModifyCase = KeyAscii
            End If
            blnSpaceFlag = False
          Else
            ModifyCase = KeyAscii
        End If
        Exit Function
    End Select

End Function

Private Sub UserControl_InitProperties()
    
    txtText2.Visible = False
    txtText.Visible = True
    
    txtText.BackColor = oleNormalBackColor
    shpShape.BorderColor = oleNormalBorderColor
    
End Sub

Private Sub UserControl_Resize()

  Dim uWidth As Long
  Dim uHeight As Long

    uWidth = UserControl.Width
    uHeight = UserControl.Height

    If (uWidth >= 190 And uHeight >= 270) Then
        shpShape.Left = 0
        shpShape.Top = 0
        shpShape.Width = uWidth
        shpShape.Height = uHeight
        
        txtText.Top = 30
        txtText.Left = 30
        txtText.Width = uWidth - 70
        txtText.Height = uHeight - 70
        
        txtText2.Top = 30
        txtText2.Left = 30
        txtText2.Width = uWidth - 70
        txtText2.Height = uHeight - 70
      Else
        If (uHeight < 270) Then
            UserControl.Height = 270
        End If
        If (uWidth < 190) Then
            UserControl.Width = 190
        End If
    End If

End Sub

'==================================
'   Properties : InputType
'==================================

Public Property Get InputType() As enmCharAccept
Attribute InputType.VB_Description = "Set or returns the input type for Enhanced Text box."

    InputType = intCharAccept

End Property

Public Property Let InputType(ByVal vNewValue As enmCharAccept)

    intCharAccept = vNewValue
    If (intCharAccept = 2) Then
        txtText.Text = 0
        txtText2.Text = 0
      Else
        txtText.Text = ""
        txtText2.Text = ""
    End If

    If Not IsNull(strFormatString) Then
        txtText.Text = Format$(txtText.Text, strFormatString)
        txtText2.Text = Format$(txtText.Text, strFormatString)
    End If
    PropertyChanged "InputType"

End Property

'==================================
'   Properties : CharCase
'==================================

Public Property Get CharCase() As enmCase
Attribute CharCase.VB_Description = "Specify the input Alphabet case."

    CharCase = intCase

End Property

Public Property Let CharCase(ByVal vNewValue As enmCase)

    intCase = vNewValue
    PropertyChanged "CharCase"

End Property

'==================================
'   Properties : Alignment
'==================================

Public Property Get Alignment() As enmAlignment
Attribute Alignment.VB_Description = "Returns or sets the Alignment of inputs."

    Alignment = txtText.Alignment

End Property

Public Property Let Alignment(ByVal vNewValue As enmAlignment)

    txtText.Alignment = vNewValue
    txtText2.Alignment = vNewValue
    PropertyChanged "Alignment"

End Property

'==================================
'   Properties : OnFocusSelect
'==================================

Public Property Get OnFocusSelect() As Boolean
Attribute OnFocusSelect.VB_Description = "Specify wheather to Select the input Text when control get focus."

    OnFocusSelect = blnFocusSelect

End Property

Public Property Let OnFocusSelect(ByVal vNewValue As Boolean)

    blnFocusSelect = vNewValue
    PropertyChanged "OnFocusSelect"

End Property

'==================================
'   Properties : NormalGrooveBackColor
'==================================

Public Property Get NormalGrooveBackColor() As OLE_COLOR

    NormalGrooveBackColor = oleNormalGrooveBackColor

End Property

Public Property Let NormalGrooveBackColor(ByVal vNewValue As OLE_COLOR)

    oleNormalGrooveBackColor = vNewValue
    UserControl.AutoRedraw = True
    
    If txtText.Enabled = True Or txtText2.Enabled = True Then
        UserControl.BackColor = vNewValue   'this is the back Color
    End If
    
    UserControl.AutoRedraw = False
    PropertyChanged "NormalGrooveBackColor"

End Property

'==================================
'   Properties : DisabledBackColor
'==================================

Public Property Get DisabledBackColor() As OLE_COLOR

    DisabledBackColor = oleDisabledColor

End Property

Public Property Let DisabledBackColor(ByVal vNewValue As OLE_COLOR)

    oleDisabledColor = vNewValue

    If txtText.Enabled = False Then txtText.BackColor = oleDisabledColor
    If txtText2.Enabled = False Then txtText2.BackColor = oleDisabledColor
    
    PropertyChanged "DisabledBackColor"

End Property

'==================================
'   Properties : DisabledBorderColor
'==================================

Public Property Get DisabledBorderColor() As OLE_COLOR

    DisabledBorderColor = oleDisabledBorderColor

End Property

Public Property Let DisabledBorderColor(ByVal vNewValue As OLE_COLOR)

    oleDisabledBorderColor = vNewValue
    
    If txtText.Enabled = False Or txtText2.Enabled = False Then
        shpShape.BorderColor = oleDisabledBorderColor
    End If
    
    PropertyChanged "DisabledBorderColor"

End Property

'==================================
'   Properties : DisabledBorderPattern
'==================================
Public Property Get DisabledBorderPattern() As BorderPattern

    DisabledBorderPattern = intDisabledBorderPattern

End Property

Public Property Let DisabledBorderPattern(ByVal vNewValue As BorderPattern)

    intDisabledBorderPattern = vNewValue
    
    If txtText.Enabled = True Or txtText2.Enabled = True Then
        shpShape.BorderStyle = intNormalBorderPattern
    Else
        shpShape.BorderStyle = intDisabledBorderPattern
    End If
    
    PropertyChanged "DisabledBorderPattern"

End Property


'==================================
'   Properties : DisabledGrooveBackColor
'==================================
Public Property Get DisabledGrooveBackColor() As OLE_COLOR

    DisabledGrooveBackColor = oleDisabledGrooveBackColor

End Property

Public Property Let DisabledGrooveBackColor(ByVal vNewValue As OLE_COLOR)

    oleDisabledGrooveBackColor = vNewValue
    
    UserControl.AutoRedraw = True
    
    If txtText.Enabled = False Or txtText2.Enabled = False Then
        UserControl.BackColor = oleDisabledGrooveBackColor
    End If
    
    UserControl.AutoRedraw = False
    
    PropertyChanged "DisabledGrooveBackColor"

End Property

'==================================
'   Properties : FocusGrooveBackColor
'==================================
Public Property Get FocusGrooveBackColor() As OLE_COLOR

    FocusGrooveBackColor = oleFocusGrooveBackColor

End Property

Public Property Let FocusGrooveBackColor(ByVal vNewValue As OLE_COLOR)

    oleFocusGrooveBackColor = vNewValue
    PropertyChanged "FocusGrooveBackColor"

End Property

'==================================
'   Properties : NormalBackColor
'==================================

Public Property Get NormalBackColor() As OLE_COLOR
Attribute NormalBackColor.VB_Description = "Sets or returns the BackGround Color when control is not focus."

    NormalBackColor = oleNormalBackColor

End Property

Public Property Let NormalBackColor(ByVal vNewValue As OLE_COLOR)

    oleNormalBackColor = vNewValue

    If txtText.Enabled = True Then txtText.BackColor = oleNormalBackColor   'this is the back Color
    If txtText2.Enabled = True Then txtText2.BackColor = oleNormalBackColor   'this is the back Color

    PropertyChanged "NormalBackColor"

End Property

'==================================
'   Properties : FocusBackColor
'==================================

Public Property Get FocusBackColor() As OLE_COLOR
Attribute FocusBackColor.VB_Description = "Sets or returns the BackGround Color when control is in focus."

    FocusBackColor = oleFocusBackColor

End Property

Public Property Let FocusBackColor(ByVal vNewValue As OLE_COLOR)

    oleFocusBackColor = vNewValue
    PropertyChanged "FocusBackColor"

End Property

'==================================
'   Properties : NormalFontColor
'==================================

Public Property Get NormalFontColor() As OLE_COLOR
Attribute NormalFontColor.VB_Description = "Sets or returns the FontColor when control is not focus."

    NormalFontColor = oleNormalFontColor

End Property

Public Property Let NormalFontColor(ByVal vNewValue As OLE_COLOR)

    oleNormalFontColor = vNewValue
    txtText.ForeColor = oleNormalFontColor
    PropertyChanged "NormalFontColor"

End Property

'==================================
'   Properties : FocusFontColor
'==================================

Public Property Get FocusFontColor() As OLE_COLOR
Attribute FocusFontColor.VB_Description = "Sets or returns the Font Color when control is in focus."

    FocusFontColor = oleFocusFontColor

End Property

Public Property Let FocusFontColor(ByVal vNewValue As OLE_COLOR)

    oleFocusFontColor = vNewValue
    PropertyChanged "FocusFontColor"

End Property

'==================================
'   Properties : NormalFont
'==================================

Public Property Get NormalFont() As Font
Attribute NormalFont.VB_Description = "Sets or returns the Font when control is not focus."

    Set NormalFont = fntNormal

End Property

Public Property Set NormalFont(ByRef vNewValue As Font) 'Make sure this is Pass ByReference and method is set

    Set fntNormal = vNewValue
    
    Set txtText.Font = vNewValue
    Set txtText2.Font = vNewValue
    
    PropertyChanged ("NormalFont")

End Property

'==================================
'   Properties : FocusFont
'==================================

Public Property Get FocusFont() As Font
Attribute FocusFont.VB_Description = "Sets or returns the Font when control is in focus."

    Set FocusFont = fntFocus

End Property

Public Property Set FocusFont(ByRef vNewValue As Font) 'Make sure this is Pass ByReference and method is set

    Set fntFocus = vNewValue
    PropertyChanged ("FocusFont")

End Property

'==================================
'   Properties : Text
'==================================

Public Property Get Text() As String
Attribute Text.VB_Description = "The Text Value that need to be displayed in control."
    
    If m_Multi = True Then
        Text = txtText2.Text
    Else
        Text = txtText.Text
    End If

End Property

Public Property Let Text(ByVal vNewValue As String)
    
    txtText.Text = vNewValue
    txtText2.Text = vNewValue
    
    PropertyChanged "Text"
End Property

'==================================
'   Properties : EnterExitKey
'==================================

Public Property Get EnterExitKey() As Boolean
Attribute EnterExitKey.VB_Description = "Specify whether to Enable the ENTER key for exit the Control."

    EnterExitKey = blnExitkey

End Property

Public Property Let EnterExitKey(ByVal vNewValue As Boolean)

    blnExitkey = vNewValue
    PropertyChanged "EnterExitKey"

End Property

'==================================
'   Properties : Enable
'==================================

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Set or return the control Enabled or not."

    Enabled = txtText.Enabled

End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)

    txtText.Enabled = vNewValue
    txtText2.Enabled = vNewValue
    
    UserControl.AutoRedraw = True
    
    If txtText.Enabled = False Or txtText2.Enabled = False Then
        txtText.BackColor = oleDisabledColor
        txtText2.BackColor = oleDisabledColor
        shpShape.BorderColor = oleDisabledBorderColor
        UserControl.BackColor = oleDisabledGrooveBackColor
        shpShape.BorderStyle = intDisabledBorderPattern
      ElseIf txtText.Enabled = True Or txtText2.Enabled = True Then
        txtText.BackColor = oleNormalBackColor
        txtText2.BackColor = oleNormalBackColor
        shpShape.BorderColor = oleNormalBorderColor
        UserControl.BackColor = oleNormalGrooveBackColor
        shpShape.BorderStyle = intNormalBorderPattern
    End If
    
    UserControl.AutoRedraw = False
    
    PropertyChanged "Enabled"

End Property

'==================================
'   Properties : Multiliner
'==================================
Public Property Get MultiLiner() As Boolean

    MultiLiner = m_Multi
    
End Property

Public Property Let MultiLiner(ByVal vNewValue As Boolean)

    m_Multi = vNewValue
    
    If m_Multi = True Then
        txtText2.Text = txtText.Text
        
        txtText.Visible = False
        txtText2.Visible = True
    Else
        txtText.Text = txtText2.Text
        
        txtText.Visible = True
        txtText2.Visible = False
    End If
    
    PropertyChanged "MultiLiner"

End Property

'==================================
'   Properties : Locked
'==================================
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Sets or returns wheather Control is locked or not."

    Locked = txtText.Locked

End Property

Public Property Let Locked(ByVal vNewValue As Boolean)

    txtText.Locked = vNewValue
    txtText2.Locked = txtText.Locked
    
    PropertyChanged "Locked"

End Property

'==================================
'   Properties : NormalBorderPattern
'==================================
Public Property Get NormalBorderPattern() As BorderPattern
Attribute NormalBorderPattern.VB_Description = "Specify the pattern for Fixed single outer Border."

    NormalBorderPattern = intNormalBorderPattern

End Property

Public Property Let NormalBorderPattern(ByVal vNewValue As BorderPattern)

    intNormalBorderPattern = vNewValue
    
    If txtText.Enabled = True Or txtText2.Enabled = True Then
        shpShape.BorderStyle = intNormalBorderPattern
    Else
        shpShape.BorderStyle = intDisabledBorderPattern
    End If
    
    PropertyChanged "NormalBorderPattern"

End Property

'==================================
'   Properties : FocusBorderPattern
'==================================
Public Property Get FocusBorderPattern() As BorderPattern
Attribute FocusBorderPattern.VB_Description = "Sets or returns the Border pattern when control is in focus."

    FocusBorderPattern = intFocusBorderPattern

End Property

Public Property Let FocusBorderPattern(ByVal vNewValue As BorderPattern)

    intFocusBorderPattern = vNewValue
    PropertyChanged "FocusBorderPattern"

End Property

'==================================
'   Properties : NormalBorderColor
'==================================
Public Property Get NormalBorderColor() As OLE_COLOR
Attribute NormalBorderColor.VB_Description = "Specify the Color For Fixed single outer Border."

    NormalBorderColor = oleNormalBorderColor

End Property

Public Property Let NormalBorderColor(ByVal vNewValue As OLE_COLOR)

    oleNormalBorderColor = vNewValue
    shpShape.BorderColor = vNewValue
    PropertyChanged ("NormalBorderColor")

End Property

'==================================
'   Properties : FocusBorderColor
'==================================

Public Property Get FocusBorderColor() As OLE_COLOR
Attribute FocusBorderColor.VB_Description = "Sets or returns the Border Color when control is in focus."

    FocusBorderColor = oleFocusBorderColor

End Property

Public Property Let FocusBorderColor(ByVal vNewValue As OLE_COLOR)

    oleFocusBorderColor = vNewValue
    PropertyChanged "FocusBorderColor"

End Property

'==================================
'   Properties : PasswordChar
'==================================
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Specify The password Character."

    PasswordChar = txtText.PasswordChar

End Property

Public Property Let PasswordChar(ByVal vNewValue As String)

    If Len(vNewValue) <= 1 Then
        txtText.PasswordChar = vNewValue
        txtText2.PasswordChar = txtText.PasswordChar
        
        PropertyChanged "PasswordChar"
      Else
        MsgBox "Invalid Character Value", vbCritical, "Enhanced Text"
    End If

End Property

'==================================
'   Properties : SpecialCharacter
'==================================

Public Property Get SpecialCharacter() As String
Attribute SpecialCharacter.VB_Description = "This property allows user to Enhanse input type allowing Special Character."

    SpecialCharacter = strSpecialChar

End Property

Public Property Let SpecialCharacter(ByVal vNewValue As String)

    strSpecialChar = vNewValue
    PropertyChanged ("SpecialCharacter")

End Property

'==================================
'   Properties : Tag
'==================================

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Provides the space for Storing the Extra Data."

    Tag = txtText.Tag

End Property

Public Property Let Tag(ByVal vNewValue As String)

    txtText.Tag = vNewValue
    txtText2.Tag = txtText.Tag
    
    PropertyChanged "Tag"

End Property

'==================================
'   Properties : RightToLeft
'==================================

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Specify default Text Right to Left Property enable or not."

    RightToLeft = txtText.RightToLeft

End Property

Public Property Let RightToLeft(ByVal vNewValue As Boolean)

    txtText.RightToLeft = vNewValue
    txtText2.RightToLeft = txtText.RightToLeft
    
    PropertyChanged "RightToLeft"

End Property

'==================================
'   Properties : TextFormat
'==================================

Public Property Get TextFormat() As String
Attribute TextFormat.VB_Description = "Specify The Input Type Formatting Formula. Ex. 0.00 for two decimal Places for Numeric input."

    TextFormat = strFormatString

End Property

Public Property Let TextFormat(ByVal vNewValue As String)

    strFormatString = vNewValue
    If Not IsNull(strFormatString) Then
        txtText.Text = Format$(txtText.Text, strFormatString)
        txtText2.Text = Format$(txtText2.Text, strFormatString)
    End If
    PropertyChanged ("TextFormat")

End Property

'==================================
'   Properties : MaxLength
'==================================

Public Property Get MaxLength() As Integer
Attribute MaxLength.VB_Description = "Specifies The MaxLength or MaxNo of Character for Control. "

    MaxLength = txtText.MaxLength

End Property

Public Property Let MaxLength(ByVal vNewValue As Integer)

    txtText.MaxLength = vNewValue
    txtText2.MaxLength = txtText.MaxLength
    
    PropertyChanged ("MaxLength")

End Property

'==================================
'   Properties : AutoTab
'==================================

Public Property Get AutoTab() As Boolean
Attribute AutoTab.VB_Description = "When this property is True, The control lost focus when the text box is filled upto MaxLength."

    AutoTab = blnAutoTab

End Property

Public Property Let AutoTab(ByVal vNewValue As Boolean)

    blnAutoTab = vNewValue
    PropertyChanged ("AutoTab")

End Property

Private Sub UserControl_Terminate()

    Set fntNormal = Nothing
    Set fntFocus = Nothing

End Sub

'==================================
'   Controls write properties
'==================================

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "InputType", intCharAccept, 0
    PropBag.WriteProperty "CharCase", intCase, 0
    PropBag.WriteProperty "Alignment", txtText.Alignment, 0
    PropBag.WriteProperty "OnFocusSelect", blnFocusSelect, True

    PropBag.WriteProperty "NormalBackColor", oleNormalBackColor, &HFFFFFF
    PropBag.WriteProperty "FocusBackColor", oleFocusBackColor, &HC0FFFF

    PropBag.WriteProperty "NormalGrooveBackColor", oleNormalGrooveBackColor, &H8000000F
    PropBag.WriteProperty "FocusGrooveBackColor", oleFocusGrooveBackColor, &H8000000F

    PropBag.WriteProperty "NormalFontColor", oleNormalFontColor, &H0
    PropBag.WriteProperty "FocusFontColor", oleFocusFontColor, &HC00000
    
    PropBag.WriteProperty "MultiLiner", m_Multi, False
    
    PropBag.WriteProperty "NormalFont", fntNormal
    PropBag.WriteProperty "FocusFont", fntFocus

    PropBag.WriteProperty "Text", txtText.Text, ""
    PropBag.WriteProperty "EnterExitKey", blnExitkey, False
    PropBag.WriteProperty "Enabled", txtText.Enabled, True
    PropBag.WriteProperty "Locked", txtText.Locked, False

    PropBag.WriteProperty "DisabledBackColor", oleDisabledColor, &H8000000F
    PropBag.WriteProperty "DisabledGrooveBackColor", oleDisabledGrooveBackColor, &H8000000F
    
    PropBag.WriteProperty "NormalBorderPattern", intNormalBorderPattern, 1
    PropBag.WriteProperty "FocusBorderPattern", intFocusBorderPattern, 1

    PropBag.WriteProperty "NormalBorderColor", oleNormalBorderColor, &H800000
    
    PropBag.WriteProperty "DisabledBorderColor", oleDisabledBorderColor, &H8000000F
    
    PropBag.WriteProperty "FocusBorderColor", oleFocusBorderColor, &H800000

    PropBag.WriteProperty "PasswordChar", txtText.PasswordChar, False
    PropBag.WriteProperty "Tag", txtText.Tag, Nothing
    PropBag.WriteProperty "RightToLeft", txtText.RightToLeft, False
    PropBag.WriteProperty "TextFormat", strFormatString, Nothing
    PropBag.WriteProperty "MaxLength", txtText.MaxLength, 25
    PropBag.WriteProperty "AutoTab", blnAutoTab, False
    PropBag.WriteProperty "SpecialCharacter", strSpecialChar, Nothing

End Sub

'==================================
'   Controls read properties
'==================================

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    intCharAccept = PropBag.ReadProperty("InputType", 0)
    intCase = PropBag.ReadProperty("CharCase", 0)
    
    txtText.Alignment = PropBag.ReadProperty("Alignment", 0)
    txtText2.Alignment = txtText.Alignment
    
    blnFocusSelect = PropBag.ReadProperty("OnFocusSelect", True)

    oleNormalBackColor = PropBag.ReadProperty("NormalBackColor", "&HFFFFFF")
    oleNormalGrooveBackColor = PropBag.ReadProperty("NormalGrooveBackColor", "&H8000000F")
    oleFocusGrooveBackColor = PropBag.ReadProperty("FocusGrooveBackColor", "&H8000000F")
    oleFocusBackColor = PropBag.ReadProperty("FocusBackColor", "&HC0FFFF")
    
    oleNormalFontColor = PropBag.ReadProperty("NormalFontColor", "&H0")
    oleFocusFontColor = PropBag.ReadProperty("FocusFontColor", "&HC00000")

    Set fntNormal = PropBag.ReadProperty("NormalFont", Nothing)
    Set fntFocus = PropBag.ReadProperty("FocusFont", Nothing)

    txtText.Text = PropBag.ReadProperty("Text", "")
    txtText2.Text = PropBag.ReadProperty("Text", "")
    
    blnExitkey = PropBag.ReadProperty("EnterExitKey", False)
    
    
    txtText.Enabled = PropBag.ReadProperty("Enabled", True)
    txtText.Locked = PropBag.ReadProperty("Locked", False)
    
    txtText2.Enabled = txtText.Enabled
    txtText2.Locked = txtText.Locked
    
    intNormalBorderPattern = PropBag.ReadProperty("NormalBorderPattern", 1)
    intFocusBorderPattern = PropBag.ReadProperty("FocusBorderPattern", 1)
    intDisabledBorderPattern = PropBag.ReadProperty("DisabledBorderPattern", 1)
    
    oleFocusBorderColor = PropBag.ReadProperty("FocusBorderColor", &H800000)
    oleNormalBorderColor = PropBag.ReadProperty("NormalBorderColor", &H800000)

    oleDisabledColor = PropBag.ReadProperty("DisabledBackColor", &H8000000F)
    oleDisabledBorderColor = PropBag.ReadProperty("DisabledBorderColor", &H8000000F)
    oleDisabledGrooveBackColor = PropBag.ReadProperty("DisabledGrooveBackColor", &H8000000F)
    
    txtText.PasswordChar = PropBag.ReadProperty("PasswordChar", Nothing)
    txtText.Tag = PropBag.ReadProperty("Tag", Nothing)
    txtText.RightToLeft = PropBag.ReadProperty("RightToleft", False)
    
    txtText2.PasswordChar = txtText.PasswordChar
    txtText2.Tag = txtText.Tag
    txtText2.RightToLeft = txtText.RightToLeft
        
    strFormatString = PropBag.ReadProperty("TextFormat", Nothing)
    
    txtText.MaxLength = PropBag.ReadProperty("MaxLength", 25)
    txtText2.MaxLength = txtText.MaxLength
    
    blnAutoTab = PropBag.ReadProperty("AutoTab", False)
    strSpecialChar = PropBag.ReadProperty("SpecialCharacter", Nothing)

    If txtText.Enabled = True Or txtText2.Enabled = True Then
        txtText.BackColor = oleNormalBackColor
        txtText2.BackColor = oleNormalBackColor
        shpShape.BorderColor = oleNormalBorderColor
        UserControl.BackColor = oleNormalGrooveBackColor
        shpShape.BorderStyle = intNormalBorderPattern
      ElseIf txtText.Enabled = False Or txtText2.Enabled = False Then
        txtText.BackColor = oleDisabledColor
        txtText2.BackColor = oleDisabledColor
        shpShape.BorderColor = oleDisabledBorderColor
        UserControl.BackColor = oleDisabledGrooveBackColor
        shpShape.BorderStyle = intDisabledBorderPattern
    End If

    txtText.ForeColor = oleNormalFontColor
    txtText2.ForeColor = txtText.ForeColor
    
    m_Multi = PropBag.ReadProperty("MultiLiner", False)
    
    If m_Multi = True Then
        txtText.Visible = False
        txtText2.Visible = True
    Else
        txtText.Visible = True
        txtText2.Visible = False
    End If

    shpShape.BorderStyle = intNormalBorderPattern
    Set txtText.Font = fntNormal
    Set txtText2.Font = fntNormal

End Sub

Sub About()
Attribute About.VB_Description = "Developed by : Priyank Modi, Visite:http://www.geocities.com/priyank_modi/"
Attribute About.VB_UserMemId = -552

    On Error Resume Next
        MsgBox "Enhanced TextControl 2" & vbCrLf & vbCrLf & "Originally prepared by:" & vbCrLf & "Priyank Modi" & vbCrLf & vbCrLf & "Modified with permission by : Francis Arnold Balatico", vbInformation

End Sub

'*******************************************************************************
'   Control commmon Events for txtText2
'*******************************************************************************

Private Sub txtText2_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub txtText2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub txtText2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub txtText2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub txtText2_Change()

    RaiseEvent Change
    If (blnAutoTab = True) Then
        If (CDbl(Len(txtText.Text)) >= CDbl(txtText.MaxLength)) Then 'autotab handling
            SendKeys "{Tab}"
            Exit Sub
        End If
    End If

End Sub

Private Sub txtText2_Click()

    RaiseEvent Click

End Sub

Private Sub txtText2_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub txtText2_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub txtText2_KeyPress(KeyAscii As Integer)

    On Error Resume Next
        RaiseEvent KeyPress(KeyAscii)

        If KeyAscii = 8 Then Exit Sub

        If (Len(txtText2.Text) = 0) Then blnSpaceFlag = True

        If (blnExitkey = True) Then         'ENTER as TAB key
            If KeyAscii = 13 Then
                SendKeys "{Tab}"
                Exit Sub
            End If
        End If

        If (InStr(strSpecialChar, Chr(KeyAscii)) <> 0) Then 'special char handling
            Exit Sub
        End If

        Select Case intCharAccept
          Case 0          '  0 - inpNone

            KeyAscii = ModifyCase(KeyAscii)
            If KeyAscii = 32 Then blnSpaceFlag = True
            Exit Sub

          Case 1          '  1 - inpAlphabetic

            If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                KeyAscii = ModifyCase(KeyAscii)
              ElseIf (KeyAscii = 32) Then
                blnSpaceFlag = True
              Else
                KeyAscii = 0
                Beep
            End If
            Exit Sub

          Case 2          '  2 - inpDecimalNumber

            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 46)) Or isDupDot(2, KeyAscii) = True Then
                KeyAscii = 0
                Beep
            End If
            Exit Sub

          Case 3          '  4 - inpAlphaNumeric

            If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                KeyAscii = ModifyCase(KeyAscii)
              ElseIf ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 46)) Then
                If blnSpaceFlag = True Then blnSpaceFlag = False
              ElseIf (KeyAscii = 32) Then
                blnSpaceFlag = True
              Else
                KeyAscii = 0
                Beep
            End If
            Exit Sub
          Case 4
            If (InStr(strSpecialChar, Chr(KeyAscii)) = 0) Then 'special char handling
                KeyAscii = 0
            End If
            Exit Sub
        End Select

End Sub

Private Sub txtText2_LostFocus()

    blnSpaceFlag = True
    UserControl.BackColor = oleNormalGrooveBackColor

    If txtText2.Enabled = True Then
        txtText2.BackColor = oleNormalBackColor
      Else
        txtText2.BackColor = oleDisabledColor
    End If

    txtText2.ForeColor = oleNormalFontColor
    shpShape.BorderColor = oleNormalBorderColor
    shpShape.BorderStyle = intNormalBorderPattern
    Set txtText2.Font = fntNormal
    If (intCharAccept = 2) Then
        If Not IsNumeric(txtText2.Text) Then txtText2.Text = 0
    End If
    If Not IsNull(strFormatString) Then
        txtText2.Text = Format$(txtText2.Text, strFormatString)
    End If

End Sub

Private Sub txtText2_GotFocus()

    If (blnFocusSelect = True) Then
        txtText2.SelStart = 0
        txtText2.SelLength = Len(txtText2)
    End If

    UserControl.BackColor = oleFocusGrooveBackColor

    If txtText2.Enabled = True Then
        txtText2.BackColor = oleFocusBackColor
      Else
        txtText2.BackColor = oleDisabledColor
    End If

    txtText2.ForeColor = oleFocusFontColor
    shpShape.BorderColor = oleFocusBorderColor
    shpShape.BorderStyle = intFocusBorderPattern
    Set txtText2.Font = fntFocus

End Sub

'***************************************************************************
'* This function checks for existence of a decimal point                   *
'***************************************************************************
Private Function isDupDot(Index As Integer, ByVal KeyAsc As Long) As Boolean
    
    If Index = 1 Then
        If InStr(1, Trim$(txtText.Text), ".", vbBinaryCompare) <> 0 Then
            If KeyAsc = 46 Then
                isDupDot = True
              Else
                isDupDot = False
            End If
          Else
            isDupDot = False
        End If
    Else
        If InStr(1, Trim$(txtText2.Text), ".", vbBinaryCompare) <> 0 Then
            If KeyAsc = 46 Then
                isDupDot = True
              Else
                isDupDot = False
            End If
          Else
            isDupDot = False
        End If
    End If
End Function
