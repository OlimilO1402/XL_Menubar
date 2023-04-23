VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Text As String
Private WithEvents mImage As Image 'MSForms.Image
Attribute mImage.VB_VarHelpID = -1
Private WithEvents mLabel As Label 'MSForms.Label
Attribute mLabel.VB_VarHelpID = -1
Private WithEvents mShort As Label 'MSForms.Label 'Another Label for Keyboard ShortCuts 'not yet implemented
Attribute mShort.VB_VarHelpID = -1
Private WithEvents mArrow As Image 'MSForms.Image 'Another Image for Arrow for SubItems 'not yet implemented
Attribute mArrow.VB_VarHelpID = -1
Private WithEvents mSubFrame As PictureBox 'MSForms.Frame 'Another Frame For holding the SubItems  'not yet implemented
Attribute mSubFrame.VB_VarHelpID = -1
Private mMenuItemFrame As PictureBox 'MSForms.Frame
Private m_MenubarItem  As MenubarItem
Private Const MouseOverBorderColor As Long = SystemColorConstants.vbHighlight ' = &HD77800 ' = RGB(  0, 120, 215)
Private Const MouseOverBackColor   As Long = &HF3D7B3 ' = RGB(179, 215, 243)
Private Const MouseDownBorderColor As Long = SystemColorConstants.vbHighlight ' = &HD77800 ' = RGB(  0, 120, 215)
Private Const MouseDownBackColor   As Long = &HEBBC80 ' = RGB(128, 188, 235)
Private Const LblSeparatorColor    As Long = &HBDBDBD ' = RGB(189, 189, 189)
Private m_IsMouseDown As Boolean
Private m_MouseDownX As Single
Private m_MouseDownY As Single
Private m_SubItems As Collection

Private Sub Class_Initialize()
    Set m_SubItems = New Collection
End Sub

'Friend Sub New_(ByVal aText As String, aLabel As MSForms.Label, aImage As MSForms.Image, Optional aLblShortCut As MSForms.Label, Optional aImgArrowSubItems As MSForms.Image, Optional aFraSubItems As MSForms.Frame)
Friend Sub New_(ByVal aText As String, aLabel As Label, aImage As Image, Optional aLblShortCut As Label, Optional aImgArrowSubItems As Image, Optional aFraSubItems As PictureBox)
    m_Text = aText: Set mLabel = aLabel:     Set mImage = aImage
    mLabel.Left = 20
    If mImage Is Nothing Then
        mLabel.BorderColor = LblSeparatorColor
        Exit Sub
    End If
    Me.Text = aText
    mImage.BorderStyle = fmBorderStyleNone
    mImage.PictureAlignment = fmPictureAlignmentBottomLeft
    mImage.Height = 18
End Sub

Friend Property Set Owner(ByVal Value As MenubarItem)
    Set m_MenubarItem = Value
End Property

Friend Sub Deselect()
    DrawMenuItem_Clear
End Sub

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal Value As String)
    m_Text = Value
    mLabel.WordWrap = False
    mLabel.Caption = m_Text
    mLabel.AutoSize = False
    mLabel.AutoSize = True
End Property

Public Property Get Height() As Single
    If mImage Is Nothing Then Height = 2.25: Exit Property
    Height = mImage.Height
End Property
'Public Property Let Height(ByVal Value As Single)
'    mLabel.Height = Value
'End Property

Public Property Get Top() As Single
    If mImage Is Nothing Then Top = mLabel.Top: Exit Property
    Top = mImage.Top
End Property
Public Property Let Top(ByVal Value As Single)
    If mImage Is Nothing Then
        mLabel.Top = Value + 0.75 '2.5
        mLabel.Height = 0.75
        Exit Property
    End If
    mImage.Top = Value
    mLabel.Top = Value + 3
End Property

Public Property Get Width() As Single
    If mImage Is Nothing Then Exit Property
    Width = mLabel.Left - mImage.Left + mLabel.Width + 20
End Property
Public Property Let Width(ByVal Value As Single)
    If mImage Is Nothing Then
        mLabel.Width = Value - 20
        Exit Property
    End If
    mImage.Width = Value
End Property

'Private Sub mLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Private Sub mLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mImage_MouseDown Button, Shift, X, Y
End Sub
'Private Sub mImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Private Sub mImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = XlMouseButton.xlPrimaryButton Then
        m_IsMouseDown = True
        m_MouseDownX = X
        m_MouseDownY = Y
        DrawMenuItem_MouseDown
    End If
End Sub

'Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Private Sub mLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mImage_MouseMove Button, Shift, X, Y
End Sub
'Private Sub mImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Private Sub mImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_MenubarItem.SelectMenuItem Me
    If m_IsMouseDown Then
        If (25 < Abs(m_MouseDownX - X)) Or (25 < Abs(m_MouseDownY - Y)) Then Exit Sub
    End If
    DrawMenuItem_MouseOver
End Sub

Private Sub mLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mImage_MouseUp Button, Shift, X, Y
End Sub
Private Sub mImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = XlMouseButton.xlPrimaryButton Then
        m_IsMouseDown = False
        Deselect
        'safety for the wrong click, just drag the mouse away and nothing happens
        If (25 < Abs(m_MouseDownX - X)) Or (25 < Abs(m_MouseDownY - Y)) Then Exit Sub
        m_MenubarItem.Owner.OnClick Me
    End If
End Sub

Private Sub DrawMenuItem_Clear()
    If mImage Is Nothing Then Exit Sub
    mImage.BorderStyle = fmBorderStyleNone
    mImage.BackStyle = fmBackStyleTransparent
End Sub

Private Sub DrawMenuItem_MouseOver()
    If mImage Is Nothing Then Exit Sub
    mImage.BorderStyle = fmBorderStyleSingle
    mImage.BorderColor = MouseOverBorderColor
    mImage.BackStyle = fmBackStyleOpaque
    mImage.BackColor = MouseOverBackColor
End Sub

Private Sub DrawMenuItem_MouseDown()
    If mImage Is Nothing Then Exit Sub
    mImage.BorderStyle = fmBorderStyleSingle
    mImage.BorderColor = MouseDownBorderColor
    mImage.BackStyle = fmBackStyleOpaque
    mImage.BackColor = MouseDownBackColor
End Sub


