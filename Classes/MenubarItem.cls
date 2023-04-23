VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenubarItem"
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
Private mMenuItemFrame    As PictureBox 'MSForms.Frame
Private m_Menubar         As Menubar
Private Const MouseOverBorderColor As Long = SystemColorConstants.vbHighlight ' = &HD77800 ' = RGB(  0, 120, 215)
Private Const MouseOverBackColor   As Long = &HF3D7B3 ' = RGB(179, 215, 243)
Private Const MouseDownBorderColor As Long = &H808080 ' = RGB(128, 128, 128) ' grey!!!
Private Const MouseDownBackColor   As Long = &HFFFFFF ' = RGB(255, 255, 255)
Private m_IsMouseDown As Boolean
Private m_MouseDownX  As Single
Private m_MouseDownY  As Single
Private m_SubItems    As Collection
Private m_LastTop     As Single

Private Sub Class_Initialize()
    Set m_SubItems = New Collection
End Sub

'Friend Sub New_(ByVal aText As String, aImage As MSForms.Image, aLabel As MSForms.Label, aMenuItemFrame As PictureBoxMSForms.Frame)
Friend Sub New_(ByVal aText As String, aImage As Image, aLabel As Label, aMenuItemFrame As PictureBox)
    m_Text = aText: Set mImage = aImage: Set mLabel = aLabel: Set mMenuItemFrame = aMenuItemFrame
    aMenuItemFrame.Visible = False
    mLabel.WordWrap = False
    mLabel.Caption = m_Text
    mLabel.AutoSize = False
    mLabel.AutoSize = True
    mImage.Top = 3
    mImage.Width = 12 + Math.Round(mLabel.Width, 0)
    mLabel.Top = 6
End Sub

Public Function Add(aMenuItem As MenuItem) As MenuItem
    Set aMenuItem.Owner = Me
    m_SubItems.Add aMenuItem
    Set Add = aMenuItem
    aMenuItem.Top = m_LastTop
    m_LastTop = m_LastTop + aMenuItem.Height '+ 1
    SetFrameSize
End Function

Friend Property Set Owner(ByVal Value As Menubar)
    Set m_Menubar = Value
End Property

Friend Property Get Owner() As Menubar
    Set Owner = m_Menubar
End Property

Private Sub SetFrameSize()
    Dim maxW As Single, maxH As Single: GetMaxSize maxW, maxH
    mMenuItemFrame.Width = 4 + maxW
    mMenuItemFrame.Height = 4 + maxH
    Dim v, obj As MenuItem
    For Each v In m_SubItems
        Set obj = v: obj.Width = maxW
    Next
End Sub

Private Sub GetMaxSize(W_out As Single, H_out As Single)
    Dim v, obj As MenuItem
    For Each v In m_SubItems
        Set obj = v
        W_out = Max(W_out, obj.Width)
        H_out = H_out + obj.Height
    Next
End Sub

Private Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

Friend Sub Deselect()
    DrawMenuItem_Clear
    m_IsMouseDown = False
    If Not mMenuItemFrame Is Nothing Then
        'and all MenuItems
        Dim v, mi As MenuItem
        For Each v In m_SubItems
            Set mi = v: mi.Deselect
        Next
        mMenuItemFrame.Visible = False
    End If
End Sub

Public Sub SelectMenuItem(aItem As MenuItem)
    Dim v, obj As MenuItem
    For Each v In m_SubItems
        Set obj = v
        If Not obj Is aItem Then
            obj.Deselect
        End If
    Next
End Sub

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Get Left() As Single
    Left = mImage.Left
End Property
Public Property Let Left(ByVal Value As Single)
    mImage.Left = Value
    mLabel.Left = Value + 6
End Property

Public Property Get Width() As Single
    Width = 12 + Math.Round(mLabel.Width, 0)
End Property

Private Sub mLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mImage_MouseDown Button, Shift, X, Y
End Sub
Private Sub mImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = XlMouseButton.xlPrimaryButton Then
        m_IsMouseDown = True
        m_Menubar.IsMenuOpen = True
        m_MouseDownX = X
        m_MouseDownY = Y
        DrawMenuItem_Open
        OpenMenu
    End If
End Sub

Public Sub CloseMenu()
    mMenuItemFrame.Visible = False
    m_IsMouseDown = False
End Sub

Public Sub OpenMenu()
    mMenuItemFrame.Visible = True
    mMenuItemFrame.Left = mImage.Left + 0.75
    mMenuItemFrame.Top = mImage.Top + mImage.Height
    mMenuItemFrame.ZOrder 0
End Sub

Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mImage_MouseMove Button, Shift, X, Y
End Sub
Private Sub mImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    m_Menubar.SelectMenubarItem Me
    If m_Menubar.IsMenuOpen Then
        m_Menubar.OpenMenu Me
        DrawMenuItem_Open
        Exit Sub
    End If
    DrawMenuItem_Hover
End Sub

Private Sub DrawMenuItem_Clear()
    mImage.BorderStyle = fmBorderStyleNone
    mImage.BackStyle = fmBackStyleTransparent
End Sub

Private Sub DrawMenuItem_Hover()
    mImage.BorderStyle = fmBorderStyleSingle
    mImage.BorderColor = MouseOverBorderColor
    mImage.BackStyle = fmBackStyleOpaque
    mImage.BackColor = MouseOverBackColor
End Sub

Private Sub DrawMenuItem_Open()
    mImage.BorderStyle = fmBorderStyleSingle
    mImage.BorderColor = MouseDownBorderColor
    mImage.BackStyle = fmBackStyleOpaque
    mImage.BackColor = MouseDownBackColor
End Sub


