VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Menubar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mOwner As Form 'MSForms.UserForm
Private WithEvents mPanel As PictureBox 'MSForms.Frame
Attribute mPanel.VB_VarHelpID = -1
Private m_IsMenuOpen As Boolean
Private m_Items As Collection
Public Event Click(aMenuItem As MenuItem)
Private m_LastLeft As Single

Private Sub Class_Initialize()
    Set m_Items = New Collection
    m_LastLeft = 6
End Sub

'Friend Sub New_(Owner As UserForm, aFrame As PictureBox) 'MSForms.Frame)
Friend Sub New_(Owner As Form, aFrame As PictureBox)
    Set mOwner = Owner: Set mPanel = aFrame
    mPanel.Left = -1
End Sub

Public Function Add(aMenubarItem As MenubarItem) As MenubarItem
    Set aMenubarItem.Owner = Me
    m_Items.Add aMenubarItem, aMenubarItem.Text
    Set Add = aMenubarItem
    aMenubarItem.Left = m_LastLeft
    m_LastLeft = m_LastLeft + aMenubarItem.Width '+ 1
End Function

Public Sub SelectMenubarItem(aItem As MenubarItem)
    Dim v, obj As MenubarItem
    For Each v In m_Items
        Set obj = v
        If Not obj Is aItem Then
            obj.Deselect
        End If
    Next
End Sub

Public Property Get Item(Key As String) As MenubarItem
    Set Item = m_Items.Item(Key)
End Property

Friend Sub OnClick(aMenuItem As MenuItem)
    m_IsMenuOpen = False
    SelectMenubarItem Nothing
    RaiseEvent Click(aMenuItem)
End Sub

Private Sub mPanel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If m_IsMenuOpen Then Exit Sub
    SelectMenubarItem Nothing
End Sub
Private Sub mPanel_Click()
    Me.IsMenuOpen = False
End Sub

Public Property Get IsMenuOpen() As Boolean
    IsMenuOpen = m_IsMenuOpen
End Property
Public Property Let IsMenuOpen(ByVal Value As Boolean)
    m_IsMenuOpen = Value
    If Not m_IsMenuOpen Then
        OpenMenu Nothing
        SelectMenubarItem Nothing
    End If
End Property

Public Sub IsMenuOpenToggle()
    m_IsMenuOpen = Not m_IsMenuOpen
End Sub

Public Sub OpenMenu(aItem As MenubarItem)
    Dim v, obj As MenubarItem
    For Each v In m_Items
        Set obj = v
        If obj Is aItem Then obj.OpenMenu Else obj.CloseMenu
    Next
End Sub


