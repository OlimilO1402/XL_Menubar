Attribute VB_Name = "MNew"
Option Explicit

Public Function Menubar(Owner As UserForm, aFraMenubar As MSForms.Frame) As Menubar
    Set Menubar = New Menubar: Menubar.New_ Owner, aFraMenubar
End Function

Public Function MenubarItem(ByVal aText As String, aImgMenubarItem As MSForms.Image, aLblMenubarItem As MSForms.Label, aMenuItemFrame As MSForms.Frame) As MenubarItem
    Set MenubarItem = New MenubarItem: MenubarItem.New_ aText, aImgMenubarItem, aLblMenubarItem, aMenuItemFrame
End Function

Public Function MenuItem(ByVal aText As String, aLabel As MSForms.Label, aImage As MSForms.Image) As MenuItem
    Set MenuItem = New MenuItem: MenuItem.New_ aText, aLabel, aImage
End Function

Public Sub ShowUserForm()
    UserForm1.Show
End Sub

