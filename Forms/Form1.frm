VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Label Label1 
      Caption         =   "This menu here is made with menu-editor of the vb-ide and is just for reference"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent Files"
         Begin VB.Menu mnuFileRecentFile1 
            Caption         =   "File1"
         End
         Begin VB.Menu mnuFileRecentFile2 
            Caption         =   "File2"
         End
         Begin VB.Menu mnuFileRecentFile3 
            Caption         =   "File3"
         End
         Begin VB.Menu mnuFileRecentFile4 
            Caption         =   "File4"
         End
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "Redo"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "?"
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mMenuBar As Menubar
Attribute mMenuBar.VB_VarHelpID = -1

Private Sub UserForm_Click()
    mMenuBar.IsMenuOpen = False
End Sub

Private Sub UserForm_Initialize()
    Set mMenuBar = MNew.Menubar(Me, FraMenubar)
    With mMenuBar
        With .Add(MNew.MenubarItem("File", ImgMnuFile, LblMnuFile, FraMnuFile))
            .Add MNew.MenuItem("New", LblMnuFileNew, ImgMnuFileNew)
            .Add MNew.MenuItem("Open...", LblMnuFileOpen, ImgMnuFileOpen)
            .Add MNew.MenuItem("-", LblMnuFileSep1, Nothing)
            .Add MNew.MenuItem("Save", LblMnuFileSave, ImgMnuFileSave)
            .Add MNew.MenuItem("Save As...", LblMnuFileSaveAs, ImgMnuFileSaveAs)
            .Add MNew.MenuItem("-", LblMnuFileSep2, Nothing)
            .Add MNew.MenuItem("Recent Files", LblMnuFileRecent, ImgMnuFileRecent)
            .Add MNew.MenuItem("-", LblMnuFileSep3, Nothing)
            .Add MNew.MenuItem("Exit", LblMnuFileExit, ImgMnuFileExit)
        End With
        With .Add(MNew.MenubarItem("Edit", ImgMnuEdit, LblMnuEdit, FraMnuEdit))
            .Add MNew.MenuItem("Undo", LblMnuEditUndo, ImgMnuEditUndo)
            .Add MNew.MenuItem("Redo", LblMnuEditRedo, ImgMnuEditRedo)
            .Add MNew.MenuItem("-", LblMnuEditSep1, Nothing)
            .Add MNew.MenuItem("Cut", LblMnuEditCut, ImgMnuEditCut)
            .Add MNew.MenuItem("Copy", LblMnuEditCopy, ImgMnuEditCopy)
            .Add MNew.MenuItem("Paste", LblMnuEditPaste, ImgMnuEditPaste)
        End With
        With .Add(MNew.MenubarItem("View", ImgMnuView, LblMnuView, FraMnuView))
            .Add MNew.MenuItem("Options", LblMnuViewOptions, ImgMnuViewOptions)
        End With
        With .Add(MNew.MenubarItem("?", ImgMnuHelp, LblMnuHelp, FraMnuHelp))
            .Add MNew.MenuItem("Info", LblMnuHelpInfo, ImgMnuHelpInfo)
        End With
    End With
End Sub

Private Sub mMenuBar_Click(aMenuItem As MenuItem)
    Select Case aMenuItem.Text
    Case "New":          mnuFileNew_Click
    Case "Open...":      mnuFileOpen_Click
    Case "Save":         mnuFileSave_Click
    Case "Save As...":   mnuFileSaveAs_Click
    Case "Recent Files": mnuFileRecent_Click
    Case "Exit":         mnuFileExit_Click
    Case "Undo":         mnuEditUndo_Click
    Case "Redo":         mnuEditRedo_Click
    Case "Cut":          mnuEditCut_Click
    Case "Copy":         mnuEditCopy_Click
    Case "Paste":        mnuEditPaste_Click
    Case "Options":      mnuViewOptions_Click
    Case "Info":         mnuHelpInfo_Click
    End Select
    'MsgBox "Clicked: " & aMenuItem.Text
End Sub

Private Sub mnuFileNew_Click()
    MsgBox "mnuFileNew_Click"
End Sub

Private Sub mnuFileOpen_Click()
    MsgBox "mnuFileOpen_Click"
End Sub

Private Sub mnuFileSave_Click()
    MsgBox "mnuFileSave_Click"
End Sub

Private Sub mnuFileSaveAs_Click()
    MsgBox "mnuFileSaveAs_Click"
End Sub

Private Sub mnuFileRecent_Click()
    MsgBox "mnuFileRecent_Click"
End Sub

Private Sub mnuFileExit_Click()
    MsgBox "mnuFileExit_Click"
    Dim mr As VbMsgBoxResult: mr = MsgBox("Save changes?", vbYesNoCancel)
    If mr = vbCancel Then Exit Sub
    If mr = vbOK Then
        'Save your data now!
        MsgBox "All Changes successfully saved!"
    End If
    Unload Me
End Sub '

Private Sub mnuEditUndo_Click()
    MsgBox "mnuEditUndo_Click"
End Sub
Private Sub mnuEditRedo_Click()
    MsgBox "mnuEditRedo_Click"
End Sub
Private Sub mnuEditCut_Click()
    MsgBox "mnuEditCut_Click"
End Sub
Private Sub mnuEditCopy_Click()
    MsgBox "mnuEditCopy_Click"
End Sub
Private Sub mnuEditPaste_Click()
    MsgBox "mnuEditPaste_Click"
End Sub
Private Sub mnuViewOptions_Click()
    MsgBox "mnuViewOptions_Click"
End Sub
Private Sub mnuHelpInfo_Click()
    MsgBox "mnuHelpInfo_Click"
End Sub

'Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    '
'End Sub
'
'Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    '
'End Sub
'
'Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    '
'End Sub

