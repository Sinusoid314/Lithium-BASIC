VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form editorWin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Lithium BASIC v1.03"
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "editorWin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   712
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox editor 
      Height          =   3795
      Left            =   795
      TabIndex        =   6
      Top             =   735
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   6694
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"editorWin.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   7170
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1365
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   2835
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   33023
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   7335
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":0D43
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":10DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":147B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":1817
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":1BB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":1F4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":22EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":2687
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":2A23
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "editorWin.frx":2DBF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statusbar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   7380
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   "Ready."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "run"
            Object.ToolTipText     =   "Run Program"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "debug"
            Object.ToolTipText     =   "Debug Program"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edittools"
            Object.ToolTipText     =   "Edit Tools"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "about"
            Object.ToolTipText     =   "About MicroByte"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   4785
         ScaleHeight     =   315
         ScaleWidth      =   660
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Width           =   660
         Begin VB.Label Label1 
            Caption         =   "Tools:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   45
            TabIndex        =   5
            Top             =   45
            Width           =   480
         End
      End
      Begin VB.ComboBox tools 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   3480
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "No File"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "No File"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "No File"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "No File"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "No File"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
      End
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "&Program"
      Begin VB.Menu mnuProgramRun 
         Caption         =   "&Run Code"
      End
      Begin VB.Menu mnuProgramDebug 
         Caption         =   "&Debug Code"
      End
      Begin VB.Menu mnuProgramSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramCreate 
         Caption         =   "&Create Runtime File"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsEdit 
         Caption         =   "&Edit Tools"
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsReg 
         Caption         =   "Register Lithium BASIC"
      End
      Begin VB.Menu mnuToolsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpFeatures 
         Caption         =   "Lithium BASIC 1.03 &Features"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpSite 
         Caption         =   "Visit the Lithium BASIC Web Site..."
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Lithium BASIC"
      End
   End
End
Attribute VB_Name = "editorWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const WM_USER As Long = &H400
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const EM_UNDO = &HC7

Private editText As String
Private filename As String

Private RecentFiles(0 To 4) As String
Private Sub AddRecentFile(fName As String)

'Check if the new filename is already there
For n = 0 To 4
    If fName = RecentFiles(n) Then Exit Sub
Next n

'Add new filename
For n = 4 To 1 Step -1
    RecentFiles(n) = RecentFiles(n - 1)
Next n
RecentFiles(0) = fName
  
'Reload menu items with filename data
For n = 0 To 4
  If RecentFiles(n) = "" Then
    mnuFileRecent(n).Caption = "No File"
    mnuFileRecent(n).Enabled = False
  Else
    mnuFileRecent(n).Caption = "&" & (n + 1) & ". " & RecentFiles(n)
    mnuFileRecent(n).Enabled = True
  End If
Next n

If Dir(App.Path & "\recent.dat") = "" Then Exit Sub

Open App.Path & "\recent.dat" For Output As #1
    For n = 0 To 4
        Print #1, RecentFiles(n)
    Next n
Close #1


End Sub

Private Sub LoadFile(fName As String)

  If editText <> editor.Text Then
    r = MsgBox("Save changes made to '" & filename & "' ?", vbYesNoCancel, "Lithium BASIC")
        If r = vbYes Then Call mnuFileSave_Click
        If r = vbCancel Then Exit Sub
  End If

  Open fName For Input As #1
    tmpTxt = Input(LOF(1), 1)
  Close #1
  
  editor.Text = tmpTxt
  editText = editor.Text
  filename = fName
  
  editorWin.Caption = "Lithium BASIC v1.0 - [" & fName & "]"
  statusbar.SimpleText = "Ready."
  
  AddRecentFile fName

End Sub

Private Sub LoadRecentFiles()

If Dir(App.Path & "\recent.dat") = "" Then Exit Sub

Dim tmpFile As String


Open App.Path & "\recent.dat" For Input As #1
    If EOF(1) Then
        Close #1
        Exit Sub
    End If
    For n = 0 To 4
        Input #1, tmpFile
        RecentFiles(n) = Trim(tmpFile)
        If Trim(tmpFile) = "" Then
            mnuFileRecent(n).Caption = "No File"
        Else
            mnuFileRecent(n).Caption = "&" & (n + 1) & ". " & tmpFile
            mnuFileRecent(n).Enabled = True
        End If
    Next n
Close #1


End Sub
Private Sub LoadTools()

If Dir(App.Path & "\tools.dat") = "" Then Exit Sub

Dim num As Integer
Dim tName, tProg As String
Dim hPic As Long

Open App.Path & "\tools.dat" For Input As #1
    Input #1, num
    If num = 0 Then
        Close #1
        Exit Sub
    End If
    For n = 1 To num
        Input #1, tName
        Input #1, tProg
        tools.AddItem tName
        ToolName.Add tName
        ToolProg.Add tProg
    Next n
Close #1

End Sub

Private Sub RemoveRecentFile(ByVal mnuIdx As Integer)

'Check if mnuIdx is in bounds
If mnuIdx < 0 Or mnuIdx > 4 Then Exit Sub

'Remove filename
For n = mnuIdx To 3
    RecentFiles(n) = RecentFiles(n + 1)
Next n
RecentFiles(4) = ""

'Reload menu items with filename data
For n = 0 To 4
  If RecentFiles(n) = "" Then
    mnuFileRecent(n).Caption = "No File"
    mnuFileRecent(n).Enabled = False
  Else
    mnuFileRecent(n).Caption = "&" & (n + 1) & ". " & RecentFiles(n)
    mnuFileRecent(n).Enabled = True
  End If
Next n

If Dir(App.Path & "\recent.dat") = "" Then Exit Sub

Open App.Path & "\recent.dat" For Output As #1
    For n = 0 To 4
        Print #1, RecentFiles(n)
    Next n
Close #1


End Sub

Private Sub editor_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then
    KeyAscii = 0
    editor.SelText = Space(4)
End If

End Sub


Private Sub Form_Load()

LoadFunctions

compileError = False

filename = ""
editText = ""

editorWin.Width = Int((630 / 800) * Screen.Width)
editorWin.Height = Int((480 / 600) * Screen.Height)
SendMessage editor.hwnd, EM_SETTARGETDEVICE, 0, 1

LoadOptions
LoadTools
LoadRecentFiles
LoadMbr

startup.Show vbModal

If Not isMbr Then mnuProgramCreate.Enabled = False

If Trim(Command$) <> "" Then
  LoadFile Mid(Command$, 2, Len(Command$) - 2)
End If

End Sub
Private Sub Form_Resize()

On Error Resume Next

editor.Move 0, 28, Me.ScaleWidth, Me.ScaleHeight - 45


End Sub


Private Sub Form_Unload(Cancel As Integer)

If editText <> editor.Text Then
  r = MsgBox("Save changes made to '" & filename & "' ?", vbYesNoCancel, "Lithium BASIC")
      If r = vbYes Then Call mnuFileSave_Click
      If r = vbCancel Then Cancel = 1
End If

If Not isMbr Then startup.Show

End Sub

Private Sub mnuEditCopy_Click()

    Clipboard.SetText editor.SelText

End Sub

Private Sub mnuEditCut_Click()

    Clipboard.SetText editor.SelText
    editor.SelText = ""

End Sub

Private Sub mnuEditFind_Click()

findDlg.Show 0, Me

End Sub

Private Sub mnuEditPaste_Click()

    editor.SelText = Clipboard.GetText

End Sub

Private Sub mnuEditSelAll_Click()

editor.SelStart = 0
editor.SelLength = Len(editor.Text)

End Sub

Private Sub mnuEditUndo_Click()

SendMessage editor.hwnd, EM_UNDO, 0, 0

End Sub

Private Sub mnuFileExit_Click()

Unload Me

End Sub

Private Sub mnuFileNew_Click()
  If editText <> editor.Text Then
    r = MsgBox("Save changes made to '" & filename & "' ?", vbYesNoCancel, "Lithium BASIC")
        If r = vbYes Then Call mnuFileSave_Click
        If r = vbCancel Then Exit Sub
  End If

editorWin.Caption = "Lithium BASIC v1.0"
statusbar.SimpleText = "Ready."

editor.Text = ""
filename = ""
editText = ""

End Sub

Private Sub mnuFileOpen_Click()

Dim tmpTxt As String
'Dim tmpFile As String

  dialog.filename = ""
  dialog.Filter = "BASIC Source File (*.bas) | *.bas"
  dialog.DialogTitle = "Open..."
  dialog.ShowOpen
  If dialog.filename = "" Then Exit Sub
  
  'tmpFile = FileDialog("Open...", "BASIC Source File (*.bas) | *.bas", _
  '              0, "bas")
  'If tmpFile = "" Then Exit Sub
  
  LoadFile dialog.filename
  
End Sub

Private Sub mnuFilePrint_Click()

On Error Resume Next

dialog.DialogTitle = "Print Code..."
dialog.CancelError = True
dialog.Flags = cdlPDReturnDC + cdlPDNoPageNums
If editor.SelLength = 0 Then
    dialog.Flags = dialog.Flags + cdlPDAllPages
Else
    dialog.Flags = dialog.Flags + cdlPDSelection
End If
dialog.ShowPrinter
If Err <> cdlCancel Then
    editor.SelPrint dialog.hDC
End If

End Sub

Private Sub mnuFileRecent_Click(Index As Integer)

On Error GoTo openError

If Dir(RecentFiles(Index)) = "" Then GoTo openError
LoadFile RecentFiles(Index)

Exit Sub

openError:
    MsgBox "File not found", vbCritical, "Lithium BASIC Error"
    RemoveRecentFile Index

End Sub
Private Sub mnuFileSave_Click()

'Dim tmpFile As String

If filename = "" Then
    dialog.filename = ""
    dialog.DialogTitle = "Save As..."
    dialog.Filter = "BASIC Source File (*.bas) | *.bas"
    dialog.ShowSave
    If dialog.filename = "" Then Exit Sub
    'tmpFile = FileDialog("Save As...", "BASIC Source File (*.bas) | *.bas", _
    '                1, "bas")
    'If tmpFile = "" Then Exit Sub
    filename = dialog.filename
    Open filename For Output As #1
        Print #1, editor.Text;
    Close #1
    editorWin.Caption = "Lithium BASIC v1.0 - [" & filename & "]"
    editText = editor.Text
    AddRecentFile filename
Else
    Open filename For Output As #1
        Print #1, editor.Text;
    Close #1
    editText = editor.Text
End If

End Sub


Private Sub mnuFileSaveAs_Click()

'Dim tmpFile As String

    dialog.filename = ""
    dialog.DialogTitle = "Save As..."
    dialog.Filter = "BASIC Source File (*.bas) | *.bas"
    dialog.ShowSave
    If dialog.filename = "" Then Exit Sub
    'tmpFile = FileDialog("Save As...", "BASIC Source File (*.bas) | *.bas", _
    '                1, "bas")
    'If tmpFile = "" Then Exit Sub
    Open dialog.filename For Output As #1
        Print #1, editor.Text;
    Close #1
    
    AddRecentFile dialog.filename

End Sub


Private Sub mnuHelpAbout_Click()

about.Show vbModal

End Sub



Private Sub mnuHelpContents_Click()

ShellExecute 0, "open", App.Path & "\help\lib help.chm", "", "", 1

End Sub

Private Sub mnuHelpFeatures_Click()

Shell "notepad " & App.Path & "\version 1.02.txt"

End Sub

Private Sub mnuHelpSite_Click()

ShellExecute 0, "open", "http://sircodezalot.britcoms.com/", "", "", 1

End Sub

Private Sub mnuProgramCreate_Click()

If Trim(editor.Text) = "" Then Exit Sub

If Not isMbr Then
    MsgBox "This unregistered copy of Lithium BASIC does not support the auto-deployment feature.", vbCritical, "Lithium BASIC"
    Exit Sub
End If

Dim compilerObj As CompilerClass
Dim tmpTxt As String

    Set compilerObj = New CompilerClass

mnuProgram.Enabled = False
toolbar1.Buttons(9).Enabled = False

compilerObj.Compile editor.Text
    If compileError Then
        compileError = False
        GoTo doneSub
    End If
    
statusbar.SimpleText = "Requesting EXE info..."
    dialog.filename = ""
    dialog.DialogTitle = "Save As..."
    dialog.Filter = "Lithium BASIC Program(*.exe) | *.exe"
    dialog.ShowSave
    If dialog.filename = "" Then GoTo doneSub
    
If Dir(App.Path & "\runtime.exe") = "" Then
    MsgBox "The file runtime.exe is missing from the Lithium BASIC directory. Runtime file cannot be created", _
            vbCritical, "Lithium BASIC"
    GoTo doneSub
End If

Open App.Path & "\runtime.exe" For Binary As #1
    tmpTxt = Space(LOF(1))
    Get #1, , tmpTxt
Close #1
Open dialog.filename For Binary As #2
    Put #2, , tmpTxt
Close #2

statusbar.SimpleText = "Writing EXE..."
WriteRunFile compilerObj.cmdList, dialog.filename

doneSub:
statusbar.SimpleText = "Ready."
mnuProgram.Enabled = True
toolbar1.Buttons(9).Enabled = True

End Sub

Private Sub mnuProgramDebug_Click()

If Trim(editor.Text) = "" Then Exit Sub

Dim compilerObj As CompilerClass

Set compilerObj = New CompilerClass

mnuProgram.Enabled = False
toolbar1.Buttons(9).Enabled = False

compilerObj.Compile editor.Text
    If compileError Then
        compileError = False
        mnuProgram.Enabled = True
        toolbar1.Buttons(9).Enabled = True
        Exit Sub
    End If
    
statusbar.SimpleText = "Writing debug EXE..."
WriteRunFile compilerObj.cmdList, App.Path & "\debug.exe"
    
statusbar.SimpleText = "Launching debugger..."
Shell App.Path & "\debug.exe", vbNormalFocus

statusbar.SimpleText = "Compile successful."

mnuProgram.Enabled = True
toolbar1.Buttons(9).Enabled = True


End Sub

Private Sub mnuProgramRun_Click()

If Trim(editor.Text) = "" Then Exit Sub

Dim compilerObj As CompilerClass

Set compilerObj = New CompilerClass

mnuProgram.Enabled = False
toolbar1.Buttons(9).Enabled = False

compilerObj.Compile editor.Text
    If compileError Then
        compileError = False
        mnuProgram.Enabled = True
        toolbar1.Buttons(9).Enabled = True
        Exit Sub
    End If
    
statusbar.SimpleText = "Writing runtime EXE..."
WriteRunFile compilerObj.cmdList, App.Path & "\runtime.exe"

statusbar.SimpleText = "Launching runtime engine..."
Shell App.Path & "\runtime.exe", vbNormalFocus

statusbar.SimpleText = "Compile successful."

mnuProgram.Enabled = True
toolbar1.Buttons(9).Enabled = True

End Sub



Private Sub mnuToolsEdit_Click()

editTools.Show vbModal

End Sub

Private Sub mnuToolsOptions_Click()

options.Show vbModal

End Sub

Private Sub mnuToolsReg_Click()

Load regWin
regWin.Show vbModal

End Sub

Private Sub toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case Button.Key
  Case "new"
    Call mnuFileNew_Click
  Case "open"
    Call mnuFileOpen_Click
  Case "save"
    Call mnuFileSave_Click
  Case "cut"
    Call mnuEditCut_Click
  Case "copy"
    Call mnuEditCopy_Click
  Case "paste"
    Call mnuEditPaste_Click
  Case "run"
    Call mnuProgramRun_Click
  Case "debug"
    Call mnuProgramDebug_Click
  Case "edittools"
    Call mnuToolsEdit_Click
  Case "about"
    Call mnuHelpAbout_Click
End Select


End Sub

Private Sub tools_Click()

If Dir(ToolProg.Item(tools.ListIndex + 1)) = "" Then Exit Sub

Shell ToolProg.Item(tools.ListIndex + 1), vbNormalFocus

End Sub


