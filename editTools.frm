VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form editTools 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Tools"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   236
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dialog 
      Left            =   3465
      Top             =   2895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox progName 
      Height          =   300
      Left            =   3045
      TabIndex        =   8
      Top             =   765
      Width           =   3525
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Browse..."
      Height          =   450
      Left            =   3045
      TabIndex        =   7
      Top             =   2175
      Width           =   1125
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   450
      Left            =   5355
      TabIndex        =   6
      Top             =   3000
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Tool"
      Height          =   450
      Left            =   1350
      TabIndex        =   5
      Top             =   3000
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Tool"
      Height          =   450
      Left            =   75
      TabIndex        =   4
      Top             =   3000
      Width           =   1125
   End
   Begin VB.TextBox progPath 
      Height          =   300
      Left            =   3045
      TabIndex        =   3
      Top             =   1740
      Width           =   3525
   End
   Begin VB.ListBox programs 
      Height          =   2400
      Left            =   75
      TabIndex        =   0
      Top             =   420
      Width           =   2805
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      Height          =   195
      Left            =   3045
      TabIndex        =   9
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "Filename:"
      Height          =   180
      Left            =   3060
      TabIndex        =   2
      Top             =   1470
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Programs:"
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "editTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Dim nameStr As String

nameStr = InputBox("Enter programs name:", "MicroByte", "Program" & str(ToolName.itemCount + 1))
    If Trim(nameStr) = "" Then Exit Sub

ToolName.Add nameStr
ToolProg.Add ""

programs.AddItem nameStr
programs.ListIndex = programs.ListCount - 1

progName.Text = nameStr
progPath.Text = ""

End Sub
Private Sub Command2_Click()

If programs.ListIndex < 0 Then Exit Sub

res = MsgBox("Are you sure you want to remove tool '" & programs.list(programs.ListIndex) & "'?", vbYesNo, "MicroByte")
    If res = vbNo Then Exit Sub

ToolName.Remove programs.ListIndex + 1
ToolProg.Remove programs.ListIndex + 1

programs.RemoveItem programs.ListIndex

progName.Text = ""
progPath.Text = ""

End Sub
Private Sub Command3_Click()

editorWin.tools.Clear

Open App.Path & "\tools.dat" For Output As #1
    Print #1, ToolProg.itemCount
    For n = 1 To ToolProg.itemCount
        Print #1, ToolName.Item(n)
        Print #1, ToolProg.Item(n)
        editorWin.tools.AddItem ToolName.Item(n)
    Next n
Close #1

Unload Me


End Sub


Private Sub Command4_Click()

'Dim tmpFile As String

dialog.filename = ""
dialog.DialogTitle = "Choose a program..."
dialog.Filter = "Executable (*.exe) | *.exe"
dialog.ShowOpen

If dialog.filename = "" Then Exit Sub

progPath.Text = dialog.filename

'tmpFile = FileDialog("Choose a program...", "Executable (*.exe) | *.exe", _
'                0, "exe")
'If tmpFile = "" Then Exit Sub
'progPath.Text = tmpFile

End Sub

Private Sub Form_Load()

For n = 1 To ToolProg.itemCount
    programs.AddItem ToolName.Item(n)
Next n

End Sub


Private Sub progName_Change()

If programs.ListIndex < 0 Then Exit Sub

ToolName.Item(programs.ListIndex + 1) = progName.Text
programs.list(programs.ListIndex) = progName.Text

End Sub


Private Sub progPath_Change()

If programs.ListIndex < 0 Then Exit Sub

ToolProg.Item(programs.ListIndex + 1) = progPath.Text


End Sub


Private Sub programs_Click()

progName.Text = ToolName.Item(programs.ListIndex + 1)
progPath.Text = ToolProg.Item(programs.ListIndex + 1)

End Sub


