VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form addtool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Tool"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dialog 
      Left            =   4890
      Top             =   195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   2970
      TabIndex        =   6
      Top             =   1590
      Width           =   1155
   End
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Height          =   450
      Left            =   1650
      TabIndex        =   5
      Top             =   1590
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   4725
      TabIndex        =   4
      Top             =   930
      Width           =   1080
   End
   Begin VB.TextBox tProg 
      Height          =   300
      Left            =   75
      TabIndex        =   3
      Top             =   975
      Width           =   4500
   End
   Begin VB.TextBox tName 
      Height          =   300
      Left            =   75
      TabIndex        =   2
      Top             =   300
      Width           =   4500
   End
   Begin VB.Label Label2 
      Caption         =   "Program:"
      Height          =   225
      Left            =   75
      TabIndex        =   1
      Top             =   750
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   165
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "addtool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

'dialog.filename = ""
'dialog.DialogTitle = "Choose a program..."
'dialog.Filter = "Executable (*.exe) | *.exe"
'dialog.ShowOpen

'If dialog.filename = "" Then Exit Sub

'tProg.Text = dialog.filename

tProg.Text = FileDialog("Choose a program...", "Executable (*.exe) | *.exe", _
                0, "exe")

End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Sub Command4_Click()

If Trim(tName.Text) = "" Or tProg.Text = "" Then
    MsgBox "Need to specify both a name and program path for the tool.", vbCritical, "MicroByte"
    Exit Sub
End If

editorWin.tools.AddItem tName.Text
ToolProg.Add tProg.Text

Open App.Path & "\tools.dat" For Output As #1
    Print #1, ToolProg.itemCount
    For n = 1 To ToolProg.itemCount
        Print #1, editorWin.tools.list(n - 1)
        Print #1, ToolProg.Item(n)
    Next n
Close #1

Unload Me

End Sub

