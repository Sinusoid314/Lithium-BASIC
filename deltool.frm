VERSION 5.00
Begin VB.Form deltool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove Tool"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   450
      Left            =   3165
      TabIndex        =   4
      Top             =   3600
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   450
      Left            =   1860
      TabIndex        =   3
      Top             =   3600
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3105
      Width           =   4350
   End
   Begin VB.ListBox tools 
      Height          =   2595
      Left            =   75
      TabIndex        =   0
      Top             =   375
      Width           =   4350
   End
   Begin VB.Label Label1 
      Caption         =   "Tools:"
      Height          =   180
      Left            =   75
      TabIndex        =   1
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "deltool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If tools.ListIndex < 0 Then Exit Sub


Dim sel As Integer

sel = tools.ListIndex + 1

ToolName.Remove sel
ToolProg.Remove sel

'Remove buttons from toolbar
editorWin.toolbar2.Visible = False
editorWin.toolbar2.ImageList = Nothing
For n = editorWin.toolbar2.Buttons.Count To 1 Step -1
    editorWin.toolbar2.Buttons.Remove n
    editorWin.img2.ListImages.Remove n
Next n

'Reset all buttons
If ToolName.itemCount > 0 Then
    editorWin.img2.ImageHeight = 16
    editorWin.img2.ImageWidth = 16
    For n = 1 To ToolName.itemCount
        DrawIcon editorWin.pic.hdc, 0, 0, ExtractIcon(App.hInstance, ToolProg.Item(n), 0)
        editorWin.img2.ListImages.Add , , editorWin.pic.Image
    Next n
    editorWin.toolbar2.ImageList = editorWin.img2
    For n = 1 To ToolName.itemCount
        editorWin.toolbar2.Buttons.Add , , , , n
        editorWin.toolbar2.Buttons.Item(n).ToolTipText = ToolName.Item(n)
    Next n
End If

Open App.Path & "\tools.dat" For Output As #1
    Print #1, ToolName.itemCount
    For n = 1 To ToolName.itemCount
        Print #1, ToolName.Item(n)
        Print #1, ToolProg.Item(n)
    Next n
Close #1

tools.RemoveItem tools.ListIndex
Text1.Text = ""

If ToolName.itemCount > 0 Then
    editorWin.toolbar2.Visible = True
Else
    Unload Me
End If


End Sub
Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

If ToolName.itemCount = 0 Then
    Command1.Enabled = False
Else
    For n = 1 To ToolName.itemCount
        tools.AddItem ToolName.Item(n)
    Next n
End If

End Sub

Private Sub tools_Click()

Text1.Text = ToolProg.Item(tools.ListIndex + 1)

End Sub
