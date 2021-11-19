VERSION 5.00
Begin VB.Form debugWin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debugging"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "debugWin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Variables:"
      Height          =   1680
      Left            =   120
      TabIndex        =   10
      Top             =   5145
      Width           =   7785
      Begin VB.OptionButton Option2 
         Caption         =   "Local"
         Height          =   300
         Left            =   6810
         TabIndex        =   14
         Top             =   930
         Width           =   750
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Global"
         Height          =   240
         Left            =   6810
         TabIndex        =   13
         Top             =   540
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.ListBox localVars 
         Height          =   1290
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Visible         =   0   'False
         Width           =   6510
      End
      Begin VB.ListBox globalVars 
         Height          =   1290
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   6510
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   420
      Left            =   6975
      TabIndex        =   9
      Top             =   4665
      Width           =   1035
   End
   Begin VB.TextBox runtimeCmd 
      Height          =   300
      Left            =   900
      TabIndex        =   8
      Top             =   4710
      Width           =   6000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Always on top"
      Height          =   255
      Left            =   4335
      TabIndex        =   6
      Top             =   4185
      Width           =   1290
   End
   Begin VB.ListBox stack 
      Height          =   3660
      IntegralHeight  =   0   'False
      Left            =   6120
      TabIndex        =   4
      Top             =   300
      Width           =   1920
   End
   Begin VB.CommandButton auto 
      Caption         =   "Auto Step"
      Height          =   450
      Left            =   1380
      TabIndex        =   3
      Top             =   4095
      Width           =   1080
   End
   Begin VB.CommandButton step 
      Caption         =   "Step"
      Height          =   450
      Left            =   75
      TabIndex        =   2
      Top             =   4095
      Width           =   1080
   End
   Begin VB.CommandButton pause 
      Caption         =   "Pause"
      Height          =   450
      Left            =   2685
      TabIndex        =   1
      Top             =   4095
      Width           =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   570
      Top             =   90
   End
   Begin VB.ListBox code 
      Height          =   3960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6075
   End
   Begin VB.Label Label2 
      Caption         =   "Command:"
      Height          =   210
      Left            =   60
      TabIndex        =   7
      Top             =   4725
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Call stack:"
      Height          =   195
      Left            =   6120
      TabIndex        =   5
      Top             =   45
      Width           =   870
   End
End
Attribute VB_Name = "debugWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Sub auto_Click()

debugState = DS_AUTO


End Sub

Private Sub Check1_Click()

If Check1.value Then
    SetWindowPos debugWin.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
Else
    SetWindowPos debugWin.hwnd, -2, 0, 0, 0, 0, &H1 Or &H2
End If


End Sub

Private Sub Command1_Click()

If Trim(runtimeCmd.Text) = "" Then Exit Sub

RunCmd Trim(runtimeCmd.Text)

End Sub

Private Sub Form_Load()

Dim idxCount As Integer
Dim lineTxt As String

ReadRunFile

LoadFunctions

App.Title = App.EXEName
output.Caption = "Debugging: " & App.EXEName

debugging = True


End Sub


Private Sub Form_Unload(Cancel As Integer)

a = MsgBox("Terminate debugging session?", vbQuestion Or vbYesNo, "Lithium BASIC Debug")

If a = vbNo Then Cancel = 1 Else EndProg


End Sub


Private Sub Option1_Click()

globalVars.visible = True
localVars.visible = False

End Sub

Private Sub Option2_Click()

localVars.visible = True
globalVars.visible = False

End Sub


Private Sub pause_Click()

debugState = DS_PAUSE


End Sub


Private Sub stack_DblClick()

If stack.ListIndex < 1 Then Exit Sub

For a = 1 To debugCode.Item(stack.ListIndex + 1).itemCount
    codeWin.code.AddItem debugCode.Item(stack.ListIndex + 1).Item(a)
Next a

codeWin.code.ListIndex = debugLneSel.Item(stack.ListIndex)
codeWin.Caption = "Code View - " & stack.list(stack.ListIndex)

codeWin.Show vbModal


End Sub


Private Sub step_Click()

debugState = DS_STEP


End Sub


Private Sub Timer1_Timer()

Timer1.Enabled = False

RunProg

output.Caption = "Execution complete: " & App.EXEName
output.display.SelStart = Len(output.display.Text)

If errorFlag Then End


End Sub





