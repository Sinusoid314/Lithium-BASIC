VERSION 5.00
Begin VB.Form codeWin 
   Caption         =   "Code View"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   405
      Left            =   4905
      TabIndex        =   1
      Top             =   3180
      Width           =   1065
   End
   Begin VB.ListBox code 
      Height          =   2925
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6090
   End
End
Attribute VB_Name = "codeWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me


End Sub

Private Sub Form_Resize()

On Error Resume Next

Command1.Top = Me.Height - 840
Command1.Left = Me.Width - 1290

code.Width = Me.Width - 100
code.Height = Me.Height - 900


End Sub


