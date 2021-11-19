VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Name Register"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   870
      TabIndex        =   4
      Top             =   1605
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   465
      Left            =   1500
      TabIndex        =   1
      Top             =   825
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   660
      TabIndex        =   0
      Top             =   240
      Width           =   3585
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   210
      Left            =   90
      TabIndex        =   3
      Top             =   1650
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   285
      Width           =   510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim nStr, pStr, tmpChar As String

nStr = Text1.Text

For n = 1 To Len(nStr) Step 3
    tmpChar = Chr(Asc(Mid(nStr, n, 1)) + 20 - n)
    If Asc(tmpChar) > 32 And Asc(tmpChar) < 127 Then
        pStr = pStr & tmpChar
    End If
Next n

Text2.Text = pStr

End Sub

