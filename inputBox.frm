VERSION 5.00
Begin VB.Form inputWin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "inputBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2730
      TabIndex        =   3
      Top             =   1155
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   435
      Left            =   1230
      TabIndex        =   2
      Top             =   1155
      Width           =   1020
   End
   Begin VB.TextBox inputVal 
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   660
      Width           =   4800
   End
   Begin VB.Label prompt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   4665
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "inputWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

userInput = inputVal.Text
Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub


