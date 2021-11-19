VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   316
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Apply"
      Height          =   450
      Left            =   3585
      TabIndex        =   7
      Top             =   1905
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   2445
      TabIndex        =   6
      Top             =   1905
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   450
      Left            =   1290
      TabIndex        =   5
      Top             =   1905
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editor"
      Height          =   1530
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   4650
      Begin VB.CommandButton Command4 
         Caption         =   "Change"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3225
         TabIndex        =   10
         Top             =   915
         Width           =   1020
      End
      Begin VB.TextBox editorFont 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   9
         Top             =   945
         Width           =   2295
      End
      Begin VB.PictureBox fcolor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3510
         ScaleHeight     =   240
         ScaleWidth      =   645
         TabIndex        =   4
         Top             =   450
         Width           =   705
      End
      Begin VB.PictureBox bcolor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1290
         ScaleHeight     =   240
         ScaleWidth      =   645
         TabIndex        =   2
         Top             =   450
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Font:"
         Enabled         =   0   'False
         Height          =   225
         Left            =   255
         TabIndex        =   8
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "Forecolor:"
         Height          =   225
         Left            =   2445
         TabIndex        =   3
         Top             =   465
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Backcolor:"
         Height          =   210
         Left            =   255
         TabIndex        =   1
         Top             =   465
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   75
      Top             =   1635
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   33023
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bcolor_Click()

dialog.ShowColor
If Not dialog.CancelError Then bcolor.BackColor = dialog.Color

End Sub
Private Sub Command1_Click()


Command3_Click

Unload Me


End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Command3_Click()


Open App.Path & "\options.dat" For Output As #1
    Print #1, bcolor.BackColor
    Print #1, fcolor.BackColor
Close #1

LoadOptions


End Sub

Private Sub fcolor_Click()

dialog.ShowColor
If Not dialog.CancelError Then fcolor.BackColor = dialog.Color

End Sub
Private Sub Form_Load()

bcolor.BackColor = editorWin.editor.BackColor
fcolor.BackColor = editorWin.editor.SelColor

End Sub
