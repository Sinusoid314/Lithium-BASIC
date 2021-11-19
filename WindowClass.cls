VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public winName As String
Public winHandle As Long
Public winDC As Long
Public redrawDC As Long
Public winType As String

Public eventName As New ArrayClass
Public eventSubIdx As New ArrayClass
Public eventSubType As New ArrayClass

Public hMenuBar As Long
Public hMenu As New ArrayClass
Public menuItemID As New ArrayClass
Public menuItemSubIdx As New ArrayClass
Public menuItemSubType As New ArrayClass

Public controls As New ArrayClass

Public sprites As New ArrayClass
Public spriteName As New ArrayClass
Public spriteBG As Long
Public spriteBGX, spriteBGY As Long
