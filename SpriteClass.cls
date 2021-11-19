VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpriteClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'THIS CLASS HAS BEEN MODIFIED TO RUN SPECIFICALLY IN
'THE Lithium BASIC RUNTIME ENGINE PROJECT

Public drawDC As Long
Public isPlaying, visible As Boolean
Public left, top, width, height, rate As Long
Private initWidth, initHeight, currTick, lastTick As Long
Public frame, cycle As Integer
Public display As Integer
Private cycleCount As Integer
Private scaling As Integer

Private spriteDC As New Collection

Private Const SD_NORMAL As Integer = 0
Private Const SD_FLIP As Integer = 1
Private Const SD_MIRROR As Integer = 2
Private Const SD_ROTATE180 As Integer = 3

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Boolean
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const OBJ_BITMAP = 7
Sub AddFrame(ByVal picName As String)

Dim hBMP As Long
Dim bmp As BITMAP
Dim n As Integer

'Get handle of image
n = ExistsIn(picName, imgName)
If n = 0 Then
    mainCode.ErrorMsg "Image '" & picName & "' does not exist"
    Exit Sub
End If
hBMP = imgHandle.Item(n)

'Fill the bmp type struct
GetObject hBMP, Len(bmp), bmp

'Create a memory DC for the new frame
spriteDC.Add CreateCompatibleDC(0)

'Select the loaded bitmap into the sprite's DC
DeleteObject SelectObject(spriteDC(spriteDC.Count), hBMP)
      
If frame = 0 Then
    frame = 1
    width = bmp.bmWidth
    height = bmp.bmHeight / 2
    initWidth = bmp.bmWidth
    initHeight = bmp.bmHeight / 2
End If

End Sub


Sub Draw()

If drawDC = 0 Then Exit Sub
If spriteDC.Count = 0 Then Exit Sub
If visible = False Then Exit Sub

If display = SD_FLIP Then
    'Transfer flipped mask onto drawDC
    StretchBlt drawDC, left, top + height, width, -height, spriteDC(frame), 0, 0, _
                initWidth, initHeight, vbSrcAnd
    'Transfer flipped sprite onto drawDC
    StretchBlt drawDC, left, top + height, width, -height, spriteDC(frame), 0, initHeight, _
                initWidth, initHeight, vbSrcPaint

ElseIf display = SD_MIRROR Then
    'Transfer mirrored mask onto drawDC
    StretchBlt drawDC, left + width, top, -width, height, spriteDC(frame), 0, 0, _
                initWidth, initHeight, vbSrcAnd
    'Transfer mirrored sprite onto drawDC
    StretchBlt drawDC, left + width, top, -width, height, spriteDC(frame), 0, initHeight, _
                initWidth, initHeight, vbSrcPaint
                
ElseIf display = SD_ROTATE180 Then
    'Transfer rotated mask onto drawDC
    StretchBlt drawDC, left + width, top + height, -width, -height, spriteDC(frame), 0, 0, _
                initWidth, initHeight, vbSrcAnd
    'Transfer rotated sprite onto drawDC
    StretchBlt drawDC, left + width, top + height, -width, -height, spriteDC(frame), 0, initHeight, _
                initWidth, initHeight, vbSrcPaint

Else
  'Transfer the sprite mask section onto drawDC
  StretchBlt drawDC, left, top, width, height, spriteDC(frame), 0, 0, _
              initWidth, initHeight, vbSrcAnd
  'Transfer the sprite image section onto drawDC
  StretchBlt drawDC, left, top, width, height, spriteDC(frame), 0, initHeight, _
              initWidth, initHeight, vbSrcPaint
End If

currTick = GetTickCount()

If isPlaying = True Then
  If currTick - lastTick > rate Then
    If cycle > 0 Then
      If cycleCount < cycle Then
        frame = (frame Mod spriteDC.Count) + 1
        If frame = spriteDC.Count Then cycleCount = cycleCount + 1
      Else
        isPlaying = False
      End If
    Else
      frame = (frame Mod spriteDC.Count) + 1
    End If
      lastTick = GetTickCount()
  End If
End If

End Sub
Sub RemoveFrame(ByVal idx As Integer)

Dim hObj As Long

If idx < 1 Or idx > spriteDC.Count Then Exit Sub

'Delete the DC of the frame to be removed
DeleteDC spriteDC(idx)

spriteDC.Remove idx

If spriteDC.Count = 0 Then
  frame = 0
  playing = False
Else
  frame = 1
End If

End Sub
Private Sub Class_Initialize()

rate = 100
drawDC = 0
isPlaying = True
visible = True
frame = 0
cycle = 0
display = 0
scaling = 100

End Sub


Private Sub Class_Terminate()

If spriteDC.Count > 0 Then
  For i = spriteDC.Count To 1 Step -1
    RemoveFrame i
  Next i
End If

End Sub
