Attribute VB_Name = "RuntimeCode"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private defPath, cmdLineStr As String

'DEBUGGER VARIABLES
Public debugging As Boolean
Public debugState As Integer
Public Const DS_PAUSE = 0
Public Const DS_STEP = 1
Public Const DS_AUTO = 2
'******************

Public lineNum As Integer

Public Const DT_STRING = 1
Public Const DT_NUMBER = 2

Public Const FT_INPUT = 1
Public Const FT_OUTPUT = 2
Public Const FT_APPEND = 3
Public Const FT_BINARY = 4

Public Const SP_SUB = 1
Public Const SP_FUNC = 2

Public varName As New ArrayClass
Public varType As New ArrayClass
Public varValue As New ArrayClass
Public varBindList As New ArrayClass

Public arrayName As New ArrayClass
Public arrayType As New ArrayClass
Public arrayValue As New ArrayClass
Public isMultiDim As New ArrayClass

Public strFunc As New ArrayClass
Public numFunc As New ArrayClass

Public userInput As String
Public inputting As Boolean
Public progDone As Boolean
Public errorFlag As Boolean

Public runCode As New ArrayClass

Public labelName As New ArrayClass
Public labelLine As New ArrayClass

Public gosubLine As Integer

Public fileHandle As New ArrayClass
Public fileNumber As New ArrayClass
Public fileType As New ArrayClass

Public subName As New ArrayClass
Public subParams As New ArrayClass
Public subRunCode As New ArrayClass

Public funcName As New ArrayClass
Public funcType As New ArrayClass
Public funcParams As New ArrayClass
Public funcRunCode As New ArrayClass

Public onErrorCmd As String

Public dataList As New ArrayClass
Public readPos As Integer

Public timerName As New ArrayClass
Public timerID As New ArrayClass
Public timerSubIdx As New ArrayClass
Public timerSubType As New ArrayClass

Public windows As New ArrayClass

Public imgName As New ArrayClass
Public imgHandle As New ArrayClass



Public Sub Cmd_ButtonImg(ByVal cmdStr As String)

Dim winName, picName As String
Dim hPic As Long
Dim n As Integer
Dim params As New ArrayClass
Dim ctlObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
picName = EvalExpression(params.Item(2))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(picName, imgName)

If n = 0 Then
    ErrorMsg "Image '" & nameStr & "' does not exist"
    Exit Sub
Else
    hPic = imgHandle.Item(n)
End If

If ctlObj.ctlType = "picbutton" Then
    SendMessage ctlObj.winHandle, BM_SETIMAGE, 0, ByVal hPic
Else
    ErrorMsg "Control '" & winName & "' needs to be a PICBUTTON"
End If

End Sub


Public Sub Cmd_CloseSound(ByVal cmdStr As String)

Dim soundName As String

cmdStr = Trim(cmdStr)

soundName = EvalExpression(cmdStr)

mciSendString "close " & soundName, "", 0, 0

End Sub


Public Sub Cmd_GetDirs(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim dirList As New ArrayClass
Dim pathName, arrName, fName As String
Dim n As Integer

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

pathName = EvalExpression(params.Item(1))
arrName = params.Item(2)

If Right(arrName, 1) = ")" Then arrName = Trim(left(arrName, Len(arrName) - 1))

fName = Dir(pathName, vbDirectory)
While fName <> ""
    dirList.Add fName
    fName = Dir
Wend

Cmd_ReDim arrName & dirList.itemCount & ")"

SetValue arrName & "0)", dirList.itemCount

For n = 1 To dirList.itemCount
    SetValue arrName & n & ")", dirList.Item(n)
Next n

End Sub


Public Sub Cmd_GetFiles(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim fileList As New ArrayClass
Dim pathName, arrName, fName As String
Dim n As Integer

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

pathName = EvalExpression(params.Item(1))
arrName = params.Item(2)

If Right(arrName, 1) = ")" Then arrName = Trim(left(arrName, Len(arrName) - 1))

fName = Dir(pathName, vbNormal)
While fName <> ""
    fileList.Add fName
    fName = Dir
Wend

Cmd_ReDim arrName & fileList.itemCount & ")"

SetValue arrName & "0)", fileList.itemCount

For n = 1 To fileList.itemCount
    SetValue arrName & n & ")", fileList.Item(n)
Next n

End Sub


Public Sub Cmd_MkDir(ByVal cmdStr As String)

On Error GoTo cmdError

Dim dirPath As String

cmdStr = Trim(cmdStr)

dirPath = EvalExpression(cmdStr)

MkDir dirPath

Exit Sub
cmdError:
    ErrorMsg "Failed to create directory '" & dirPath & "'"

End Sub


Public Sub Cmd_Name(ByVal cmdStr As String)

On Error GoTo cmdError

Dim params As New ArrayClass
Dim oldName, newName As String

cmdStr = Trim(cmdStr)

params.Add GetString(1, LCase(cmdStr), " as ")
params.Item(1) = Mid(cmdStr, 1, Len(params.Item(1)))
params.Add Mid(cmdStr, Len(params.Item(1)) + 5)

oldName = EvalExpression(params.Item(1))
newName = EvalExpression(params.Item(2))

Name oldName As newName

Exit Sub
cmdError:
    ErrorMsg "Failed to rename '" & oldName & "' to '" & newName & "'"

End Sub


Public Sub Cmd_OpenSound(ByVal cmdStr As String)

Dim soundName, soundFile As String
Dim params As New ArrayClass
Dim n As Integer

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

soundName = EvalExpression(params.Item(1))
soundFile = EvalExpression(params.Item(2))

n = GetShortPathName(soundFile, soundFile, Len(soundFile))
soundFile = left(soundFile, n)

mciSendString "open " & soundFile & " alias " & soundName, "", 0, 0

End Sub


Public Sub Cmd_PauseSound(ByVal cmdStr As String)

Dim soundName As String

cmdStr = Trim(cmdStr)

soundName = EvalExpression(cmdStr)

mciSendString "pause " & soundName, "", 0, 0

End Sub



Public Sub Cmd_PlaySound(ByVal cmdStr As String)

Dim soundName As String

cmdStr = Trim(cmdStr)

soundName = EvalExpression(cmdStr)

mciSendString "play " & soundName, "", 0, 0

End Sub



Public Sub Cmd_ResumeSound(ByVal cmdStr As String)

Dim soundName As String

cmdStr = Trim(cmdStr)

soundName = EvalExpression(cmdStr)

mciSendString "resume " & soundName, "", 0, 0

End Sub



Public Sub Cmd_RmDir(ByVal cmdStr As String)

On Error GoTo cmdError

Dim dirPath As String

cmdStr = Trim(cmdStr)

dirPath = EvalExpression(cmdStr)

RmDir dirPath

Exit Sub
cmdError:
    ErrorMsg "Failed to remove directory '" & dirPath & "'"

End Sub


Public Sub Cmd_SetState(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim ctlState As Integer
Dim ctlObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
ctlState = EvalExpression(params.Item(2))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

If ctlObj.ctlType = "radiobutton" Or ctlObj.ctlType = "checkbox" Then
    SendMessage ctlObj.winHandle, BM_SETCHECK, ctlState, 0
Else
    ErrorMsg "Control '" & winName & "' needs to be a RADIOBUTTON or CHECKBOX"
End If


End Sub


Public Sub Cmd_StopSound(ByVal cmdStr As String)

Dim soundName As String

cmdStr = Trim(cmdStr)

soundName = EvalExpression(cmdStr)

mciSendString "stop " & soundName, "", 0, 0

End Sub



Public Sub Cmd_AddItem(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim itemText As String
Dim itemIdx As Integer
Dim ctlObj As Object

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
itemText = EvalExpression(params.Item(2))
If params.itemCount = 3 Then itemIdx = EvalExpression(params.Item(3))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

If ctlObj.ctlType = "listbox" Then
    SendMessage ctlObj.winHandle, LB_INSERTSTRING, itemIdx - 1, ByVal itemText
ElseIf ctlObj.ctlType = "combobox" Then
    SendMessage ctlObj.winHandle, CB_INSERTSTRING, itemIdx - 1, ByVal itemText
Else
    ErrorMsg "Control '" & winName & "' needs to be a COMBOBOX or LISTBOX"
End If

End Sub

Public Sub Cmd_DelItem(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim itemIdx As Integer
Dim ctlObj As Object

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
itemIdx = EvalExpression(params.Item(2))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

If ctlObj.ctlType = "listbox" Then
    SendMessage ctlObj.winHandle, LB_DELETESTRING, itemIdx - 1, 0
ElseIf ctlObj.ctlType = "combobox" Then
    SendMessage ctlObj.winHandle, CB_DELETESTRING, itemIdx - 1, 0
Else
    ErrorMsg "Control '" & winName & "' needs to be a COMBOBOX or LISTBOX"
End If

End Sub

Public Sub Cmd_Disable(ByVal cmdStr As String)

Dim winName As String
Dim winHandle As Long

cmdStr = Trim(cmdStr)

winName = EvalExpression(cmdStr)

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

EnableWindow winHandle, False

End Sub

Public Sub Cmd_DrawText(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim left, top As Integer
Dim drawText As Variant
Dim winDC As Long

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
drawText = EvalExpression(params.Item(2))
left = EvalExpression(params.Item(3))
top = EvalExpression(params.Item(4))

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

TextOut winDC, left, top, drawText, Len(drawText)

End Sub

Public Sub Cmd_Enable(ByVal cmdStr As String)

Dim winName As String
Dim winHandle As Long

cmdStr = Trim(cmdStr)

winName = EvalExpression(cmdStr)

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

EnableWindow winHandle, True

End Sub

Public Sub Cmd_GetSel(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, startVar, endVar As String
Dim startPos, endPos, selected As Long
Dim ctlObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
startVar = params.Item(2)
endVar = params.Item(3)

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

If ctlObj.ctlType = "textbox" Or ctlObj.ctlType = "texteditor" Then
    selected = SendMessage(ctlObj.winHandle, EM_GETSEL, 0, 0)
    startPos = LOWORD(selected) + 1
    endPos = HIWORD(selected) + 1
    SetValue startVar, startPos
    SetValue endVar, endPos
Else
    ErrorMsg "Control '" & winName & "' needs to be a TEXTBOX or TEXTEDITOR"
End If

End Sub


Public Sub Cmd_Line(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim startX, startY As Integer
Dim endX, endY As Integer
Dim winDC As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
startX = EvalExpression(params.Item(2))
startY = EvalExpression(params.Item(3))
endX = EvalExpression(params.Item(4))
endY = EvalExpression(params.Item(5))

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

MoveToEx winDC, startX, startY, vbNull
LineTo winDC, endX, endY

End Sub


Public Sub Cmd_Circle(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim centerX, centerY, radius As Integer
Dim left, top, Right, Bottom As Integer
Dim winDC As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
centerX = EvalExpression(params.Item(2))
centerY = EvalExpression(params.Item(3))
radius = EvalExpression(params.Item(4))

left = centerX - radius
top = centerY - radius
Right = centerX + radius
Bottom = centerY + radius

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

Ellipse winDC, left, top, Right, Bottom

End Sub


Public Sub Cmd_Box(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim left, top, width, height As Integer
Dim Right, Bottom As Integer
Dim winDC As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
left = EvalExpression(params.Item(2))
top = EvalExpression(params.Item(3))
width = EvalExpression(params.Item(4))
height = EvalExpression(params.Item(5))

Right = left + width
Bottom = top + height

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

Rectangle winDC, left, top, Right, Bottom

End Sub


Public Sub Cmd_LineSize(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim lineSize As Integer
Dim winDC As Long
Dim pen As LOGPEN

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
lineSize = EvalExpression(params.Item(2))

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

GetObject GetCurrentObject(winDC, OBJ_PEN), Len(pen), pen
DeleteObject SelectObject(winDC, CreatePen(PS_SOLID, lineSize, pen.lopnColor))

End Sub


Public Sub Cmd_LoadImg(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim nameStr As String
Dim fileStr As String
Dim hImg As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

nameStr = EvalExpression(params.Item(1))
fileStr = EvalExpression(params.Item(2))

hImg = LoadImage(fileStr)

If hImg = 0 Then
    ErrorMsg "Unable to load bitmap '" & fileStr & "'"
    Exit Sub
End If

imgName.Add nameStr
imgHandle.Add hImg

End Sub


Public Sub Cmd_SetFont(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, fontName As String
Dim fontWidth, fontHeight As Integer
Dim winDC As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
fontName = EvalExpression(params.Item(2))
fontWidth = EvalExpression(params.Item(3))
fontHeight = EvalExpression(params.Item(4))

winDC = GetWinDC(winName)

DeleteObject GetCurrentObject(winDC, OBJ_FONT)

SelectObject winDC, CreateFont(fontHeight, fontWidth, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, fontName)

End Sub

Public Sub Cmd_Sprite(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr, imgStr As String
Dim left, top, n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
left = EvalExpression(params.Item(3))
top = EvalExpression(params.Item(4))
imgStr = EvalExpression(params.Item(5))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n Then
    ErrorMsg "Sprite '" & nameStr & "' already exists in window or control '" & winName & "'"
    Exit Sub
End If

winObj.spriteName.Add nameStr
winObj.sprites.Add New SpriteClass

winObj.sprites.Item(winObj.sprites.itemCount).left = left
winObj.sprites.Item(winObj.sprites.itemCount).top = top
winObj.sprites.Item(winObj.sprites.itemCount).AddFrame imgStr

End Sub


Public Sub Cmd_DelSprite(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.spriteName.Remove n
winObj.sprites.Remove n

End Sub


Public Sub Cmd_DrawSprites(ByVal cmdStr As String)

Dim winName As String
Dim winObj As Object
Dim width, height, n, a, b As Integer
Dim backDC, backBMP As Long
Dim hBrush, bgDC As Long
Dim winRect As RECT
Dim bmpInfo As BITMAP
Dim cutX, cutY, cutWidth, cutHeight As Integer

cmdStr = Trim(cmdStr)

winName = EvalExpression(cmdStr)

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

GetWindowRect winObj.winHandle, winRect
width = winRect.Right - winRect.left
height = winRect.Bottom - winRect.top

backDC = CreateCompatibleDC(winObj.winDC)
backBMP = CreateCompatibleBitmap(winObj.winDC, width, height)
SelectObject backDC, backBMP

If winObj.spriteBG = 0 Then
    winRect.left = 0: winRect.top = 0
    winRect.Right = width: winRect.Bottom = height
    hBrush = CreateSolidBrush(vbWhite)
    FillRect backDC, winRect, hBrush
    DeleteObject hBrush
Else
    GetObject winObj.spriteBG, Len(bmpInfo), bmpInfo
    bgDC = CreateCompatibleDC(0)
    SelectObject bgDC, winObj.spriteBG
    For a = 0 To Int(height / bmpInfo.bmHeight)
        For b = 0 To Int(width / bmpInfo.bmWidth)
            BitBlt backDC, (b * bmpInfo.bmWidth), (a * bmpInfo.bmHeight), _
                    bmpInfo.bmWidth, bmpInfo.bmHeight, bgDC, 0, 0, vbSrcCopy
        Next b
    Next a
    DeleteDC bgDC
End If

For n = 1 To winObj.sprites.itemCount
    winObj.sprites.Item(n).drawDC = backDC
    winObj.sprites.Item(n).Draw
Next n

BitBlt winObj.winDC, 0, 0, width, height, backDC, 0, 0, vbSrcCopy

DeleteDC backDC
DeleteObject backBMP

End Sub


Public Sub Cmd_AddFrame(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr, imgStr As String
Dim n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
imgStr = EvalExpression(params.Item(3))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).AddFrame imgStr

End Sub


Public Sub Cmd_DelFrame(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim frameIdx, n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
frameIdx = EvalExpression(params.Item(3))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).RemoveFrame frameIdx

End Sub


Public Sub Cmd_SpriteBG(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, imgStr As String
Dim n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
imgStr = EvalExpression(params.Item(2))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

If Trim(imgStr) = "" Then
    winObj.spriteBG = 0
    Exit Sub
End If

n = ExistsIn(imgStr, imgName)
If n = 0 Then
    ErrorMsg "Image '" & imgStr & "' does not exist"
    Exit Sub
End If

winObj.spriteBG = imgHandle.Item(n)

End Sub

Public Sub Cmd_SpritePause(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).isPlaying = False

End Sub


Public Sub Cmd_SpritePlay(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).isPlaying = True

End Sub


Public Sub Cmd_SpritePos(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim left, top, n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
left = EvalExpression(params.Item(3))
top = EvalExpression(params.Item(4))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).left = left
winObj.sprites.Item(n).top = top

End Sub


Public Sub Cmd_SpriteSize(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim width, height, n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
width = EvalExpression(params.Item(3))
height = EvalExpression(params.Item(4))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).width = width
winObj.sprites.Item(n).height = height

End Sub


Public Sub Cmd_SpriteRate(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim rate, n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
rate = EvalExpression(params.Item(3))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).rate = rate

End Sub


Public Sub Cmd_SpriteRotate(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr, rotation As String
Dim n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
rotation = LCase(EvalExpression(params.Item(3)))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

Select Case rotation
    Case "normal"
        winObj.sprites.Item(n).display = 0
    Case "flip"
        winObj.sprites.Item(n).display = 1
    Case "mirror"
        winObj.sprites.Item(n).display = 2
    Case "rotate180"
        winObj.sprites.Item(n).display = 3
End Select

End Sub


Public Sub Cmd_SpriteShow(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).Visible = True

End Sub


Public Sub Cmd_SpriteHide(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim n As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, winObj.spriteName)
If n = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Sub
End If

winObj.sprites.Item(n).Visible = False

End Sub


Public Sub Cmd_UnloadImg(ByVal cmdStr As String)

Dim nameStr As String
Dim n As Integer

cmdStr = Trim(cmdStr)

nameStr = EvalExpression(cmdStr)

n = ExistsIn(nameStr, imgName)

If n = 0 Then
    ErrorMsg "Image '" & nameStr & "' does not exist"
    Exit Sub
End If

DeleteObject imgHandle.Item(n)

imgName.Remove n
imgHandle.Remove n

End Sub


Public Sub Cmd_DrawImg(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim nameStr As String
Dim left, top, n As Integer
Dim winDC, tempDC As Long
Dim bmpInfo As BITMAP

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
left = EvalExpression(params.Item(3))
top = EvalExpression(params.Item(4))

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

n = ExistsIn(nameStr, imgName)

If n = 0 Then
    ErrorMsg "Image '" & nameStr & "' does not exist"
    Exit Sub
End If

GetObject imgHandle.Item(n), Len(bmpInfo), bmpInfo
tempDC = CreateCompatibleDC(0)
SelectObject tempDC, imgHandle.Item(n)
BitBlt winDC, left, top, bmpInfo.bmWidth, bmpInfo.bmHeight, tempDC, 0, 0, vbSrcCopy
DeleteDC tempDC

End Sub


Public Sub Cmd_GetImg(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, nameStr As String
Dim left, top, width, height As Integer
Dim winDC As Long
Dim tempDC, tempBMP As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
nameStr = EvalExpression(params.Item(2))
left = EvalExpression(params.Item(3))
top = EvalExpression(params.Item(4))
width = EvalExpression(params.Item(5))
height = EvalExpression(params.Item(6))

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

tempDC = CreateCompatibleDC(winDC)
tempBMP = CreateCompatibleBitmap(winDC, width, height)
SelectObject tempDC, tempBMP
BitBlt tempDC, 0, 0, width, height, winDC, left, top, vbSrcCopy
DeleteDC tempDC

imgName.Add nameStr
imgHandle.Add tempBMP

End Sub


Public Sub Cmd_Stick(ByVal cmdStr As String)

Dim winName As String
Dim winDC, redrawDC, winHandle, tmpBmp As Long
Dim winRect As RECT
Dim width, height As Integer

cmdStr = Trim(cmdStr)

winName = EvalExpression(cmdStr)

winHandle = GetWinHandle(winName)
If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If
winDC = GetWinDC(winName)
redrawDC = GetWinRedrawDC(winName)

GetWindowRect winHandle, winRect
width = winRect.Right - winRect.left
height = winRect.Bottom - winRect.top

tmpBmp = CreateCompatibleBitmap(winDC, width, height)
DeleteObject SelectObject(redrawDC, tmpBmp)
BitBlt redrawDC, 0, 0, width, height, winDC, 0, 0, vbSrcCopy

End Sub


Public Sub Cmd_Refresh(ByVal cmdStr As String)

Dim winName As String
Dim winHandle, winDC, redrawDC As Long

cmdStr = Trim(cmdStr)

winName = EvalExpression(cmdStr)

winHandle = GetWinHandle(winName)
If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If
winDC = GetWinDC(winName)
redrawDC = GetWinRedrawDC(winName)

SendMessage winHandle, WM_ERASEBKGND, winDC, ByVal 0
RedrawWindow winHandle, ByVal 0&, 0&, RDW_INVALIDATE

End Sub


Public Sub Cmd_Clear(ByVal cmdStr As String)

Dim winName As String
Dim winObj As Object

cmdStr = Trim(cmdStr)

winName = EvalExpression(cmdStr)

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

DeleteObject GetCurrentObject(winObj.redrawDC, OBJ_BITMAP)
DeleteDC winObj.redrawDC
winObj.redrawDC = CreateCompatibleDC(winObj.winDC)
SendMessage winObj.winHandle, WM_ERASEBKGND, winObj.winDC, ByVal 0
RedrawWindow winObj.winHandle, ByVal 0&, 0&, RDW_INVALIDATE

End Sub


Public Sub Cmd_BackColor(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim color As Long
Dim winDC As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
color = EvalExpression(params.Item(2))

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

DeleteObject SelectObject(winDC, CreateSolidBrush(color))
SetBkColor winDC, color

End Sub


Public Sub Cmd_ForeColor(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim color As Long
Dim winDC As Long
Dim pen As LOGPEN

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
color = EvalExpression(params.Item(2))

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

GetObject GetCurrentObject(winDC, OBJ_PEN), Len(pen), pen
DeleteObject SelectObject(winDC, CreatePen(PS_SOLID, pen.lopnWidth.X, color))
SetTextColor winDC, color

End Sub


Public Sub Cmd_SetPixel(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim X, Y As Integer
Dim color As Long
Dim winDC As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
X = EvalExpression(params.Item(2))
Y = EvalExpression(params.Item(3))
color = EvalExpression(params.Item(4))

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Sub
End If

SetPixel winDC, X, Y, color

End Sub


Public Sub Cmd_SetClipboardText(ByVal cmdStr As String)

Dim cbText As String

cmdStr = Trim(cmdStr)

cbText = EvalExpression(cmdStr)

Clipboard.SetText cbText

End Sub


Public Sub Cmd_SetSel(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim startPos As Long
Dim endPos As Long
Dim ctlObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
startPos = EvalExpression(params.Item(2))
endPos = EvalExpression(params.Item(3))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

If ctlObj.ctlType = "textbox" Or ctlObj.ctlType = "texteditor" Then
    SendMessage ctlObj.winHandle, EM_SETSEL, startPos - 1, ByVal endPos - 1
Else
    ErrorMsg "Control '" & winName & "' needs to be a TEXTBOX or TEXTEDITOR"
End If

End Sub


Public Sub Cmd_Hide(ByVal cmdStr As String)

Dim winHandle As Long
Dim winName As String

cmdStr = Trim(cmdStr)

winName = EvalExpression(cmdStr)

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

ShowWindow winHandle, SW_HIDE

End Sub

Public Sub Cmd_Menu(ByVal cmdStr As String)

Dim winName, mText As String
Dim subStr, miText As String
Dim n, a, b, C, mCount As Integer
Dim miStyle As Long
Dim miID As Long
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
mText = EvalExpression(params.Item(2))

For n = 1 To windows.itemCount
    If winName = windows.Item(n).winName Then
    With windows.Item(n)
        If .hMenu.itemCount = 0 Then
            .hMenuBar = CreateMenu()
            SetMenu .winHandle, .hMenuBar
        End If
        .hMenu.Add CreatePopupMenu()
        .menuItemID.Add New ArrayClass
        .menuItemSubIdx.Add New ArrayClass
        .menuItemSubType.Add New ArrayClass
        AppendMenu .hMenuBar, MF_STRING Or MF_POPUP, .hMenu.Item(.hMenu.itemCount), mText
        DrawMenuBar .winHandle
        If params.itemCount = 2 Then Exit Sub
        mCount = .hMenu.itemCount
        a = 2
        While a < params.itemCount
            a = a + 1
            If params.Item(a) = "|" Then
                miStyle = MF_STRING Or MF_SEPARATOR
            Else
                miText = EvalExpression(params.Item(a))
                subStr = params.Item(a + 1)
                miStyle = MF_STRING
                miID = mCount & .menuItemID.Item(mCount).itemCount + 1
                .menuItemID.Item(mCount).Add miID
                b = ExistsIn(subStr, subName)
                If b Then
                    .menuItemSubIdx.Item(mCount).Add b
                    .menuItemSubType.Item(mCount).Add SP_SUB
                Else
                    C = ExistsIn(subStr, funcName)
                    If C Then
                        .menuItemSubIdx.Item(mCount).Add C
                        .menuItemSubType.Item(mCount).Add SP_FUNC
                    End If
                End If
                a = a + 1
            End If
            AppendMenu .hMenu.Item(mCount), miStyle, miID, miText
            DrawMenuBar .winHandle
        Wend
        Exit Sub
    End With
    End If
Next n

ErrorMsg "Window '" & winName & "' does exist"

End Sub

Public Sub Cmd_SetItem(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim itemText As String
Dim itemIdx As Integer
Dim ctlObj As Object

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
itemIdx = EvalExpression(params.Item(2))
itemText = EvalExpression(params.Item(3))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

If ctlObj.ctlType = "listbox" Then
    SendMessage ctlObj.winHandle, LB_DELETESTRING, itemIdx - 1, 0
    SendMessage ctlObj.winHandle, LB_INSERTSTRING, itemIdx - 1, ByVal itemText
ElseIf ctlObj.ctlType = "combobox" Then
    SendMessage ctlObj.winHandle, CB_DELETESTRING, itemIdx - 1, 0
    SendMessage ctlObj.winHandle, CB_INSERTSTRING, itemIdx - 1, ByVal itemText
Else
    ErrorMsg "Control '" & winName & "' needs to be a COMBOBOX or LISTBOX"
End If

End Sub

Public Sub Cmd_SetSelIdx(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim itemIdx As Integer
Dim ctlObj As Object

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
itemIdx = EvalExpression(params.Item(2))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

If ctlObj.ctlType = "listbox" Then
    SendMessage ctlObj.winHandle, LB_SETCURSEL, itemIdx - 1, 0
ElseIf ctlObj.ctlType = "combobox" Then
    SendMessage ctlObj.winHandle, CB_SETCURSEL, itemIdx - 1, 0
Else
    ErrorMsg "Control '" & winName & "' needs to be a COMBOBOX or LISTBOX"
End If

End Sub



Public Sub Cmd_SetSelText(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim repStr As String
Dim ctlObj As Object

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
repStr = EvalExpression(params.Item(2))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Sub
End If

If ctlObj.ctlType = "textbox" Or ctlObj.ctlType = "texteditor" Then
    SendMessage ctlObj.winHandle, EM_REPLACESEL, 1, ByVal repStr
Else
    ErrorMsg "Control '" & winName & "' needs to be a TEXTBOX or TEXTEDITOR"
End If

End Sub


Public Sub Cmd_Show(ByVal cmdStr As String)

Dim winHandle As Long
Dim winName As String
Dim a, b, n As Integer

cmdStr = Trim(cmdStr)

winName = EvalExpression(cmdStr)

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

ShowWindow winHandle, SW_SHOW

End Sub





Public Sub DebugUpdateArrays(Optional projObj As Object)

End Sub

Public Sub DebugUpdateCode(Optional progObj As Object)

debugWin.code.Clear

If progObj Is Nothing Then
    For n = 1 To runCode.itemCount
        debugWin.code.AddItem runCode.Item(n)
    Next n
    debugWin.code.AddItem ""
Else
    For n = 1 To progObj.subProg_runCode.itemCount
        debugWin.code.AddItem progObj.subProg_runCode.Item(n)
    Next n
    debugWin.code.AddItem ""
End If

End Sub


Public Sub DebugUpdateVars(Optional progObj As Object)

Dim n As Integer

debugWin.localVars.Clear

If progObj Is Nothing Then
    debugWin.globalVars.Clear
    For n = 1 To varName.itemCount
        If varType.Item(n) = DT_STRING Then
            debugWin.globalVars.AddItem varName.Item(n) & "  =  " & Chr(34) & varValue.Item(n) & Chr(34)
        Else
            debugWin.globalVars.AddItem varName.Item(n) & "  =  " & varValue.Item(n)
        End If
    Next n
Else
    With progObj
    For n = 1 To .subProg_varName.itemCount
        If .subProg_varType.Item(n) = DT_STRING Then
            debugWin.localVars.AddItem .subProg_varName.Item(n) & "  =  " & Chr(34) & .subProg_varValue.Item(n) & Chr(34)
        Else
            debugWin.localVars.AddItem .subProg_varName.Item(n) & "  =  " & .subProg_varValue.Item(n)
        End If
    Next n
    End With
End If

End Sub



Public Sub DefineSysVars()

Cmd_Var "ErrorMsg as string"
Cmd_Var "Red as number": SetValue "Red", vbRed
Cmd_Var "Yellow as number": SetValue "Yellow", vbYellow
Cmd_Var "Orange as number": SetValue "Orange", CLng(&H80FF&)
Cmd_Var "Blue as number": SetValue "Blue", vbBlue
Cmd_Var "Green as number": SetValue "Green", vbGreen
Cmd_Var "Purple as number": SetValue "Purple", CLng(&H800080)
Cmd_Var "Black as number": SetValue "Black", vbBlack
Cmd_Var "White as number": SetValue "White", vbWhite
Cmd_Var "Brown as number": SetValue "Brown", CLng(&H45371)
Cmd_Var "ButtonFace as number": SetValue "ButtonFace", GetSysColor(COLOR_BTNFACE)
Cmd_Var "DefPath as string": SetValue "DefPath", App.Path
Cmd_Var "CommandLine as string": SetValue "CommandLine", Command
Cmd_Var "False as number": SetValue "False", 0
Cmd_Var "True as number": SetValue "True", 1
Cmd_Var "ScreenWidth as number": SetValue "ScreenWidth", (Screen.width / Screen.TwipsPerPixelX)
Cmd_Var "ScreenHeight as number": SetValue "ScreenHeight", (Screen.height / Screen.TwipsPerPixelY)
Cmd_Var "Inkey as string"

End Sub

Public Function FileDialog(ByVal titleStr As String, ByVal filterStr As String, ByVal dialogType As Integer) As String

'0 = Open
'1 = Save

Dim fdInfo As OPENFILENAME
Dim n As Integer

fdInfo.lStructSize = Len(fdInfo)
fdInfo.flags = OFN_EXPLORER Or OFN_NODEREFERENCELINKS Or _
                OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
fdInfo.nMaxFile = 260
fdInfo.lpstrFile = Space(260)
fdInfo.lpstrFilter = Replace(filterStr, "|", Chr(0)) & Chr(0)
fdInfo.lpstrTitle = titleStr

If dialogType Then
    GetSaveFileName fdInfo
Else
    GetOpenFileName fdInfo
End If

FileDialog = fdInfo.lpstrFile

End Function



Public Function Func_Collide(ByVal paramStr As String) As Variant

Dim params As New ArrayClass
Dim winName, sprite1, sprite2 As String
Dim a, b As Integer
Dim winObj As Object

ParseParams paramStr, params

winName = EvalExpression(params.Item(1))
sprite1 = EvalExpression(params.Item(2))
sprite2 = EvalExpression(params.Item(3))

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then
    ErrorMsg "Window or control '" & winName & "' does not exist"
    Exit Function
End If

a = ExistsIn(sprite1, winObj.spriteName)
If a = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Function
End If
b = ExistsIn(sprite2, winObj.spriteName)
If b = 0 Then
    ErrorMsg "Sprite '" & nameStr & "' does not exist in window or control '" & winName & "'"
    Exit Function
End If

With winObj.sprites
    If .Item(a).left < .Item(b).left + .Item(b).width And _
       .Item(a).left + .Item(a).width > .Item(b).left And _
       .Item(a).top < .Item(b).top + .Item(b).height And _
       .Item(a).top + .Item(a).height > .Item(b).top Then
          Func_Collide = 1
    End If
End With

End Function


Public Function Func_Date() As Variant

Func_Date = Date

End Function


Public Function Func_EOF(ByVal paramStr As String) As Variant

Dim a As Integer

paramStr = Trim(paramStr)

a = ExistsIn(paramStr, fileHandle)

If a Then
    Func_EOF = Abs(EOF(fileNumber.Item(a)))
Else
    ErrorMsg "File handle does not exist: " & paramStr
End If

End Function


Public Function Func_FileOpen(ByVal paramStr As String) As Variant

Dim titleStr, filterStr, filename As String
Dim params As New ArrayClass

paramStr = Trim(paramStr)

ParseParams paramStr, params

titleStr = EvalExpression(params.Item(1))
filterStr = EvalExpression(params.Item(2))

filename = FileDialog(titleStr, filterStr, 0)

Func_FileOpen = Mid(filename, 1, InStr(1, filename, Chr(0)) - 1)

End Function


Public Function Func_FileSave(ByVal paramStr As String) As Variant

Dim titleStr, filterStr, filename As String
Dim params As New ArrayClass

paramStr = Trim(paramStr)

ParseParams paramStr, params

titleStr = EvalExpression(params.Item(1))
filterStr = EvalExpression(params.Item(2))

filename = FileDialog(titleStr, filterStr, 1)

Func_FileSave = Mid(filename, 1, InStr(1, filename, Chr(0)) - 1)

End Function


Public Function Func_hBmp(ByVal paramStr As String) As Variant

Dim picName As String
Dim n As Integer

paramStr = Trim(paramStr)

picName = EvalExpression(paramStr)

n = ExistsIn(picName, imgName)

If n Then
    Func_hBmp = imgHandle.Item(n)
Else
    ErrorMsg "Image '" & picName & "' does not exist"
End If

End Function


Public Function Func_Input(ByVal paramStr As String) As Variant

Dim params As New ArrayClass
Dim lenVal As Long
Dim a As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

a = ExistsIn(params.Item(1), fileHandle)
lenVal = EvalExpression(params.Item(2))

If a Then
    Select Case fileType.Item(a)
        Case FT_OUTPUT
            ErrorMsg "File opened in OUTPUT mode cannot be inputted from"
        Case FT_APPEND
            ErrorMsg "File opened in APPEND mode cannot be inputted from"
    End Select
    Func_Input = Input(lenVal, fileNumber.Item(a))
Else
    ErrorMsg "File handle does not exist: " & params.Item(1)
End If

End Function

Public Function Func_LOF(ByVal paramStr As String) As Variant

Dim a As Integer

paramStr = Trim(paramStr)

a = ExistsIn(paramStr, fileHandle)

If a Then
    Func_LOF = LOF(fileNumber.Item(a))
Else
    ErrorMsg "File handle does not exist: " & paramStr
End If

End Function


Public Function Func_Time() As Variant

Func_Time = Time

End Function


Public Function Func_GetClipboardText() As Variant

Func_GetClipboardText = Clipboard.GetText()

End Function

Public Function Func_GetLineText(ByVal paramStr As String) As Variant

Dim winName, lineText As String
Dim lLen, lNum As Long
Dim params As New ArrayClass
Dim ctlObj As Object

paramStr = Trim(paramStr)

ParseParams paramStr, params

winName = EvalExpression(params.Item(1))
lNum = EvalExpression(params.Item(2))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Function
End If

If ctlObj.ctlType = "textbox" Or ctlObj.ctlType = "texteditor" Then
    lineText = Space(255)
    lLen = SendMessage(ctlObj.winHandle, EM_GETLINE, lNum - 1, ByVal lineText)
    lineText = left(lineText, lLen)
    Func_GetLineText = lineText
Else
    ErrorMsg "Control '" & winName & "' needs to be a TEXTBOX or TEXTEDITOR"
End If

End Function

Public Function Func_GetSelIdx(ByVal paramStr As String) As Variant

Dim winName As String
Dim itemIdx As Integer
Dim ctlObj As Object

paramStr = Trim(paramStr)

winName = EvalExpression(paramStr)

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Function
End If

If ctlObj.ctlType = "listbox" Then
    itemIdx = SendMessage(ctlObj.winHandle, LB_GETCURSEL, 0, 0)
    Func_GetSelIdx = itemIdx + 1
ElseIf ctlObj.ctlType = "combobox" Then
    itemIdx = SendMessage(ctlObj.winHandle, CB_GETCURSEL, 0, 0)
    Func_GetSelIdx = itemIdx + 1
Else
    ErrorMsg "Control '" & winName & "' needs to be a COMBOBOX or LISTBOX"
End If

End Function


Public Function Func_GetItem(ByVal paramStr As String) As Variant

Dim winName, itemText As String
Dim itemIdx As Integer
Dim itemLen As Long
Dim params As New ArrayClass
Dim ctlObj As Object

paramStr = Trim(paramStr)

ParseParams paramStr, params

winName = EvalExpression(params.Item(1))
itemIdx = EvalExpression(params.Item(2))

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Function
End If

If ctlObj.ctlType = "listbox" Then
    itemLen = SendMessage(ctlObj.winHandle, LB_GETTEXTLEN, itemIdx - 1, 0)
    itemText = Space(Abs(itemLen))
    SendMessage ctlObj.winHandle, LB_GETTEXT, itemIdx - 1, ByVal itemText
    Func_GetItem = itemText
ElseIf ctlObj.ctlType = "combobox" Then
    itemLen = SendMessage(ctlObj.winHandle, CB_GETLBTEXTLEN, itemIdx - 1, 0)
    itemText = Space(Abs(itemLen))
    SendMessage ctlObj.winHandle, CB_GETLBTEXT, itemIdx - 1, ByVal itemText
    Func_GetItem = itemText
Else
    ErrorMsg "Control '" & winName & "' needs to be a COMBOBOX or LISTBOX"
End If

End Function

Public Function Func_GetSelText(ByVal paramStr As String) As Variant

Dim winName, selText As String
Dim winText As String
Dim startPos, endPos As Integer
Dim textLen, selected As Long
Dim ctlObj As Object

paramStr = Trim(paramStr)

winName = EvalExpression(paramStr)

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Function
End If

If ctlObj.ctlType = "textbox" Or ctlObj.ctlType = "texteditor" Then
    textLen = GetWindowTextLength(ctlObj.winHandle)
    winText = Space(textLen + 1)
    GetWindowText ctlObj.winHandle, winText, textLen + 1
    selected = SendMessage(ctlObj.winHandle, EM_GETSEL, 0, 0)
    startPos = LOWORD(selected) + 1
    endPos = HIWORD(selected) + 1
    selText = Mid(winText, startPos, endPos - startPos)
    Func_GetSelText = selText
Else
    ErrorMsg "Control '" & winName & "' needs to be a TEXTBOX or TEXTEDITOR"
End If

End Function

Public Function Func_InputBox(ByVal paramStr As String) As Variant

Dim params As New ArrayClass
Dim promptStr, titleStr As String
Dim defVal As Variant

paramStr = Trim(paramStr)

ParseParams paramStr, params

promptStr = EvalExpression(params.Item(1))
titleStr = EvalExpression(params.Item(2))
If params.itemCount = 3 Then defVal = EvalExpression(params.Item(3))

inputWin.prompt.Caption = promptStr
inputWin.Caption = titleStr
inputWin.inputVal.Text = defVal
inputWin.Show vbModal

Func_InputBox = userInput
userInput = ""

End Function

Public Function Func_ItemCount(ByVal paramStr As String) As Variant

Dim winName As String
Dim itemNum As Integer
Dim ctlObj As Object

paramStr = Trim(paramStr)

winName = EvalExpression(paramStr)

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Function
End If
        
If ctlObj.ctlType = "listbox" Then
    itemNum = SendMessage(ctlObj.winHandle, LB_GETCOUNT, 0, 0)
    Func_ItemCount = itemNum
ElseIf ctlObj.ctlType = "combobox" Then
    itemNum = SendMessage(ctlObj.winHandle, CB_GETCOUNT, 0, 0)
    Func_ItemCount = itemNum
Else
    ErrorMsg "Control '" & winName & "' needs to be a COMBOBOX or LISTBOX"
End If

End Function



Public Function Func_LineCount(ByVal paramStr As String) As Variant

Dim winName As String
Dim lNum As Long
Dim ctlObj As Object

paramStr = Trim(paramStr)

winName = EvalExpression(paramStr)

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Function
End If

If ctlObj.ctlType = "textbox" Or ctlObj.ctlType = "texteditor" Then
    lNum = SendMessage(ctlObj.winHandle, EM_GETLINECOUNT, 0, 0)
    Func_LineCount = lNum
Else
    ErrorMsg "Control '" & winName & "' needs to be a TEXTBOX or TEXTEDITOR"
End If

End Function


Public Function GetControlObj(ByVal ctlName As String) As Object

'Get the object of the given control, returning
'an empty object if the control name was not found

Dim a, b As Integer

Set GetControlObj = Nothing

For a = 1 To windows.itemCount
    With windows.Item(a)
    For b = 1 To .Controls.itemCount
        If ctlName = .Controls.Item(b).winName Then
            Set GetControlObj = .Controls.Item(b)
            Exit Function
        End If
    Next b
    End With
Next a

End Function

Public Function Func_GetState(ByVal paramStr As String)

Dim winName As String
Dim ctlState As Integer
Dim ctlObj As Object

paramStr = Trim(paramStr)

winName = EvalExpression(paramStr)

Set ctlObj = GetControlObj(winName)

If ctlObj Is Nothing Then
    ErrorMsg "Control '" & winName & "' does not exist"
    Exit Function
End If

If ctlObj.ctlType = "radiobutton" Or ctlObj.ctlType = "checkbox" Then
    ctlState = SendMessage(ctlObj.winHandle, BM_GETCHECK, 0, 0)
    Func_GetState = ctlState
Else
    ErrorMsg "Control '" & winName & "' needs to be a RADIOBUTTON or CHECKBOX"
End If

End Function

Public Function GetWindowObj(ByVal winName As String) As Object

'Get the object of the given window or control, returning
'an empty object if the window/control name was not found

Dim a As Integer

Set GetWindowObj = Nothing

For a = 1 To windows.itemCount
    With windows.Item(a)
    If winName = windows.Item(a).winName Then
        Set GetWindowObj = windows.Item(a)
        Exit Function
    Else
        For b = 1 To .Controls.itemCount
            If winName = .Controls.Item(b).winName Then
                Set GetWindowObj = .Controls.Item(b)
                Exit Function
            End If
        Next b
    End If
    End With
Next a

End Function

Public Function GetWinHandle(ByVal winName As String) As Long

Dim winObj As Object

GetWinHandle = 0

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then Exit Function

GetWinHandle = winObj.winHandle

End Function

Public Function GetWinDC(ByVal winName As String) As Long

Dim winObj As Object

GetWinDC = 0

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then Exit Function

GetWinDC = winObj.winDC

End Function


Public Function GetWinRedrawDC(ByVal winName As String) As Long

Dim winObj As Object

GetWinRedrawDC = 0

Set winObj = GetWindowObj(winName)

If winObj Is Nothing Then Exit Function

GetWinRedrawDC = winObj.redrawDC

End Function
Public Function LoadImage(ByVal filename As String) As Long

On Error GoTo loadError

Dim hBMP As Long
Dim pic As IPictureDisp
Dim hDC1, hDC2 As Long
Dim width, height As Integer
Dim bmpInfo As BITMAP

Set pic = LoadPicture(filename)
hDC1 = CreateCompatibleDC(0)
DeleteObject SelectObject(hDC1, pic.handle)

GetObject pic.handle, Len(bmpInfo), bmpInfo
width = bmpInfo.bmWidth
height = bmpInfo.bmHeight

hDC2 = CreateCompatibleDC(hDC1)
hBMP = CreateCompatibleBitmap(hDC1, width, height)
DeleteObject SelectObject(hDC2, hBMP)

BitBlt hDC2, 0, 0, width, height, hDC1, 0, 0, vbSrcCopy

DeleteDC hDC1
DeleteDC hDC2

Set pic = Nothing

LoadImage = hBMP
Exit Function

loadError:

End Function

Public Function Min(ByVal val1 As Long, ByVal val2 As Long) As Long

If val1 < val2 Then
    Min = val1
Else
    Min = val2
End If

End Function



Public Function Max(ByVal val1 As Long, ByVal val2 As Long) As Long

If val1 > val2 Then
    Max = val1
Else
    Max = val2
End If

End Function



Public Sub ParseParams(ByVal paramStr As String, ByRef paramList As ArrayClass)

Dim b As Integer
Dim tmpParam As String

b = 1
While b <= Len(paramStr)
    tmpParam = GetString(b, paramStr, ",")
    b = Len(tmpParam) + b + 1
    paramList.Add Trim(tmpParam)
Wend

End Sub

Public Sub CallSubProg(ByVal subIdx As Integer, ByVal subType As Integer)

'This sub makes a quick call to a sub program, without
'any arguments or return value
'(This is used with timer and window events)

Dim subObj As SubProgClass
Dim tmpName As String
Dim b As Integer

Set subObj = New SubProgClass

If subType = SP_SUB Then
    For b = 1 To subParams.Item(subIdx).itemCount
        subObj.Cmd_Var subParams.Item(subIdx).Item(b)
    Next b
    For b = 1 To subRunCode.Item(subIdx).itemCount
        subObj.subProg_runCode.Add subRunCode.Item(subIdx).Item(b)
    Next b
    subObj.subProg_name = subName.Item(subIdx)
    subObj.RunProg
Else
    tmpName = funcName.Item(subIdx)
    subObj.Cmd_Var left(tmpName, Len(tmpName) - 1)
    For b = 1 To funcParams.Item(subIdx).itemCount
        subObj.Cmd_Var funcParams.Item(subIdx).Item(b)
    Next b
    For b = 1 To funcRunCode.Item(subIdx).itemCount
        subObj.subProg_runCode.Add funcRunCode.Item(subIdx).Item(b)
    Next b
    subObj.subProg_name = tmpName
    subObj.RunProg
End If

Set subObj = Nothing


End Sub
Public Sub Cmd_Control(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim tmpParam As String
Dim b, winIdx As Integer
Dim style, exStyle, defStyle As Long
Dim winName, winTitle, winType, parentName As String
Dim winX, winY, winWidth, winHeight As Integer
Dim className As String

cmdStr = Trim(cmdStr)

b = 1
While b <= Len(cmdStr)
    tmpParam = GetString(b, cmdStr, ",")
    b = Len(tmpParam) + b + 1
    params.Add Trim(tmpParam)
Wend

winName = EvalExpression(params.Item(1))
parentName = EvalExpression(params.Item(2))
winTitle = EvalExpression(params.Item(3))
winType = LCase(params.Item(4))
winX = EvalExpression(params.Item(5))
winY = EvalExpression(params.Item(6))
winWidth = EvalExpression(params.Item(7))
winHeight = EvalExpression(params.Item(8))

style = 0
exStyle = WS_EX_CLIENTEDGE
defStyle = WS_VISIBLE Or WS_CHILD

Select Case winType
    Case "button"
        exStyle = 0
        className = "BUTTON"
    Case "statictext"
        exStyle = 0
        className = "STATIC"
    Case "texteditor"
        style = WS_VSCROLL Or WS_HSCROLL Or ES_MULTILINE Or ES_NOHIDESEL
        className = "EDIT"
    Case "textbox"
        style = ES_NOHIDESEL
        className = "EDIT"
    Case "listbox"
        style = LBS_NOTIFY
        className = "LISTBOX"
    Case "combobox"
        style = CBS_DROPDOWN
        className = "COMBOBOX"
    Case "drawbox"
        style = WS_BORDER
        className = "MBGraphWin"
    Case "picbutton"
        exStyle = 0
        style = BS_BITMAP
        className = "BUTTON"
    Case "checkbox"
        exStyle = 0
        style = BS_AUTOCHECKBOX
        className = "BUTTON"
    Case "radiobutton"
        exStyle = 0
        style = BS_AUTORADIOBUTTON
        className = "BUTTON"
    Case "groupbox"
        exStyle = 0
        style = BS_GROUPBOX
        className = "BUTTON"
    Case Else
        Exit Sub
End Select

If windows.itemCount = 0 Then
    ErrorMsg "Window does not exist: '" & parentName & "'"
    Exit Sub
End If

For b = 1 To windows.itemCount
    If parentName = windows.Item(b).winName Then
        winIdx = b
        Exit For
    ElseIf b = windows.itemCount Then
        ErrorMsg "Window does not exist: '" & parentName & "'"
        Exit Sub
    End If
Next b

With windows.Item(winIdx).Controls
    .Add New ControlClass
    .Item(.itemCount).winName = winName
    .Item(.itemCount).ctlType = winType
    .Item(.itemCount).winHandle = CreateWindowEx(exStyle, className, winTitle, _
                                    style Or defStyle, winX, winY, winWidth, winHeight, _
                                    windows.Item(winIdx).winHandle, 0&, App.hInstance, 0&)
    .Item(.itemCount).ctlWndProc = SetWindowLong(.Item(.itemCount).winHandle, GWL_WNDPROC, AddressOf CtlProc)
    .Item(.itemCount).winDC = GetDC(.Item(.itemCount).winHandle)
    .Item(.itemCount).redrawDC = CreateCompatibleDC(.Item(.itemCount).winDC)
    UpdateWindow .Item(.itemCount).winHandle
End With


End Sub
Public Sub Cmd_Event(ByVal cmdStr As String)

Dim nameStr, eventStr, subStr As String
Dim a, b, n, eventIdx As Integer
Dim winObj As Object

cmdStr = Trim(cmdStr)

nameStr = GetString(1, cmdStr, ",")
b = Len(nameStr) + 2
eventStr = GetString(b, cmdStr, ",")
b = b + Len(eventStr) + 1
subStr = Trim(Mid(cmdStr, b))

nameStr = EvalExpression(nameStr)
eventStr = LCase(EvalExpression(eventStr))

For n = 1 To windows.itemCount
    If nameStr = windows.Item(n).winName Then
        Set winObj = windows.Item(n)
        GoTo addEvent
    End If
Next n

For n = 1 To windows.itemCount
    With windows.Item(n).Controls
    For I = 1 To .itemCount
        If nameStr = .Item(I).winName Then
            Set winObj = .Item(I)
            GoTo addEvent
        End If
    Next I
    End With
Next n

ErrorMsg "Window or control does not exist: '" & nameStr & "'"

Exit Sub


addEvent:
    For a = 1 To subName.itemCount
        If subStr = subName.Item(a) Then
            eventIdx = ExistsIn(eventStr, winObj.eventName)
            If eventIdx Then
                winObj.eventName.Item(eventIdx) = eventStr
                winObj.eventSubIdx.Item(eventIdx) = a
                winObj.eventSubType.Item(eventIdx) = SP_SUB
            Else
                winObj.eventName.Add eventStr
                winObj.eventSubIdx.Add a
                winObj.eventSubType.Add SP_SUB
            End If
            Exit Sub
        End If
    Next a
    For b = 1 To funcName.itemCount
        If (subStr & "()" = funcName.Item(b) & ")") Or (subStr = left(funcName.Item(b), Len(funcName.Item(b)) - 1)) Then
            eventIdx = ExistsIn(eventStr, winObj.eventName)
            If eventIdx Then
                winObj.eventName.Item(eventIdx) = eventStr
                winObj.eventSubIdx.Item(eventIdx) = b
                winObj.eventSubType.Item(eventIdx) = SP_FUNC
            Else
                winObj.eventName.Add eventStr
                winObj.eventSubIdx.Add b
                winObj.eventSubType.Add SP_FUNC
            End If
            Exit Sub
        End If
    Next b


End Sub
Public Sub Cmd_Message(ByVal cmdStr As String)

Dim paramStr, msgStr, titleStr As String

cmdStr = Trim(cmdStr)

paramStr = GetString(1, cmdStr, ",")
msgStr = EvalExpression(paramStr)

paramStr = Mid(cmdStr, Len(paramStr) + 2)
titleStr = EvalExpression(paramStr)

MsgBox msgStr, vbExclamation, titleStr


End Sub


Public Sub Cmd_Run(ByVal cmdStr As String)

Dim fileStr, modeStr As String
Dim mode As Integer

cmdStr = Trim(cmdStr)

mode = vbNormalFocus

fileStr = GetString(1, cmdStr, ",")

If Len(fileStr) < Len(cmdStr) Then
    modeStr = Trim(LCase(Mid(cmdStr, Len(fileStr) + 2)))
    Select Case modeStr
        Case "hide"
            mode = vbHide
        Case "minimized"
            mode = vbMinimizedFocus
        Case "maximized"
            mode = vbMaximizedFocus
    End Select
End If

fileStr = EvalExpression(fileStr)

Shell fileStr, mode


End Sub

Public Sub Cmd_Seek(ByVal cmdStr As String)

Dim handleStr As String
Dim seekVal As Long
Dim a As Integer

cmdStr = Trim(cmdStr)

handleStr = GetString(1, cmdStr, ",")
seekVal = EvalExpression(Mid(cmdStr, Len(handleStr) + 2))
handleStr = Trim(handleStr)

For a = 1 To fileHandle.itemCount
    If fileHandle.Item(a) = handleStr Then
        If seekVal > LOF(fileNumber.Item(a)) Then
            ErrorMsg "Past end of file"
            Exit Sub
        End If
        Seek #fileNumber.Item(a), seekVal
        Exit Sub
    End If
Next a

ErrorMsg "File handle does not exist: " & handleStr


End Sub
Public Sub Cmd_ShowConsol()

Output.Show

End Sub
Public Sub Cmd_HideConsol()

Output.Hide

End Sub
Public Function ExistsIn(ByVal val As Variant, ByRef list As ArrayClass) As Integer

Dim n As Integer

For n = 1 To list.itemCount
    If list.Item(n) = val Then
        ExistsIn = n
        Exit Function
    End If
Next n

ExistsIn = 0

End Function
Public Function Func_Loc(ByVal paramStr As String) As Variant

Dim a As Integer

paramStr = Trim(paramStr)

a = ExistsIn(paramStr, fileHandle)

If a Then
    Func_Loc = Loc(fileNumber.Item(a))
Else
    ErrorMsg "File handle does not exist: " & paramStr
End If

End Function


Public Function CtlProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim winIdx, ctlIdx, a, b As Integer
Dim eventStr As String
Dim hdc As Long
Dim ps As PAINTSTRUCT

For a = 1 To windows.itemCount
  With windows.Item(a).Controls
    For b = 1 To .itemCount
        If hwnd = .Item(b).winHandle Then
            winIdx = a
            ctlIdx = b
            a = windows.itemCount
            Exit For
        End If
    Next b
  End With
Next a

If (ctlIdx = 0) Or (winIdx = 0) Then GoTo callDef

With windows.Item(winIdx).Controls.Item(ctlIdx)
Select Case uMsg
    Case WM_LBUTTONUP
        n = ExistsIn("leftbuttonup", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_LBUTTONDOWN
        n = ExistsIn("leftbuttondown", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_LBUTTONDBLCLK
        n = ExistsIn("leftbuttondouble", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_RBUTTONUP
        n = ExistsIn("rightbuttonup", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_RBUTTONDOWN
        n = ExistsIn("rightbuttondown", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_RBUTTONDBLCLK
        n = ExistsIn("rightbuttondouble", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_MOUSEMOVE
        n = ExistsIn("mousemove", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_KEYDOWN
        SetValue "Inkey", Chr(wParam)
        n = ExistsIn("keydown", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_KEYUP
        SetValue "Inkey", Chr(wParam)
        n = ExistsIn("keyup", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_PAINT
        If .ctlType = "drawbox" Then
            hdc = BeginPaint(.winHandle, ps)
            BitBlt hdc, ps.rcPaint.left, ps.rcPaint.top, ps.rcPaint.Right - ps.rcPaint.left, _
                    ps.rcPaint.Bottom - ps.rcPaint.top, .redrawDC, ps.rcPaint.left, ps.rcPaint.top, vbSrcCopy
            EndPaint .winHandle, ps
            n = ExistsIn("paint", .eventName)
            If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
        End If
    Case WM_SIZE
        n = ExistsIn("resize", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
End Select
End With

callDef:
    CtlProc = CallWindowProc(windows.Item(winIdx).Controls.Item(ctlIdx).ctlWndProc, _
                            hwnd, uMsg, wParam, lParam)


End Function
Public Sub ReadRunFile()

Dim fileData, dData As String
Dim listArray() As String

If Dir(App.Path & "\" & App.EXEName & ".lbr") = "" Then
    MsgBox "Lithium BASIC Runtime Engine requiers that there be a runtime file " & _
           "in the same directory as itself in order to run.", vbCritical, _
           "Lithium BASIC Runtime Engine"
    End
End If

Open App.Path & "\" & App.EXEName & ".lbr" For Binary As #1
    fileData = Input(LOF(1), 1)
Close #1

dData = DTask(fileData)

listArray = Split(dData, vbCrLf)

For n = 0 To UBound(listArray) - 1
    runCode.Add listArray(n)
Next n

End Sub


Public Function DTask(ByVal dat As String) As String

Dim dData, dKey, tmpDat As String

dKey = Mid(dat, 11, 1)
tmpDat = Mid(dat, 20, Len(dat) - 29)

For n = 1 To Len(tmpDat) Step 2
    Mid(tmpDat, n, 1) = Chr(Asc(Mid(tmpDat, n, 1)) + 3)
Next n

For n = 1 To Len(tmpDat)
    dData = dData & Chr(Asc(Mid(tmpDat, n, 1)) - Asc(dKey))
Next n

DTask = dData

End Function


Public Function WinProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim winIdx, ctlIdx, n, a, b As Integer
Dim nMsg, menuID, hdc As Long
Dim eventStr As String
Dim ps As PAINTSTRUCT

For n = 1 To windows.itemCount
    If hwnd = windows.Item(n).winHandle Then
        winIdx = n
        Exit For
    End If
Next n

If winIdx = 0 Then GoTo callDef

With windows.Item(winIdx)
Select Case uMsg
    Case WM_DESTROY
        ReleaseDC .winHandle, .winDC
        DeleteObject GetCurrentObject(.redrawDC, OBJ_BITMAP)
        DeleteDC .redrawDC
        For n = 1 To .Controls.itemCount
            ReleaseDC .Controls.Item(n).winHandle, .Controls.Item(n).winDC
            DeleteObject GetCurrentObject(.Controls.Item(n).redrawDC, OBJ_BITMAP)
            DeleteDC .Controls.Item(n).redrawDC
            DestroyWindow .Controls.Item(n).winHandle
        Next n
        If windows.Item(winIdx).winType = "dialog_modal" Then
            If winIdx > 1 Then
                EnableWindow windows.Item(winIdx - 1).winHandle, True
            End If
        End If
        windows.Remove winIdx
        If (windows.itemCount = 0) And (Output.Visible = False) Then progDone = True
    Case WM_CLOSE
        n = ExistsIn("close", .eventName)
        If n Then
            CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
            Exit Function
        End If
    Case WM_LBUTTONUP
        n = ExistsIn("leftbuttonup", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_LBUTTONDOWN
        n = ExistsIn("leftbuttondown", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_LBUTTONDBLCLK
        n = ExistsIn("leftbuttondouble", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_RBUTTONUP
        n = ExistsIn("rightbuttonup", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_RBUTTONDOWN
        n = ExistsIn("rightbuttondown", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_RBUTTONDBLCLK
        n = ExistsIn("rightbuttondouble", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_MOUSEMOVE
        n = ExistsIn("mousemove", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_KEYDOWN
        SetValue "Inkey", Chr(wParam)
        n = ExistsIn("keydown", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_KEYUP
        SetValue "Inkey", Chr(wParam)
        n = ExistsIn("keyup", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_PAINT
        hdc = BeginPaint(.winHandle, ps)
        BitBlt hdc, ps.rcPaint.left, ps.rcPaint.top, ps.rcPaint.Right - ps.rcPaint.left, _
               ps.rcPaint.Bottom - ps.rcPaint.top, .redrawDC, ps.rcPaint.left, ps.rcPaint.top, vbSrcCopy
        EndPaint .winHandle, ps
        n = ExistsIn("paint", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_SIZE
        n = ExistsIn("resize", .eventName)
        If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
    Case WM_COMMAND
        If lParam = 0 Then
            menuID = LOWORD(wParam)
            For a = 1 To .hMenu.itemCount
                For b = 1 To .menuItemID.Item(a).itemCount
                    If menuID = .menuItemID.Item(a).Item(b) Then
                        CallSubProg .menuItemSubIdx.Item(a).Item(b), .menuItemSubType.Item(a).Item(b)
                        GoTo callDef
                    End If
                Next b
            Next a
        End If
        For a = 1 To .Controls.itemCount
            If .Controls.Item(a).winHandle = lParam Then
                ctlIdx = a
                Exit For
            End If
        Next a
        If ctlIdx = 0 Then GoTo callDef
        nMsg = HIWORD(wParam)
        With .Controls.Item(ctlIdx)
        Select Case nMsg
            Case BN_CLICKED
                n = ExistsIn("click", .eventName)
                If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
            Case EN_CHANGE
                n = ExistsIn("change", .eventName)
                If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
            Case LBN_DBLCLK
                n = ExistsIn("doubleselect", .eventName)
                If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
            Case LBN_SELCHANGE 'same as CBN_SELCHANGE
                n = ExistsIn("select", .eventName)
                If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
            Case CBN_EDITCHANGE
                n = ExistsIn("change", .eventName)
                If n Then CallSubProg .eventSubIdx.Item(n), .eventSubType.Item(n)
        End Select
        End With
End Select
End With

callDef:
    WinProc = DefWindowProc(hwnd, uMsg, wParam, lParam)

End Function

Public Sub RegClass(ByVal clsName As String)

Dim winCls As WNDCLASSEX

winCls.cbSize = Len(winCls)
winCls.style = CS_DBLCLKS Or CS_OWNDC
winCls.lpfnWndProc = GetFuncPtr(AddressOf WinProc)
winCls.cbClsExtra = 0&
winCls.cbWndExtra = 0&
winCls.hInstance = App.hInstance
winCls.hIcon = LoadIcon(App.hInstance, IDI_APPLICATION)
winCls.hCursor = LoadCursor(App.hInstance, IDC_ARROW)
winCls.hbrBackground = GetStockObject(1)
winCls.lpszMenuName = 0&
winCls.lpszClassName = clsName
winCls.hIconSm = LoadIcon(App.hInstance, IDI_APPLICATION)

RegisterClassEx winCls

End Sub

Public Sub Cmd_Question(ByVal cmdStr As String)

Dim paramStr, msgStr, titleStr, varStr As String
Dim b, aVal As Integer

cmdStr = Trim(cmdStr)

paramStr = GetString(1, cmdStr, ",")
msgStr = EvalExpression(paramStr)

b = Len(paramStr) + 2

paramStr = GetString(b, cmdStr, ",")
titleStr = EvalExpression(paramStr)

b = b + Len(paramStr) + 1

varStr = Mid(cmdStr, b)

aVal = MsgBox(msgStr, vbYesNo Or vbQuestion, titleStr)
If aVal = vbYes Then
    SetValue varStr, "yes"
Else
    SetValue varStr, "no"
End If


End Sub
Public Sub Cmd_Window(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim tmpParam As String
Dim b As Integer
Dim style, parent As Long
Dim winName, winTitle, winType As String
Dim winX, winY, winWidth, winHeight As Integer

cmdStr = Trim(cmdStr)

b = 1
While b <= Len(cmdStr)
    tmpParam = GetString(b, cmdStr, ",")
    b = Len(tmpParam) + b + 1
    params.Add Trim(tmpParam)
Wend

winName = EvalExpression(params.Item(1))
winTitle = EvalExpression(params.Item(2))
winType = LCase(params.Item(3))
winX = EvalExpression(params.Item(4))
winY = EvalExpression(params.Item(5))
winWidth = EvalExpression(params.Item(6))
winHeight = EvalExpression(params.Item(7))

Select Case winType
    Case "normal"
        style = WS_OVERLAPPEDWINDOW
    Case "dialog"
        style = WS_SYSMENU Or WS_DLGFRAME
    Case "dialog_modal"
        style = WS_SYSMENU Or WS_DLGFRAME Or WS_EX_DLGMODALFRAME
        If windows.itemCount > 0 Then
            parent = windows.Item(windows.itemCount).winHandle
            EnableWindow parent, False
        End If
    Case "popup"
        style = WS_POPUP
End Select

windows.Add New WindowClass

With windows.Item(windows.itemCount)
    .winName = winName
    .winType = winType
    .winHandle = CreateWindowEx(0&, "MicroByteWin", winTitle, _
                            style Or WS_VISIBLE, _
                            winX, winY, winWidth, winHeight, parent, 0&, App.hInstance, 0&)
    .winDC = GetDC(.winHandle)
    .redrawDC = CreateCompatibleDC(.winDC)
    UpdateWindow winHandle
End With


End Sub

Public Sub Cmd_CloseWindow(ByVal cmdStr As String)

Dim n As Integer

cmdStr = EvalExpression(Trim(cmdStr))

For n = 1 To windows.itemCount
    If cmdStr = windows.Item(n).winName Then
        DestroyWindow windows.Item(n).winHandle
        Exit Sub
    End If
Next n

ErrorMsg "Window does not exist: '" & cmdStr & "'"

End Sub
Public Sub AddFuncDef(ByVal cmdLine As Integer)

'Starts loading lines of code from the given line until
'it reaches an END FUNCTION statement, deleting the lines
'along the way

Dim tmpName, tmpParam, cmdStr, paramStr, typeStr As String

cmdStr = Trim(Right(runCode.Item(cmdLine), Len(runCode.Item(cmdLine)) - 9))

tmpName = GetString(1, cmdStr, "(")
paramStr = GetString(Len(tmpName) + 1, LCase(cmdStr), " as ")
paramStr = Mid(cmdStr, Len(tmpName) + 1, Len(paramStr))
typeStr = Right(cmdStr, Len(cmdStr) - (Len(tmpName) + Len(paramStr)))
paramStr = GetString(2, paramStr, ")")

funcName.Add tmpName
funcType.Add typeStr
funcParams.Add New ArrayClass
funcRunCode.Add New ArrayClass

b = 1
While b <= Len(paramStr)
    tmpParam = GetString(b, paramStr, ",")
    b = Len(tmpParam) + b + 1
    funcParams.Item(funcParams.itemCount).Add Trim(tmpParam)
Wend

runCode.Remove cmdLine

n = cmdLine
While LCase(Trim(runCode.Item(n))) <> "end function"
    funcRunCode.Item(funcRunCode.itemCount).Add runCode.Item(n)
    n = n + 1
Wend

For a = n To cmdLine Step -1
    runCode.Remove a
Next a


End Sub
Public Sub AddSubDef(ByVal cmdLine As Integer)

'Starts loading lines of code from the given line until
'it reaches an END SUB statement, deleting the lines along
'the way

Dim tmpName, tmpParam, cmdStr As String

cmdStr = Trim(Right(runCode.Item(cmdLine), Len(runCode.Item(cmdLine)) - 4))

tmpName = GetString(1, cmdStr, " ")

subName.Add tmpName
subParams.Add New ArrayClass
subRunCode.Add New ArrayClass

b = Len(tmpName) + 2
While b <= Len(cmdStr)
    tmpParam = GetString(b, cmdStr, ",")
    b = Len(tmpParam) + b + 1
    subParams.Item(subParams.itemCount).Add Trim(tmpParam)
Wend

runCode.Remove cmdLine

n = cmdLine
While LCase(Trim(runCode.Item(n))) <> "end sub"
    subRunCode.Item(subRunCode.itemCount).Add runCode.Item(n)
    n = n + 1
Wend

For a = n To cmdLine Step -1
    runCode.Remove a
Next a


End Sub

Public Function CallUserFunc(ByVal cmdStr As String) As Variant

Dim funcObj As New SubProgClass
Dim tmpName, tmpParam, paramStr As String

cmdStr = Trim(cmdStr)

tmpName = GetString(1, cmdStr, "(")

For a = 1 To funcName.itemCount
    If tmpName = funcName.Item(a) Then
        'Make the function name a local variable
        funcObj.Cmd_Var tmpName & funcType.Item(a)
        For b = 1 To funcParams.Item(a).itemCount
            'Make the parameter definitions into local function variables
            funcObj.Cmd_Var funcParams.Item(a).Item(b)
        Next b
        'Load the lines of code into the function
        For b = 1 To funcRunCode.Item(a).itemCount
            funcObj.subProg_runCode.Add funcRunCode.Item(a).Item(b)
        Next b
        'Tell the function its name
        funcObj.subProg_name = tmpName
    End If
Next a

paramStr = GetString(Len(tmpName) + 2, cmdStr, ")")

'Fill the parameter variables with the arguments
n = 2
b = 1
While b <= Len(paramStr)
    tmpParam = GetString(b, paramStr, ",")
    b = Len(tmpParam) + b + 1
    funcObj.subProg_varValue.Item(n) = EvalExpression(tmpParam)
    n = n + 1
Wend

'Start running the function
funcObj.RunProg

'Return the value that was stored in the function name variable
CallUserFunc = funcObj.subProg_varValue.Item(1)

'Clean up sub program object
Set funcObj = Nothing

End Function


Public Sub Cmd_Array(ByVal cmdStr As String)

Dim str As String

  cmdStr = Trim(cmdStr)
  varStr = GetString(1, cmdStr, " ")
  idxStr = GetString(Len(GetString(1, varStr, "(")) + 2, varStr, ")")
  firstIdx = GetString(1, idxStr, ",")
  firstVal = EvalExpression(firstIdx)
  If firstVal < 0 Then ErrorMsg "Runtime error: Illegal index value": Exit Sub
  isMultiDim.Add False
  If Len(firstIdx) < Len(idxStr) Then
      secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
      secondVal = EvalExpression(secondIdx)
      If secondVal < 0 Then ErrorMsg "Runtime error: Illegal index value": Exit Sub
      isMultiDim.Item(isMultiDim.itemCount) = True
  End If
  str = Trim(Right(cmdStr, Len(cmdStr) - Len(varStr)))
  If LCase(Trim(Right(str, Len(str) - 3))) = "string" Then
    arrayType.Add DT_STRING
  ElseIf LCase(Trim(Right(str, Len(str) - 3))) = "number" Then
    arrayType.Add DT_NUMBER
  End If
  arrayValue.Add New ArrayClass
  For n = 0 To firstVal
      If isMultiDim.Item(isMultiDim.itemCount) Then
          arrayValue.Item(arrayValue.itemCount).Add New ArrayClass
          For I = 0 To secondVal
            If arrayType.Item(arrayType.itemCount) = DT_STRING Then
                arrayValue.Item(arrayValue.itemCount).Item(n + 1).Add ""
            Else
                arrayValue.Item(arrayValue.itemCount).Item(n + 1).Add 0
            End If
          Next I
      Else
          If arrayType.Item(arrayType.itemCount) = DT_STRING Then
              arrayValue.Item(arrayValue.itemCount).Add ""
          Else
              arrayValue.Item(arrayValue.itemCount).Add 0
          End If
      End If
  Next n
  arrayName.Add left(varStr, Len(GetString(1, varStr, "(")) + 1)

End Sub

Public Sub Cmd_BindVar(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  
  tmpStr = GetString(1, LCase(cmdStr), " to ")
  var1Str = Trim(Mid(tmpStr, 1, Len(tmpStr)))
  var2Str = Trim(Right(cmdStr, Len(cmdStr) - (Len(tmpStr) + 4)))
  For n = 1 To varName.itemCount
    If varName.Item(n) = var2Str Then
        For a = 1 To varName.itemCount
          If varName.Item(a) = var1Str Then
            varBindList.Item(n).Add a
            Return
          End If
        Next a
    End If
  Next n

End Sub


Public Sub Cmd_Pause()

While Not progDone
    DoEvents
    Sleep 1
Wend

End Sub
Public Sub Cmd_Read(ByVal cmdStr As String)

Dim varList As New ArrayClass
Dim a As Integer

cmdStr = Trim(cmdStr)

ParseParams cmdStr, varList

For a = 1 To varList.itemCount
    If readPos > dataList.itemCount Then
        ErrorMsg "Read past end of Data"
        Exit Sub
    End If
    SetValue varList.Item(a), dataList.Item(readPos)
    readPos = readPos + 1
Next a

End Sub
Public Sub Cmd_Restore(ByVal cmdStr As String)

Dim resVal As Integer

    cmdStr = Trim(cmdStr)

If cmdStr = "" Then
    readPos = 1
Else
    resVal = EvalExpression(cmdStr)
    If resVal > 0 And resVal <= dataList.itemCount Then
        readPos = resVal
    Else
        ErrorMsg "Restore beyond Data bounds"
        Exit Sub
    End If
End If
    

End Sub
Public Sub Cmd_Data(ByVal cmdStr As String)

Dim dataItems As New ArrayClass
Dim a As Integer

cmdStr = Trim(cmdStr)

ParseParams cmdStr, dataItems

For a = dataItems.itemCount To 1 Step -1
    dataList.Add EvalExpression(dataItems.Item(a)), 1
Next a

End Sub
Public Sub Cmd_Call(ByVal cmdStr As String)

Dim subObj As New SubProgClass
Dim tmpName, tmpParam As String

cmdStr = Trim(cmdStr)

tmpName = GetString(1, cmdStr, " ")

For a = 1 To subName.itemCount
    If tmpName = subName.Item(a) Then
        For b = 1 To subParams.Item(a).itemCount
            'Make the parameter definitions into local sub variables
            subObj.Cmd_Var subParams.Item(a).Item(b)
        Next b
        'Load the lines of code into the sub
        For b = 1 To subRunCode.Item(a).itemCount
            subObj.subProg_runCode.Add subRunCode.Item(a).Item(b)
        Next b
        'Tell the sub its name
        subObj.subProg_name = tmpName
    End If
Next a

'Fill the parameter variables with the arguments
n = 1
b = Len(tmpName) + 2
While b <= Len(cmdStr)
    tmpParam = GetString(b, cmdStr, ",")
    b = Len(tmpParam) + b + 1
    subObj.subProg_varValue.Item(n) = EvalExpression(tmpParam)
    n = n + 1
Wend

'Start running the sub
subObj.RunProg

'Clean up sub program object
Set subObj = Nothing


End Sub

Public Sub Cmd_ConsolTitle(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  Output.Caption = EvalExpression(cmdStr)

End Sub

Public Sub Cmd_End()

  progDone = True

End Sub

Public Sub Cmd_Error(ByVal cmdStr As String)

Dim paramStr, msgStr, titleStr As String

cmdStr = Trim(cmdStr)

paramStr = GetString(1, cmdStr, ",")
msgStr = EvalExpression(paramStr)

paramStr = Mid(cmdStr, Len(paramStr) + 2)
titleStr = EvalExpression(paramStr)

MsgBox msgStr, vbCritical, titleStr


End Sub
Public Sub Cmd_GoSub(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  
  For n = 1 To labelName.itemCount
    If cmdStr = labelName.Item(n) Then
        gosubLine = lineNum
        lineNum = labelLine.Item(n)
    End If
  Next n


End Sub

Public Sub Cmd_GoTo(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  
  For n = 1 To labelName.itemCount
    If cmdStr = labelName.Item(n) Then lineNum = labelLine.Item(n)
  Next n

End Sub



Public Sub Cmd_SetText(ByVal cmdStr As String)

Dim winName As String
Dim winText As Variant
Dim winHandle As Long
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
winText = EvalExpression(params.Item(2))

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

SetWindowText winHandle, CStr(winText)

End Sub


Public Sub Cmd_SetXY(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim winLeft, winTop As Integer
Dim winHandle As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
winLeft = EvalExpression(params.Item(2))
winTop = EvalExpression(params.Item(3))

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

SetWindowPos winHandle, 0, winLeft, winTop, winWidth, winHeight, SWP_NOSIZE Or SWP_NOZORDER

End Sub


Public Sub Cmd_GetXY(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, leftVar, topVar As String
Dim winHandle As Long
Dim posRect As RECT

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
leftVar = params.Item(2)
topVar = params.Item(3)

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

GetWindowRect winHandle, posRect

SetValue leftVar, posRect.left
SetValue topVar, posRect.top

End Sub


Public Sub Cmd_SetSize(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim winWidth, winHeight As Integer
Dim winHandle As Long

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
winWidth = EvalExpression(params.Item(2))
winHeight = EvalExpression(params.Item(3))

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

SetWindowPos winHandle, 0, 0, 0, winWidth, winHeight, SWP_NOMOVE Or SWP_NOZORDER

End Sub


Public Sub Cmd_GetSize(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim winName, widthVar, heightVar As String
Dim winWidth, winHeight As Integer
Dim winHandle As Long
Dim sizeRect As RECT

cmdStr = Trim(cmdStr)

ParseParams cmdStr, params

winName = EvalExpression(params.Item(1))
widthVar = params.Item(2)
heightVar = params.Item(3)

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Sub
End If

GetWindowRect winHandle, sizeRect
winWidth = sizeRect.Right - sizeRect.left
winHeight = sizeRect.Bottom - sizeRect.top

SetValue widthVar, winWidth
SetValue heightVar, winHeight

End Sub




Public Function Func_GetText(ByVal paramStr As String) As Variant

Dim winName, winText As String
Dim winHandle, textLen As Long

paramStr = Trim(paramStr)

winName = EvalExpression(paramStr)

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Function
End If

textLen = GetWindowTextLength(winHandle)
winText = Space(textLen)
GetWindowText winHandle, winText, textLen + 1

Func_GetText = winText

End Function




Public Function Func_hDC(ByVal paramStr As String) As Variant

Dim winName As String
Dim winDC As Long

paramStr = Trim(paramStr)

winName = EvalExpression(paramStr)

winDC = GetWinDC(winName)

If winDC = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Function
End If

Func_hDC = winDC

End Function


Public Function Func_hWnd(ByVal paramStr As String) As Variant

Dim winName As String
Dim winHandle As Long

paramStr = Trim(paramStr)

winName = EvalExpression(paramStr)

winHandle = GetWinHandle(winName)

If winHandle = 0 Then
    ErrorMsg "Window or control '" & winName & "' does exist"
    Exit Function
End If

Func_hWnd = winHandle

End Function
Public Sub Cmd_Cls()

  Output.display.Text = ""

End Sub

Public Sub Cmd_Input(ByVal cmdStr As String)
  
Dim varStr, expStr, tmpHandle, tmpVal As String
Dim a, b As Integer

  cmdStr = Trim(cmdStr)
  
If left(cmdStr, 1) = "#" Then
    tmpHandle = GetString(1, cmdStr, ",")
    b = Len(tmpHandle) + 2
    tmpHandle = Trim(tmpHandle)
    For a = 1 To fileHandle.itemCount
      If fileHandle.Item(a) = tmpHandle Then
        Select Case fileType.Item(a)
            Case FT_OUTPUT
                ErrorMsg "File opened in OUTPUT mode cannot be inputted from"
            Case FT_APPEND
                ErrorMsg "File opened in APPEND mode cannot be inputted from"
            Case FT_BINARY
                varStr = Mid(cmdStr, b)
                Input #fileNumber.Item(a), tmpVal
                SetValue varStr, tmpVal
            Case Else
                While b <= Len(cmdStr)
                    varStr = GetString(b, cmdStr, ",")
                    b = b + Len(varStr) + 1
                    Input #fileNumber.Item(a), tmpVal
                    SetValue varStr, tmpVal
                Wend
        End Select
        Exit Sub
      End If
    Next a
    ErrorMsg "File handle does not exist: " & tmpHandle
    Exit Sub
Else
    expStr = GetString(1, cmdStr, ",")
  '*** If there is a prompt in the command ***
    If expStr <> cmdStr Then
      Output.display.Text = Output.display.Text & EvalExpression(expStr)
      expStr = Mid(cmdStr, Len(expStr) + 2, Len(cmdStr) - (Len(expStr) + 1))
    End If
  '*******************************************
    expStr = Trim(expStr)
    inputting = True
    While inputting And Not progDone
      DoEvents
      Sleep 1
    Wend
    SetValue expStr, userInput
    userInput = ""
End If


End Sub


Public Sub Cmd_Let(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  varStr = GetString(1, cmdStr, "=")
  expStr = Mid(cmdStr, Len(varStr) + 2, Len(cmdStr) - (Len(varStr) + 1))
  varStr = Trim(varStr)
  SetValue varStr, EvalExpression(expStr)

End Sub
Public Sub Cmd_On(ByVal cmdStr As String)

Dim tmpLabelList As New ArrayClass
Dim valueStr, labelStr, tmpLabel, cmdType As String
Dim valNum As Variant
Dim a, b As Integer

    cmdStr = Trim(cmdStr)

'Get the value if it is an ON...GOTO command
valueStr = GetString(1, LCase(cmdStr), " goto ")
valueStr = left(cmdStr, Len(valueStr))
b = Len(valueStr) + 6
cmdType = "goto"

'If not, then get the value if it is an ON...GOSUB command
If Len(valueStr) = Len(cmdStr) Then
    valueStr = GetString(1, LCase(cmdStr), " gosub ")
    valueStr = left(cmdStr, Len(valueStr))
    b = Len(valueStr) + 7
    cmdType = "gosub"
End If

'Evalute the value
valNum = EvalExpression(valueStr)

'Get list of branch labels
labelStr = Trim(Right(cmdStr, Len(cmdStr) - b))

'Section out each branch label
b = 1
While b <= Len(labelStr)
    tmpLabel = GetString(b, labelStr, ",")
    b = Len(tmpLabel) + b + 1
    tmpLabelList.Add Trim(tmpLabel)
Wend

'Decide which label to jump to
If valNum > 0 And valNum <= tmpLabelList.itemCount Then
    If cmdType = "goto" Then
        Cmd_GoTo tmpLabelList.Item(valNum)
    ElseIf cmdType = "gosub" Then
        Cmd_GoSub tmpLabelList.Item(valNum)
    End If
End If


End Sub
Public Sub Cmd_OnError(ByVal cmdStr As String)

    cmdStr = Trim(cmdStr)

onErrorCmd = cmdStr


End Sub
Public Sub Cmd_Print(ByVal cmdStr As String)

Dim tmpHandle, tmpExp As String
Dim tmpVal As String
Dim b As Integer
Dim noRet As Boolean

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = ";" Then
    noRet = True
    cmdStr = left(cmdStr, Len(cmdStr) - 1)
Else
    noRet = False
End If

If left(cmdStr, 1) = "#" Then
    tmpHandle = GetString(1, cmdStr, ",")
    tmpExp = Right(cmdStr, Len(cmdStr) - (Len(tmpHandle) + 1))
    tmpHandle = Trim(tmpHandle)
    tmpVal = EvalExpression(tmpExp)
    For b = 1 To fileHandle.itemCount
        If fileHandle.Item(b) = tmpHandle Then
            Select Case fileType.Item(b)
                Case FT_INPUT
                    ErrorMsg "File opened in INPUT mode cannot be printed to"
                Case FT_BINARY
                    Put #fileNumber.Item(b), , tmpVal
                Case Else
                    If noRet Then
                        Print #fileNumber.Item(b), tmpVal;
                    Else
                        Print #fileNumber.Item(b), tmpVal
                    End If
            End Select
            Exit Sub
        End If
    Next b
    ErrorMsg "File handle does not exist: " & tmpHandle
Else
    If noRet Then
        Output.display.Text = Output.display.Text & EvalExpression(cmdStr)
    Else
        Output.display.Text = Output.display.Text & EvalExpression(cmdStr) & vbCrLf
    End If
End If


End Sub

Public Sub Cmd_ReDim(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  For n = 1 To arrayName.itemCount
    If left(cmdStr, Len(arrayName.Item(n))) = arrayName.Item(n) And Right(cmdStr, 1) = ")" Then
        idxStr = GetString(Len(GetString(1, cmdStr, "(")) + 2, cmdStr, ")")
        firstIdx = GetString(1, idxStr, ",")
        firstVal = EvalExpression(firstIdx)
        If firstVal < 0 Then MsgBox "Runtime error: Illegal index value", vbCritical, "Lithium BASIC Runtime": End
        isMultiDim.Item(n) = False
        If Len(firstIdx) < Len(idxStr) Then
            secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
            secondVal = EvalExpression(secondIdx)
            If secondVal < 0 Then MsgBox "Runtime error: Illegal index value", vbCritical, "Lithium BASIC Runtime": End
            isMultiDim.Item(n) = True
        End If
        For I = arrayValue.Item(n).itemCount To 1 Step -1
            arrayValue.Item(n).Remove I
        Next I
        For I = 0 To firstVal
            If isMultiDim.Item(n) Then
                arrayValue.Item(n).Add New ArrayClass
                For a = 0 To secondVal
                  If arrayType.Item(n) = DT_STRING Then
                      arrayValue.Item(n).Item(I + 1).Add ""
                  Else
                      arrayValue.Item(n).Item(I + 1).Add 0
                  End If
                Next a
            Else
                If arrayType.Item(n) = DT_STRING Then
                    arrayValue.Item(n).Add ""
                Else
                    arrayValue.Item(n).Add 0
                End If
            End If
        Next I
        Exit Sub
    End If
  Next n

End Sub

Public Sub Cmd_Return()

cmdStr = Trim(cmdStr)

If gosubLine = 0 Then
    ErrorMsg "RETURN without GOSUB"
    Exit Sub
Else
    lineNum = gosubLine
    gosubLine = 0
End If

End Sub
Public Sub Cmd_StopTimer(ByVal cmdStr As String)

Dim tmpName As String

'If debugging Then Exit Sub

tmpName = EvalExpression(cmdStr)

For n = 1 To timerName.itemCount
    If timerName.Item(n) = tmpName Then
        KillTimer 0, timerID.Item(n)
        timerName.Remove n
        timerID.Remove n
        timerSubIdx.Remove n
        timerSubType.Remove n
        Exit Sub
    End If
Next n

End Sub

Public Sub Cmd_Swap(ByVal cmdStr As String)

Dim var1, var2 As String
Dim value1, value2 As Variant
Dim a As Integer

    cmdStr = Trim(cmdStr)

'Seperate the two variables/arrays
var1 = GetString(1, cmdStr, ",")
a = Len(var1) + 1
var2 = Right(cmdStr, Len(cmdStr) - a)

'Trim the spaces from them
var1 = Trim(var1)
var2 = Trim(var2)

'Get the value of each
value1 = EvalExpression(var1)
value2 = EvalExpression(var2)

'Swap the two values
SetValue var1, value2
SetValue var2, value1


End Sub
Public Sub Cmd_TextColor(ByVal cmdStr As String)

Dim color As Variant

cmdStr = Trim(cmdStr)

color = EvalExpression(cmdStr)

Output.display.ForeColor = color

End Sub


Public Sub Cmd_BGColor(ByVal cmdStr As String)

Dim color As Variant

cmdStr = Trim(cmdStr)

color = EvalExpression(cmdStr)

Output.display.BackColor = color

End Sub

Public Sub Cmd_Timer(ByVal cmdStr As String)

Dim nameStr, subStr, timeStr As String
Dim timeVal As Variant

cmdStr = Trim(cmdStr)

nameStr = GetString(1, cmdStr, ",")
timeStr = GetString(Len(nameStr) + 2, cmdStr, ",")
subStr = Trim(Mid(cmdStr, Len(nameStr) + Len(timeStr) + 3))

timerName.Add EvalExpression(nameStr)
timeVal = EvalExpression(timeStr)

For a = 1 To subName.itemCount
    If subStr = subName.Item(a) Then
        timerSubIdx.Add a
        timerSubType.Add SP_SUB
        timerID.Add SetTimer(0, 0, timeVal, AddressOf TimerProc)
        Exit Sub
    End If
Next a

For b = 1 To funcName.itemCount
    If (subStr & "()" = funcName.Item(b) & ")") Or (subStr = left(funcName.Item(b), Len(funcName.Item(b)) - 1)) Then
        timerSubIdx.Add b
        timerSubType.Add SP_FUNC
        timerID.Add SetTimer(0, 0, timeVal, AddressOf TimerProc)
        Exit Sub
    End If
Next b


End Sub
Public Sub Cmd_UnbindVar(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  tmpStr = GetString(1, LCase(cmdStr), " from ")
  var1Str = Trim(Mid(tmpStr, 1, Len(tmpStr)))
  var2Str = Trim(Right(cmdStr, Len(cmdStr) - (Len(tmpStr) + 6)))
  For n = 1 To varName.itemCount
    If varName.Item(n) = var2Str Then
        For a = 1 To varBindList.Item(n).itemCount
            If varName.Item(varBindList.Item(n).Item(a)) = var1Str Then
                varBindList.Item(n).Remove a
                Return
            End If
        Next a
        ErrorMsg "Variable cannot be unbound"
        Exit Sub
    End If
  Next n

End Sub
Public Sub Cmd_Var(ByVal cmdStr As String)

Dim tmpVar, varStr, typeStr As String

  cmdStr = Trim(cmdStr)

varStr = GetString(1, LCase(cmdStr), " as ")
varStr = left(cmdStr, Len(varStr))

typeStr = LCase(Trim(Right(cmdStr, Len(cmdStr) - (Len(varStr) + 4))))

b = 1
While b <= Len(varStr)
    tmpVar = GetString(b, varStr, ",")
    b = Len(tmpVar) + b + 1
    If typeStr = "number" Then
        varType.Add DT_NUMBER
        varValue.Add 0
    ElseIf typeStr = "string" Then
        varType.Add DT_STRING
        varValue.Add ""
    End If
    varName.Add Trim(tmpVar)
    varBindList.Add New ArrayClass
Wend


End Sub


Public Sub DebugWait()

If debugging Then
    If debugState = DS_STEP Then
        debugState = DS_PAUSE
    ElseIf debugState = DS_PAUSE Then
        While (debugState = DS_PAUSE) And (Not progDone)
            DoEvents
            Sleep 1
        Wend
        If debugState = DS_STEP Then debugState = DS_PAUSE
    End If
End If


End Sub


Public Sub EndProg()

For n = 1 To fileHandle.itemCount
    Close fileNumber.Item(n)
Next n

For n = 1 To timerID.itemCount
    KillTimer 0, timerID.Item(n)
Next n

For n = windows.itemCount To 1 Step -1
    DestroyWindow windows.Item(n).winHandle
Next n

End

End Sub
Public Sub ErrorMsg(msgStr As String)

If onErrorCmd = "" Then
    MsgBox "Something unexpected has happened and I can't go on: " & vbCrLf & _
            vbCrLf & _
            Space(10) & msgStr & vbCrLf & _
            vbCrLf & _
            "Awfully sorry about that. Please don't blame yourself.", _
            vbCritical, "Lithium BASIC Runtime"
    progDone = True
    errorFlag = True
Else
    SetValue "ErrorMsg", msgStr
    RunCmd onErrorCmd
End If

End Sub

Public Function Func_Asc(ByVal paramStr As String) As Variant

On Error Resume Next

paramStr = Trim(paramStr)

Func_Asc = Asc(EvalExpression(paramStr))

End Function

Public Function Func_Not(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Not = (Not EvalExpression(paramStr))

End Function

Public Function Func_Len(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Len = Len(EvalExpression(paramStr))


End Function
Public Function Func_Rnd() As Variant

  Randomize
  Func_Rnd = Rnd

End Function
Public Function Func_Val(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Val = val(EvalExpression(paramStr))

End Function
Public Function Func_Chr(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

Func_Chr = Chr(EvalExpression(paramStr))

End Function
Public Function Func_Str(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Str = str(EvalExpression(paramStr))

End Function
Public Function Func_Upper(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Upper = UCase(EvalExpression(paramStr))

End Function
Public Function Func_Lower(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Lower = LCase(EvalExpression(paramStr))

End Function
Public Function Func_Trim(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Trim = Trim(EvalExpression(paramStr))

End Function

Public Function Func_Mid(ByVal paramStr As String) As Variant

Dim params As New ArrayClass
Dim stringVal As String
Dim starting, length As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

stringVal = EvalExpression(params.Item(1))
starting = EvalExpression(params.Item(2))

If params.itemCount = 3 Then
    length = EvalExpression(params.Item(3))
    Func_Mid = Mid(stringVal, starting, length)
Else
    Func_Mid = Mid(stringVal, starting)
End If

End Function
Public Function Func_Left(ByVal paramStr As String) As Variant

Dim paramList As New ArrayClass

paramStr = Trim(paramStr)

    param = GetString(1, paramStr, ",")
    paramList.Add EvalExpression(param)
    a = Len(param) + 2
    paramList.Add EvalExpression(GetString(a, paramStr, ")"))
    Func_Left = left(paramList.Item(1), paramList.Item(2))

End Function
Public Function Func_Right(ByVal paramStr As String) As Variant

Dim paramList As New ArrayClass

paramStr = Trim(paramStr)

    param = GetString(1, paramStr, ",")
    paramList.Add EvalExpression(param)
    a = Len(param) + 2
    paramList.Add EvalExpression(GetString(a, paramStr, ")"))
    Func_Right = Right(paramList.Item(1), paramList.Item(2))

End Function
Public Function Func_InStr(ByVal paramStr As String) As Variant

Dim params As New ArrayClass
Dim string1, string2 As String
Dim starting As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

string1 = EvalExpression(params.Item(1))
string2 = EvalExpression(params.Item(2))
If params.itemCount = 3 Then starting = EvalExpression(params.Item(3))

If params.itemCount = 3 Then
    Func_InStr = InStr(starting, string1, string2)
Else
    Func_InStr = InStr(1, string1, string2)
End If

End Function
Public Function Func_Int(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Int = Int(EvalExpression(paramStr))

End Function
Public Function Func_Abs(ByVal paramStr As String) As Variant

paramStr = Trim(paramStr)

  Func_Abs = Abs(EvalExpression(paramStr))


End Function
Public Function GetString(ByVal start As Integer, ByVal str As String, ByVal endStr As String) As String

'This sub sections out a substring from within the given string,
'starting at the given point, ending at the given character,
'and not counting any character within parentheses or quotes.

Dim inString As Boolean
Dim parNum As Integer

inString = False
parNum = 0

For a = start To Len(str)
  If Mid(str, a, Len(endStr)) = endStr Then
    If parNum = 0 And inString = False Then
      Exit Function
    End If
  End If
  If Mid(str, a, 1) = Chr(34) Then
    If inString = False Then inString = True Else inString = False
  End If
  If Mid(str, a, 1) = "(" And inString = False Then
    parNum = parNum + 1
  ElseIf Mid(str, a, 1) = ")" And inString = False Then
    If parNum > 0 Then parNum = parNum - 1
  End If
  GetString = GetString & Mid(str, a, 1)
Next a

End Function

Public Sub LoadFunctions()

numFunc.Add "abs("
numFunc.Add "asc("
numFunc.Add "not("
numFunc.Add "int("
numFunc.Add "len("
numFunc.Add "rnd("
numFunc.Add "val("
numFunc.Add "loc("
numFunc.Add "hwnd("
numFunc.Add "hdc("
numFunc.Add "getselidx("
numFunc.Add "itemcount("
numFunc.Add "linecount("
numFunc.Add "min("
numFunc.Add "max("
numFunc.Add "lof("
numFunc.Add "eof("
numFunc.Add "hbmp("
numFunc.Add "sqr("
numFunc.Add "getstate("
numFunc.Add "collide("

strFunc.Add "chr("
strFunc.Add "str("
strFunc.Add "upper("
strFunc.Add "lower("
strFunc.Add "trim("
strFunc.Add "left("
strFunc.Add "mid("
strFunc.Add "right("
strFunc.Add "instr("
strFunc.Add "gettext("
strFunc.Add "getitem("
strFunc.Add "getseltext("
strFunc.Add "getlinetext("
strFunc.Add "getclipboardtext("
strFunc.Add "inputbox("
strFunc.Add "space("
strFunc.Add "fileopen("
strFunc.Add "filesave("
strFunc.Add "date("
strFunc.Add "time("
strFunc.Add "input("

End Sub


Public Function RunBlock(ByVal startLne As Integer, ByVal endLne As Integer) As Boolean

Dim ifNum, whileNum, forNum, selectNum As Integer
Dim n As Integer

RunBlock = True

If LCase(left(runCode.Item(startLne), 3)) = "if " Then
    ifNum = 0
    If LCase(Right(runCode.Item(startLne), 5)) = " then" Then
        For n = startLne To endLne
            If (LCase(Trim(runCode.Item(n))) = "end if" And ifNum = 0) Then
                RunIfBlock startLne, n - 1, False
                n = endLne
            ElseIf (LCase(left(LTrim(runCode.Item(n)), 3)) = "if ") And (n > startLne) Then
                If LCase(Right(runCode.Item(n), 5)) = " then" Then
                    ifNum = ifNum + 1
                End If
            ElseIf LCase(Trim(runCode.Item(n))) = "end if" Then
                ifNum = ifNum - 1
            End If
        Next n
    Else
       RunIfBlock startLne, 0, True
    End If
ElseIf LCase(left(runCode.Item(startLne), 6)) = "while " Then
    whileNum = 0
    For n = startLne To endLne
        If (LCase(Trim(runCode.Item(n))) = "wend" And whileNum = 0) Then
            RunWhileBlock startLne, n - 1
            n = endLne
        ElseIf (LCase(left(LTrim(runCode.Item(n)), 6)) = "while ") And (n > startLne) Then
            whileNum = whileNum + 1
        ElseIf LCase(Trim(runCode.Item(n))) = "wend" Then
            whileNum = whileNum - 1
        End If
    Next n
ElseIf LCase(left(runCode.Item(startLne), 4)) = "for " Then
    forNum = 0
    For n = startLne To endLne
        If (LCase(left(runCode.Item(n), 5)) = "next " And forNum = 0) Then
            RunForBlock startLne, n - 1
            n = endLne
        ElseIf (LCase(left(LTrim(runCode.Item(n)), 4)) = "for ") And (n > startLne) Then
            forNum = forNum + 1
        ElseIf LCase(left(runCode.Item(n), 5)) = "next " Then
            forNum = forNum - 1
        End If
    Next n
ElseIf LCase(left(runCode.Item(startLne), 12)) = "select case " Then
    selectNum = 0
    For n = startLne To endLne
        If (LCase(Trim(runCode.Item(n))) = "end select" And selectNum = 0) Then
            RunSelectBlock startLne, n - 1
            n = endLne
        ElseIf (LCase(left(LTrim(runCode.Item(n)), 12)) = "select case ") And (n > startLne) Then
            selectNum = selectNum + 1
        ElseIf LCase(Trim(runCode.Item(n))) = "end select" Then
            selectNum = selectNum - 1
        End If
    Next n
Else
    RunBlock = False
End If

End Function
Public Sub RunCmd(ByVal cmdStr As String)

Dim tmpCmdStr As String

cmdStr = Trim(cmdStr)

If LCase(left(cmdStr, 6)) = "print " Then
  Cmd_Print Mid(cmdStr, 7)
ElseIf LCase(cmdStr) = "print" Then
  Cmd_Print ""
ElseIf LCase(left(cmdStr, 6)) = "input " Then
  Cmd_Input Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 4)) = "let " Then
  Cmd_Let Mid(cmdStr, 5)
ElseIf LCase(left(cmdStr, 4)) = "var " Then
  Cmd_Var Mid(cmdStr, 5)
ElseIf LCase(cmdStr) = "cls" Then
  Cmd_Cls
ElseIf LCase(left(cmdStr, 5)) = "goto " Then
  Cmd_GoTo Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 6)) = "gosub " Then
  Cmd_GoSub Mid(cmdStr, 7)
ElseIf LCase(cmdStr) = "return" Then
  Cmd_Return
ElseIf LCase(cmdStr) = "end" Then
  Cmd_End
ElseIf LCase(left(cmdStr, 12)) = "consoltitle " Then
  Cmd_ConsolTitle Mid(cmdStr, 13)
ElseIf LCase(left(cmdStr, 6)) = "array " Then
  Cmd_Array Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 6)) = "redim " Then
  Cmd_ReDim Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 8)) = "bindvar " Then
  Cmd_BindVar Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 10)) = "unbindvar " Then
  Cmd_UnbindVar Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 5)) = "open " Then
  Cmd_Open Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 6)) = "close " Then
  Cmd_Close Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 5)) = "call " Then
  Cmd_Call Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 4)) = "dim " Then
  Cmd_Array Mid(cmdStr, 5)
ElseIf LCase(left(cmdStr, 8)) = "onerror " Then
  Cmd_OnError Mid(cmdStr, 9)
ElseIf LCase(cmdStr) = "onerror" Then
  Cmd_OnError ""
ElseIf LCase(left(cmdStr, 5)) = "swap " Then
  Cmd_Swap Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 3)) = "on " Then
  Cmd_On Mid(cmdStr, 4)
ElseIf LCase(left(cmdStr, 5)) = "read " Then
  Cmd_Read Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 8)) = "restore " Then
  Cmd_Restore Mid(cmdStr, 9)
ElseIf LCase(cmdStr) = "restore" Then
  Cmd_Restore ""
ElseIf LCase(left(cmdStr, 10)) = "textcolor " Then
  Cmd_TextColor Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 8)) = "bgcolor " Then
  Cmd_BGColor Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 6)) = "timer " Then
  Cmd_Timer Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 10)) = "stoptimer " Then
  Cmd_StopTimer Mid(cmdStr, 11)
ElseIf LCase(Trim(cmdStr)) = "pause" Then
  Cmd_Pause
ElseIf LCase(left(cmdStr, 8)) = "message " Then
  Cmd_Message Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 9)) = "question " Then
  Cmd_Question Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 6)) = "error " Then
  Cmd_Message Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 4)) = "run " Then
  Cmd_Run Mid(cmdStr, 5)
ElseIf LCase(left(cmdStr, 5)) = "seek " Then
  Cmd_Seek Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 7)) = "window " Then
  Cmd_Window Mid(cmdStr, 8)
ElseIf LCase(left(cmdStr, 12)) = "closewindow " Then
  Cmd_CloseWindow Mid(cmdStr, 13)
ElseIf LCase(left(cmdStr, 6)) = "event " Then
  Cmd_Event Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 8)) = "control " Then
  Cmd_Control Mid(cmdStr, 9)
ElseIf LCase(Trim(cmdStr)) = "showconsol" Then
  Cmd_ShowConsol
ElseIf LCase(Trim(cmdStr)) = "hideconsol" Then
  Cmd_HideConsol
ElseIf LCase(left(cmdStr, 7)) = "enable " Then
  Cmd_Enable Mid(cmdStr, 8)
ElseIf LCase(left(cmdStr, 8)) = "disable " Then
  Cmd_Disable Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 5)) = "show " Then
  Cmd_Show Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 5)) = "hide " Then
  Cmd_Hide Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 8)) = "getsize " Then
  Cmd_GetSize Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 6)) = "getxy " Then
  Cmd_GetXY Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 8)) = "settext " Then
  Cmd_SetText Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 8)) = "setsize " Then
  Cmd_SetSize Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 6)) = "setxy " Then
  Cmd_SetXY Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 5)) = "menu " Then
  Cmd_Menu Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 8)) = "additem " Then
  Cmd_AddItem Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 8)) = "delitem " Then
  Cmd_DelItem Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 10)) = "setselidx " Then
  Cmd_SetSelIdx Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 8)) = "setitem " Then
  Cmd_SetItem Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 7)) = "getsel " Then
  Cmd_GetSel Mid(cmdStr, 8)
ElseIf LCase(left(cmdStr, 7)) = "setsel " Then
  Cmd_SetSel Mid(cmdStr, 8)
ElseIf LCase(left(cmdStr, 11)) = "setseltext " Then
  Cmd_SetSelText Mid(cmdStr, 12)
ElseIf LCase(left(cmdStr, 17)) = "setclipboardtext " Then
  Cmd_SetClipboardText Mid(cmdStr, 18)
ElseIf LCase(left(cmdStr, 5)) = "line " Then
  Cmd_Line Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 4)) = "box " Then
  Cmd_Box Mid(cmdStr, 5)
ElseIf LCase(left(cmdStr, 7)) = "circle " Then
  Cmd_Circle Mid(cmdStr, 8)
ElseIf LCase(left(cmdStr, 9)) = "linesize " Then
  Cmd_LineSize Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 8)) = "loadimg " Then
  Cmd_LoadImg Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 10)) = "unloadimg " Then
  Cmd_UnloadImg Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 8)) = "drawimg " Then
  Cmd_DrawImg Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 7)) = "getimg " Then
  Cmd_GetImg Mid(cmdStr, 8)
ElseIf LCase(left(cmdStr, 6)) = "stick " Then
  Cmd_Stick Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 8)) = "refresh " Then
  Cmd_Refresh Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 6)) = "clear " Then
  Cmd_Clear Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 10)) = "backcolor " Then
  Cmd_BackColor Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 10)) = "forecolor " Then
  Cmd_ForeColor Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 9)) = "setpixel " Then
  Cmd_SetPixel Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 9)) = "drawtext " Then
  Cmd_DrawText Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 8)) = "setfont " Then
  Cmd_SetFont Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 7)) = "sprite " Then
  Cmd_Sprite Mid(cmdStr, 8)
ElseIf LCase(left(cmdStr, 10)) = "delsprite " Then
  Cmd_DelSprite Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 12)) = "drawsprites " Then
  Cmd_DrawSprites Mid(cmdStr, 13)
ElseIf LCase(left(cmdStr, 9)) = "addframe " Then
  Cmd_AddFrame Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 9)) = "delframe " Then
  Cmd_DelFrame Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 10)) = "spritepos " Then
  Cmd_SpritePos Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 11)) = "spritesize " Then
  Cmd_SpriteSize Mid(cmdStr, 12)
ElseIf LCase(left(cmdStr, 11)) = "spriterate " Then
  Cmd_SpriteRate Mid(cmdStr, 12)
ElseIf LCase(left(cmdStr, 13)) = "spriterotate " Then
  Cmd_SpriteRotate Mid(cmdStr, 14)
ElseIf LCase(left(cmdStr, 11)) = "spriteshow " Then
  Cmd_SpriteShow Mid(cmdStr, 12)
ElseIf LCase(left(cmdStr, 11)) = "spritehide " Then
  Cmd_SpriteHide Mid(cmdStr, 12)
ElseIf LCase(left(cmdStr, 11)) = "spriteplay " Then
  Cmd_SpritePlay Mid(cmdStr, 12)
ElseIf LCase(left(cmdStr, 12)) = "spritepause " Then
  Cmd_SpritePause Mid(cmdStr, 13)
ElseIf LCase(left(cmdStr, 9)) = "spritebg " Then
  Cmd_SpriteBG Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 7)) = "runcmd " Then
  Cmd_RunCmd Mid(cmdStr, 8)
ElseIf LCase(left(cmdStr, 11)) = "closesound " Then
  Cmd_CloseSound Mid(cmdStr, 12)
ElseIf LCase(left(cmdStr, 10)) = "opensound " Then
  Cmd_OpenSound Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 11)) = "pausesound " Then
  Cmd_PauseSound Mid(cmdStr, 12)
ElseIf LCase(left(cmdStr, 10)) = "playsound " Then
  Cmd_PlaySound Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 12)) = "resumesound " Then
  Cmd_ResumeSound Mid(cmdStr, 13)
ElseIf LCase(left(cmdStr, 10)) = "stopsound " Then
  Cmd_StopSound Mid(cmdStr, 11)
ElseIf LCase(Trim(cmdStr)) = "beep" Then
  Cmd_Beep
ElseIf LCase(left(cmdStr, 9)) = "getfiles " Then
  Cmd_GetFiles Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 8)) = "getdirs " Then
  Cmd_GetDirs Mid(cmdStr, 9)
ElseIf LCase(left(cmdStr, 5)) = "name " Then
  Cmd_Name Mid(cmdStr, 6)
ElseIf LCase(left(cmdStr, 9)) = "setstate " Then
  Cmd_SetState Mid(cmdStr, 10)
ElseIf LCase(left(cmdStr, 10)) = "buttonimg " Then
  Cmd_ButtonImg Mid(cmdStr, 11)
ElseIf LCase(left(cmdStr, 6)) = "mkdir " Then
  Cmd_MkDir Mid(cmdStr, 7)
ElseIf LCase(left(cmdStr, 6)) = "rmdir " Then
  Cmd_RmDir Mid(cmdStr, 7)

ElseIf left(cmdStr, 1) = "@" Then
  'NULL
Else
  If Len(GetString(1, cmdStr, "=")) = Len(cmdStr) Then
    ErrorMsg "Bad or unexpected command string: " & UCase(cmdStr)
    Exit Sub
  Else
    Cmd_Let cmdStr
  End If
End If

DoEvents

End Sub



Public Sub Cmd_Beep()

Beep

End Sub


Public Sub Cmd_RunCmd(ByVal cmdStr As String)

Dim cmdLine As String

cmdStr = Trim(cmdStr)

cmdLine = EvalExpression(cmdStr)

RunCmd cmdLine

End Sub


Public Sub RunForBlock(ByVal startLne As Integer, ByVal endLne As Integer)

Dim expStr, varStr As String
Dim varVal, val1, val2, stepVal As Integer

varStr = GetString(5, runCode.Item(startLne), "=")

a = Len(varStr) + 6

expStr = GetString(a, LCase(runCode.Item(startLne) & " "), " to ")
expStr = Mid(runCode.Item(startLne), a, Len(expStr))
val1 = EvalExpression(expStr)

a = a + Len(expStr) + 3

expStr = GetString(a, LCase(runCode.Item(startLne) & " "), " step ")
expStr = Mid(runCode.Item(startLne), a, Len(expStr))
val2 = EvalExpression(expStr)

If a + Len(expStr) < Len(runCode.Item(startLne)) Then
    a = a + Len(expStr) + 5
    expStr = Right(runCode.Item(startLne), Len(runCode.Item(startLne)) - a)
    stepVal = EvalExpression(expStr)
Else
    stepVal = 1
End If

For varVal = val1 To val2 Step stepVal
  SetValue varStr, varVal
  If debugging Then debugWin.code.ListIndex = startLne
  If endLne > startLne Then
    For lineNum = startLne + 1 To endLne
        DebugWait
        If Not RunBlock(lineNum, endLne) Then
            RunCmd runCode.Item(lineNum)
        End If
        If progDone Then Exit Sub
        If lineNum < startLne Or lineNum > endLne Then Exit Sub
        If debugging Then debugWin.code.ListIndex = lineNum
    Next lineNum
  End If
  DebugWait
  If debugging Then debugWin.code.ListIndex = startLne - 1
  DebugWait
Next varVal

lineNum = endLne + 1


End Sub
Public Sub RunIfBlock(ByVal startLne As Integer, ByVal endLne As Integer, ByVal oneLine As Boolean)

Dim cmdStr, expStr, tmpLine As String
Dim ifExp As New ArrayClass
Dim ifStart As New ArrayClass
Dim ifEnd As New ArrayClass
Dim ifNum, n As Integer

expStr = GetString(4, LCase(runCode.Item(startLne)) & " ", " then ")
expStr = Mid(runCode.Item(startLne), 4, Len(expStr))

If oneLine Then
    cmdStr = LTrim(Right(runCode.Item(startLne), Len(runCode.Item(startLne)) - (Len(expStr) + 8)))
    If EvalExpression(expStr) Then RunCmd cmdStr
    Exit Sub
Else
  If endLne > startLne Then
    ifExp.Add expStr
    ifStart.Add startLne + 1
    For n = startLne + 1 To endLne
        If LCase(left(runCode.Item(n), 3)) = "if " Then
            If LCase(Right(runCode.Item(n), 5)) = " then" Then
                ifNum = ifNum + 1
            End If
        ElseIf LCase(Trim(runCode.Item(n))) = "end if" Then
            ifNum = ifNum - 1
        ElseIf LCase(left(runCode.Item(n), 7)) = "elseif " And ifNum = 0 Then
            ifEnd.Add n - 1
            expStr = GetString(8, LCase(runCode.Item(n)) & " ", " then ")
            expStr = Mid(runCode.Item(n), 8, Len(expStr))
            ifExp.Add expStr
            ifStart.Add n + 1
        ElseIf LCase(Trim((runCode.Item(n)))) = "else" And ifNum = 0 Then
            ifEnd.Add n - 1
            ifExp.Add "1"
            ifStart.Add n + 1
        End If
    Next n
    ifEnd.Add endLne
    For n = 1 To ifExp.itemCount
        If EvalExpression(ifExp.Item(n)) Then
            If debugging Then debugWin.code.ListIndex = ifStart.Item(n) - 1
            For lineNum = ifStart.Item(n) To ifEnd.Item(n)
                DebugWait
                If Not RunBlock(lineNum, endLne) Then
                    RunCmd runCode.Item(lineNum)
                End If
                If progDone Then Exit Sub
                If lineNum < startLne Or lineNum > endLne Then Exit Sub
                If debugging Then debugWin.code.ListIndex = lineNum
            Next lineNum
            lineNum = endLne + 1
            Exit Sub
        End If
        If debugging Then debugWin.code.ListIndex = ifEnd.Item(n)
        DebugWait
    Next n
    lineNum = endLne + 1
  End If
End If

End Sub

Public Sub OldRunIfBlock(ByVal startLne As Integer, ByVal endLne As Integer, ByVal oneLine As Boolean)

Dim cmdStr, expStr, tmpLine As String
Dim ifIsTrue As Boolean

expStr = GetString(4, LCase(runCode.Item(startLne)) & " ", " then ")
expStr = Mid(runCode.Item(startLne), 4, Len(expStr))

If EvalExpression(expStr) Then
    ifIsTrue = True
Else
    ifIsTrue = False
End If

If oneLine Then
    cmdStr = LTrim(Right(runCode.Item(startLne), Len(runCode.Item(startLne)) - (Len(expStr) + 8)))
    If ifIsTrue Then RunCmd cmdStr
Else
    If Len(expStr) + 8 < Len(runCode.Item(startLne)) Then
        cmdStr = LTrim(Right(runCode.Item(startLne), Len(runCode.Item(startLne)) - (Len(expStr) + 8)))
        If ifIsTrue Then RunCmd cmdStr: If progDone Then Exit Sub
    End If
    For lineNum = startLne + 1 To endLne
        If debugging Then
            If debugState = DS_STEP Then
                debugState = DS_PAUSE
            ElseIf debugState = DS_PAUSE Then
                While debugState = DS_PAUSE
                    DoEvents
                    Sleep 1
                Wend
                If debugState = DS_STEP Then debugState = DS_PAUSE
            End If
            debugWin.code.ListIndex = lineNum - 1
        End If
        If LCase(left(runCode.Item(lineNum), 7)) = "elseif " Then
            If ifIsTrue Then
                lineNum = endLne + 1
                Exit Sub
            Else
                'Evaluate expression in between ELSEIF and THEN
                expStr = GetString(8, LCase(runCode.Item(lineNum)) & " ", " then ")
                expStr = Mid(runCode.Item(lineNum), 8, Len(expStr))
                If EvalExpression(expStr) Then
                    ifIsTrue = True
                End If
                'Run any command on the same line after ELSEIF...THEN
                If Len(expStr) + 12 < Len(runCode.Item(lineNum)) Then
                    cmdStr = LTrim(Right(runCode.Item(lineNum), Len(runCode.Item(lineNum)) - (Len(expStr) + 12)))
                    If ifIsTrue Then RunCmd cmdStr: If progDone Then Exit Sub
                End If
            End If
        ElseIf LCase(runCode.Item(lineNum)) = "else" Then
            If ifIsTrue Then
                lineNum = endLne + 1
                Exit Sub
            Else
                ifIsTrue = True
            End If
        Else
            If ifIsTrue Then
                If Not RunBlock(lineNum, endLne) Then
                    RunCmd runCode.Item(lineNum)
                End If
                If progDone Then Exit Sub
            End If
        End If
    If lineNum < startLne Or lineNum > endLne Then Exit Sub
    Next lineNum
    lineNum = endLne + 1
End If


End Sub

Public Sub RunProg()

Dim tmpLine, str As String
Dim a As Integer

'Read and cut out sub definitions
For a = runCode.itemCount To 1 Step -1
  If LCase(left(runCode.Item(a), 4)) = "sub " Then
    AddSubDef a
  End If
Next a

'Read and cut out function definitions
For a = runCode.itemCount To 1 Step -1
  If LCase(left(runCode.Item(a), 9)) = "function " Then
    AddFuncDef a
  End If
Next a

'Define system variables
DefineSysVars

'Define all other variables in the code
For a = runCode.itemCount To 1 Step -1
    If LCase(left(runCode.Item(a), 4)) = "var " Then
        Cmd_Var Mid(runCode.Item(a), 5)
        runCode.Remove a
    End If
Next a

'Read and cut out data commands
For a = runCode.itemCount To 1 Step -1
  If LCase(left(runCode.Item(a), 5)) = "data " Then
    Cmd_Data Mid(runCode.Item(a), 6)
    runCode.Remove a
  End If
Next a

'Read branch labels (AFTER command cut-outs to insure accurate line numbers)
For a = 1 To runCode.itemCount
  If left(runCode.Item(a), 1) = "@" Then
    labelLine.Add a
    labelName.Add runCode.Item(a)
  End If
Next a

'If debugging, add code to debugger display
If debugging Then
    'debugCode.Add New ArrayClass, 1
    'For a = 1 To runCode.itemCount
    '    debugWin.code.AddItem runCode.Item(a)
    '    debugCode.Item(1).Add runCode.Item(a)
    'Next a
    'debugWin.code.AddItem ""
    debugWin.stack.AddItem "<main>"
    debugWin.Caption = "Debug - <main>"
    DebugUpdateCode
    debugWin.code.ListIndex = 0
    DebugUpdateVars
End If

progDone = False
gosubLine = 0
onErrorCmd = ""
readPos = 1
errorFlag = False

'Register the window class
RegClass "MicroByteWin"
RegClass "MBGraphWin"

'Loop through every line, running each command within the line
For lineNum = 1 To runCode.itemCount
    DebugWait
    If Not RunBlock(lineNum, runCode.itemCount) Then
        RunCmd runCode.Item(lineNum)
    End If
    If progDone Then Exit For
    If debugging Then debugWin.code.ListIndex = lineNum
Next lineNum


If debugging Then debugWin.Caption = "Debug Finished"

For n = 1 To fileHandle.itemCount
    Close fileNumber.Item(n)
Next n

For n = 1 To timerID.itemCount
    KillTimer 0, timerID.Item(n)
Next n

For n = windows.itemCount To 1 Step -1
    DestroyWindow windows.Item(n).winHandle
Next n

For n = 1 To imgName.itemCount
    DeleteObject imgHandle.Item(n)
Next n


End Sub

Public Function GetFuncPtr(ByVal funcPtr As Long) As Long

GetFuncPtr = funcPtr

End Function
Private Function EvalExpression(ByVal expStr As String) As Variant

Dim operand As New ArrayClass
Dim operator As New ArrayClass
Dim tmpExp As String
Dim inString As Boolean

If Trim(expStr) = "" Then EvalExpression = "": Exit Function

'parse out each operand
For I = 1 To Len(expStr)

  If Mid(expStr, I, 1) = Chr(34) Then
    If inString = False Then inString = True Else inString = False
  End If
  
  If Mid(expStr, I, 1) = "(" And inString = False Then
    temp = GetString(I + 1, expStr, ")")
    tmpExp = tmpExp & "(" & temp & ")"
    I = I + Len(temp) + 1
  
  ElseIf (Mid(expStr, I, 2) = ">=" Or Mid(expStr, I, 2) = "<=" Or Mid(expStr, I, 2) = "<>") _
  And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add Mid(expStr, I, 2), 1: I = I + 1
    
  ElseIf (Mid(expStr, I, 1) = "+" Or Mid(expStr, I, 1) = "%" _
  Or Mid(expStr, I, 1) = "*" Or Mid(expStr, I, 1) = "/" Or Mid(expStr, I, 1) = "^" _
  Or Mid(expStr, I, 1) = "=" Or Mid(expStr, I, 1) = "<" Or Mid(expStr, I, 1) = ">" _
  Or Mid(expStr, I, 1) = "&") And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add Mid(expStr, I, 1), 1
  
  ElseIf (UCase(Mid(expStr, I, 5)) = " AND " Or UCase(Mid(expStr, I, 5)) = " XOR ") _
  And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add UCase(Trim(Mid(expStr, I, 5))), 1: I = I + 4
  
  ElseIf (UCase(Mid(expStr, I, 4)) = " OR ") And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add UCase(Trim(Mid(expStr, I, 4))), 1: I = I + 3
  
  ElseIf Mid(expStr, I, 1) = "-" And inString = False Then
    If Len(Trim(tmpExp)) = 0 Then
      tmpExp = tmpExp & "-"
    Else
      GoSub AddType: tmpExp = ""
      operator.Add "-", 1
    End If
        
  Else
    tmpExp = tmpExp & Mid(expStr, I, 1)
  End If

Next

    GoSub AddType


'Evaluate operator: ^
For a = operator.itemCount To 1 Step -1
  If operator.Item(a) = "^" Then
    operand.Item(a) = operand.Item(a + 1) ^ operand.Item(a)
    operand.Remove (a + 1): operator.Remove (a)
  End If
Next a

'Evaluate operators: *, /, and %
For a = operator.itemCount To 1 Step -1
  If operator.Item(a) = "*" Then
    operand.Item(a) = operand.Item(a + 1) * operand.Item(a)
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = "/" Then
    operand.Item(a) = operand.Item(a + 1) / operand.Item(a)
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = "%" Then
    operand.Item(a) = operand.Item(a + 1) Mod operand.Item(a)
    operand.Remove (a + 1): operator.Remove (a)
  End If
Next a

'Evaluate operators: + and -
For a = operator.itemCount To 1 Step -1
  If operator.Item(a) = "+" Then
    operand.Item(a) = operand.Item(a + 1) + operand.Item(a)
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = "-" Then
    operand.Item(a) = operand.Item(a + 1) - operand.Item(a)
    operand.Remove (a + 1): operator.Remove (a)
  End If
Next a

'Evaluate operator: &
For a = operator.itemCount To 1 Step -1
  If operator.Item(a) = "&" Then
    operand.Item(a) = operand.Item(a + 1) & operand.Item(a)
    operand.Remove (a + 1): operator.Remove (a)
  End If
Next a

'Evaluate operators: =, <, >, <=, >=, and <>
For a = operator.itemCount To 1 Step -1
  If operator.Item(a) = "=" Then
    operand.Item(a) = CDbl(operand.Item(a + 1) = operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = "<" Then
    operand.Item(a) = CDbl(operand.Item(a + 1) < operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = ">" Then
    operand.Item(a) = CDbl(operand.Item(a + 1) > operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = "<=" Then
    operand.Item(a) = CDbl(operand.Item(a + 1) <= operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = ">=" Then
    operand.Item(a) = CDbl(operand.Item(a + 1) >= operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = "<>" Then
    operand.Item(a) = CDbl(operand.Item(a + 1) <> operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  End If
Next a

'Evaluate operators: AND, OR, and XOR
For a = operator.itemCount To 1 Step -1
  If operator.Item(a) = "AND" Then
    operand.Item(a) = (operand.Item(a + 1) And operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = "OR" Then
    operand.Item(a) = (operand.Item(a + 1) Or operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  ElseIf operator.Item(a) = "XOR" Then
    operand.Item(a) = (operand.Item(a + 1) Xor operand.Item(a))
    operand.Remove (a + 1): operator.Remove (a)
  End If
Next a

EvalExpression = operand.Item(1)

Exit Function


'routine to add an operand to the stack
AddType:
    tmpExp = Trim(tmpExp)
    If IsNumeric(tmpExp) Then
          operand.Add CDbl(tmpExp), 1
    ElseIf (left(tmpExp, 1) = Chr(34)) And (Right(tmpExp, 1) = Chr(34)) Then
      operand.Add Mid(tmpExp, 2, Len(tmpExp) - 2), 1
    ElseIf (left(tmpExp, 1) = "(") And (Right(tmpExp, 1) = ")") Then
      operand.Add EvalExpression(Mid(tmpExp, 2, Len(tmpExp) - 2)), 1
    Else
        For a = 1 To varName.itemCount
          If tmpExp = varName.Item(a) Then
              operand.Add varValue.Item(a), 1
              Return
          End If
        Next a
        For a = 1 To arrayName.itemCount
            If left(tmpExp, Len(arrayName.Item(a))) = arrayName.Item(a) And (Right(tmpExp, 1) = ")") Then
                idxStr = GetString(Len(GetString(1, tmpExp, "(")) + 2, tmpExp, ")")
                firstIdx = GetString(1, idxStr, ",")
                firstVal = EvalExpression(firstIdx)
                If firstVal < 0 Or (firstVal + 1) > arrayValue.Item(a).itemCount Then
                    ErrorMsg "Illegal index value"
                    Exit Function
                End If
                If Len(firstIdx) < Len(idxStr) Then
                    If Not isMultiDim.Item(a) Then
                        ErrorMsg "Array is one dimensional"
                        Exit Function
                    End If
                    secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
                    secondVal = EvalExpression(secondIdx)
                    If secondVal < 0 Or (secondVal + 1) > arrayValue.Item(a).Item(1).itemCount Then
                        ErrorMsg "Illegal index value"
                        Exit Function
                    End If
                    If arrayType.Item(a) = DT_NUMBER Then
                        operand.Add arrayValue.Item(a).Item(firstVal + 1).Item(secondVal + 1), 1
                    Else
                        operand.Add arrayValue.Item(a).Item(firstVal + 1).Item(secondVal + 1), 1
                    End If
                Else
                    If isMultiDim.Item(a) Then
                        ErrorMsg "Array is two dimensional"
                        Exit Function
                    End If
                    If arrayType.Item(a) = DT_NUMBER Then
                        operand.Add arrayValue.Item(a).Item(firstVal + 1), 1
                    Else
                        operand.Add arrayValue.Item(a).Item(firstVal + 1), 1
                    End If
                End If
                Return
            End If
        Next a
        For b = 1 To strFunc.itemCount
          If LCase(left(tmpExp, Len(strFunc.Item(b)))) = LCase(strFunc.Item(b)) Then
              operand.Add EvalFunction(tmpExp), 1
              Return
          End If
        Next b
        For C = 1 To numFunc.itemCount
          If LCase(left(tmpExp, Len(numFunc.Item(C)))) = LCase(numFunc.Item(C)) Then
              operand.Add CDbl(EvalFunction(tmpExp)), 1
              Return
          End If
        Next C
        For C = 1 To funcName.itemCount
          If left(tmpExp, Len(funcName.Item(C))) = funcName.Item(C) Then
              operand.Add CallUserFunc(tmpExp), 1
              Return
          End If
        Next C
    End If
Return

End Function
Private Function EvalFunction(funcStr As String) As Variant

Dim paramStr As String

funcStr = Trim(funcStr)

If LCase(left(funcStr, 4)) = "abs(" Then
    EvalFunction = Func_Abs(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "asc(" Then
    EvalFunction = Func_Asc(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "not(" Then
    EvalFunction = Func_Not(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "int(" Then
    EvalFunction = Func_Int(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "len(" Then
    EvalFunction = Func_Len(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "rnd(" Then
    EvalFunction = Func_Rnd()
ElseIf LCase(left(funcStr, 4)) = "val(" Then
    EvalFunction = Func_Val(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "chr(" Then
    EvalFunction = Func_Chr(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "str(" Then
    EvalFunction = Func_Str(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 6)) = "upper(" Then
    EvalFunction = Func_Upper(GetString(7, funcStr, ")"))
ElseIf LCase(left(funcStr, 6)) = "lower(" Then
    EvalFunction = Func_Lower(GetString(7, funcStr, ")"))
ElseIf LCase(left(funcStr, 5)) = "trim(" Then
    EvalFunction = Func_Trim(GetString(6, funcStr, ")"))
ElseIf LCase(left(funcStr, 5)) = "left(" Then
    EvalFunction = Func_Left(GetString(6, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "mid(" Then
    EvalFunction = Func_Mid(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 6)) = "right(" Then
    EvalFunction = Func_Right(GetString(7, funcStr, ")"))
ElseIf LCase(left(funcStr, 6)) = "instr(" Then
    EvalFunction = Func_InStr(GetString(7, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "loc(" Then
    EvalFunction = Func_Loc(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 8)) = "gettext(" Then
    EvalFunction = Func_GetText(GetString(9, funcStr, ")"))
ElseIf LCase(left(funcStr, 5)) = "hwnd(" Then
    EvalFunction = Func_hWnd(GetString(6, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "hdc(" Then
    EvalFunction = Func_hDC(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 10)) = "getselidx(" Then
    EvalFunction = Func_GetSelIdx(GetString(11, funcStr, ")"))
ElseIf LCase(left(funcStr, 8)) = "getitem(" Then
    EvalFunction = Func_GetItem(GetString(9, funcStr, ")"))
ElseIf LCase(left(funcStr, 10)) = "itemcount(" Then
    EvalFunction = Func_ItemCount(GetString(11, funcStr, ")"))
ElseIf LCase(left(funcStr, 11)) = "getseltext(" Then
    EvalFunction = Func_GetSelText(GetString(12, funcStr, ")"))
ElseIf LCase(left(funcStr, 10)) = "linecount(" Then
    EvalFunction = Func_LineCount(GetString(11, funcStr, ")"))
ElseIf LCase(left(funcStr, 12)) = "getlinetext(" Then
    EvalFunction = Func_GetLineText(GetString(13, funcStr, ")"))
ElseIf LCase(left(funcStr, 17)) = "getclipboardtext(" Then
    EvalFunction = Func_GetClipboardText()
ElseIf LCase(left(funcStr, 9)) = "inputbox(" Then
    EvalFunction = Func_InputBox(GetString(10, funcStr, ")"))
ElseIf LCase(left(funcStr, 5)) = "date(" Then
    EvalFunction = Func_Date()
ElseIf LCase(left(funcStr, 5)) = "time(" Then
    EvalFunction = Func_Time()
ElseIf LCase(left(funcStr, 9)) = "getstate(" Then
    EvalFunction = Func_GetState(GetString(10, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "min(" Then
    EvalFunction = Func_Min(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "max(" Then
    EvalFunction = Func_Max(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "sqr(" Then
    EvalFunction = Func_Sqr(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 6)) = "space(" Then
    EvalFunction = Func_Space(GetString(7, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "lof(" Then
    EvalFunction = Func_LOF(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 4)) = "eof(" Then
    EvalFunction = Func_EOF(GetString(5, funcStr, ")"))
ElseIf LCase(left(funcStr, 5)) = "hbmp(" Then
    EvalFunction = Func_hBmp(GetString(6, funcStr, ")"))
ElseIf LCase(left(funcStr, 9)) = "fileopen(" Then
    EvalFunction = Func_FileOpen(GetString(10, funcStr, ")"))
ElseIf LCase(left(funcStr, 9)) = "filesave(" Then
    EvalFunction = Func_FileSave(GetString(10, funcStr, ")"))
ElseIf LCase(left(funcStr, 8)) = "collide(" Then
    EvalFunction = Func_Collide(GetString(9, funcStr, ")"))
ElseIf LCase(left(funcStr, 6)) = "input(" Then
    EvalFunction = Func_Input(GetString(7, funcStr, ")"))
End If

End Function


Public Function Func_Min(ByVal paramStr As String) As Variant

Dim num1, num2 As Long
Dim params As New ArrayClass

paramStr = Trim(paramStr)

ParseParams paramStr, params

num1 = EvalExpression(params.Item(1))
num2 = EvalExpression(params.Item(2))

Func_Min = Min(num1, num2)

End Function


Public Function Func_Max(ByVal paramStr As String) As Variant

Dim num1, num2 As Long
Dim params As New ArrayClass

paramStr = Trim(paramStr)

ParseParams paramStr, params

num1 = EvalExpression(params.Item(1))
num2 = EvalExpression(params.Item(2))

Func_Max = Max(num1, num2)

End Function



Public Function Func_Sqr(ByVal paramStr As String) As Variant

Dim num As Long

paramStr = Trim(paramStr)

num = EvalExpression(paramStr)

Func_Sqr = Sqr(num)

End Function


Public Function Func_Space(ByVal paramStr As String) As Variant

Dim num As Long

paramStr = Trim(paramStr)

num = EvalExpression(paramStr)

Func_Space = Space(num)

End Function



Sub Main()

Dim idxCount As Integer
Dim lineTxt As String

ReadRunFile

LoadFunctions

App.Title = App.EXEName
Output.Caption = "Running: " & App.EXEName

debugging = False

RunProg

Output.Caption = "Execution complete: " & App.EXEName
Output.display.SelStart = Len(Output.display.Text)

If (Not Output.Visible) Or (errorFlag) Then End

End Sub


Public Sub OldRunSelectBlock(ByVal startLne As Integer, ByVal endLne As Integer)

Dim expStr As String
Dim selectVal As Variant
Dim caseIsTrue As Boolean

selectVal = EvalExpression(Mid(runCode.Item(startLne), 13))
caseIsTrue = False

If endLne > startLne Then
  For lineNum = startLne + 1 To endLne
      If debugging Then
          If debugState = DS_STEP Then
              debugState = DS_PAUSE
          ElseIf debugState = DS_PAUSE Then
              While debugState = DS_PAUSE
                  DoEvents
                  Sleep 1
              Wend
              If debugState = DS_STEP Then debugState = DS_PAUSE
          End If
          debugWin.code.ListIndex = lineNum - 1
      End If
      If LCase(Trim(runCode.Item(lineNum))) = "case else" Then
          If caseIsTrue Then
            lineNum = endLne + 1
            Exit Sub
          Else
            caseIsTrue = True
          End If
      ElseIf LCase(left(runCode.Item(lineNum), 5)) = "case " Then
          If caseIsTrue Then
            lineNum = endLne + 1
            Exit Sub
          Else
            expStr = Mid(runCode.Item(lineNum), 6)
            If selectVal = EvalExpression(expStr) Then
              caseIsTrue = True
            End If
          End If
      Else
          If caseIsTrue Then
            If Not RunBlock(lineNum, endLne) Then
                RunCmd runCode.Item(lineNum)
            End If
            If progDone Then Exit Sub
          End If
      End If
      If lineNum < startLne Or lineNum > endLne Then Exit Sub
  Next lineNum
End If

lineNum = endLne + 1


End Sub
Public Sub RunSelectBlock(ByVal startLne As Integer, ByVal endLne As Integer)

Dim expStr As String
Dim caseExp As New ArrayClass
Dim caseStart As New ArrayClass
Dim caseEnd As New ArrayClass
Dim selectNum, n As Integer
Dim selectVal As Variant

selectVal = EvalExpression(Mid(runCode.Item(startLne), 13))

If endLne > startLne Then
  caseExp.Add "1"
  caseStart.Add startLne + 1
  For n = startLne + 1 To endLne
      If LCase(left(runCode.Item(n), 12)) = "select case " Then
          selectNum = selectNum + 1
      ElseIf LCase(Trim(runCode.Item(n))) = "end select" Then
          selectNum = selectNum - 1
      ElseIf LCase(Trim(runCode.Item(n))) = "case else" And selectNum = 0 Then
          caseEnd.Add n - 1
          caseExp.Add selectVal
          caseStart.Add n + 1
      ElseIf LCase(left(runCode.Item(n), 5)) = "case " And selectNum = 0 Then
          caseEnd.Add n - 1
          caseExp.Add Mid(runCode.Item(n), 6)
          caseStart.Add n + 1
      End If
  Next n
  
  caseEnd.Add endLne
  
  If debugging Then debugWin.code.ListIndex = lineNum
  DebugWait
  
  For n = 2 To caseExp.itemCount
      If EvalExpression(caseExp.Item(n)) = selectVal Then
          If debugging Then debugWin.code.ListIndex = caseStart.Item(n) - 1
          For lineNum = caseStart.Item(n) To caseEnd.Item(n)
              DebugWait
              If Not RunBlock(lineNum, endLne) Then
                  RunCmd runCode.Item(lineNum)
              End If
              If progDone Then Exit Sub
              If lineNum < startLne Or lineNum > endLne Then Exit Sub
              If debugging Then debugWin.code.ListIndex = lineNum
          Next lineNum
          lineNum = endLne + 1
          Exit Sub
      End If
      If debugging Then debugWin.code.ListIndex = caseEnd.Item(n)
      DebugWait
  Next n
End If

lineNum = endLne + 1


End Sub


Public Sub RunWhileBlock(ByVal startLne As Integer, ByVal endLne As Integer)

Dim str, expStr, tmpLine As String

expStr = Right(runCode.Item(startLne), Len(runCode.Item(startLne)) - 6)

While EvalExpression(expStr)
  If debugging Then debugWin.code.ListIndex = startLne
  If endLne > startLne Then
    For lineNum = startLne + 1 To endLne
        DebugWait
        If Not RunBlock(lineNum, endLne) Then
            RunCmd runCode.Item(lineNum)
        End If
        If progDone Then Exit Sub
        If lineNum < startLne Or lineNum > endLne Then Exit Sub
        If debugging Then debugWin.code.ListIndex = lineNum
    Next lineNum
  End If
  DebugWait
  If debugging Then debugWin.code.ListIndex = startLne - 1
  DebugWait
Wend

lineNum = endLne + 1


End Sub


Public Sub SetValue(ByVal varStr As String, ByVal value As Variant)

varStr = Trim(varStr)

For n = 1 To varName.itemCount
    If varName.Item(n) = varStr Then
        If varType.Item(n) = DT_NUMBER Then
            varValue.Item(n) = val(value)
            For a = 1 To varBindList.Item(n).itemCount
                varValue.Item(varBindList.Item(n).Item(a)) = val(value)
            Next a
        Else
            varValue.Item(n) = value
            For a = 1 To varBindList.Item(n).itemCount
                varValue.Item(varBindList.Item(n).Item(a)) = value
            Next a
        End If
        If debugging Then DebugUpdateVars
        Exit Sub
    End If
Next n

For n = 1 To arrayName.itemCount
    If left(varStr, Len(arrayName.Item(n))) = arrayName.Item(n) And (Right(varStr, 1) = ")") Then
        idxStr = GetString(Len(GetString(1, varStr, "(")) + 2, varStr, ")")
        firstIdx = GetString(1, idxStr, ",")
        firstVal = EvalExpression(firstIdx)
        If firstVal < 0 Or (firstVal + 1) > arrayValue.Item(n).itemCount Then
            ErrorMsg "Illegal index value"
            Exit Sub
        End If
        If Len(firstIdx) < Len(idxStr) Then
            If Not isMultiDim.Item(n) Then
                ErrorMsg "Array is one dimensional"
                Exit Sub
            End If
            secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
            secondVal = EvalExpression(secondIdx)
            If secondVal < 0 Or (secondVal + 1) > arrayValue.Item(n).Item(1).itemCount Then
                ErrorMsg "Illegal index value"
                Exit Sub
            End If
            If arrayType.Item(n) = DT_NUMBER Then
                arrayValue.Item(n).Item(firstVal + 1).Item(secondVal + 1) = val(value)
            Else
                arrayValue.Item(n).Item(firstVal + 1).Item(secondVal + 1) = value
            End If
        Else
            If isMultiDim.Item(n) Then
                ErrorMsg "Array is two dimensional"
                Exit Sub
            End If
            If arrayType.Item(n) = DT_NUMBER Then
                arrayValue.Item(n).Item(firstVal + 1) = val(value)
            Else
                arrayValue.Item(n).Item(firstVal + 1) = value
            End If
        End If
        Exit Sub
    End If
Next n


End Sub

Public Sub Cmd_Open(ByVal cmdStr As String)

Dim tmpFile, tmpType, tmpHandle As String
Dim a, tmpFileNum As Integer

    cmdStr = Trim(cmdStr)

'Parse out file path
tmpFile = GetString(1, LCase(cmdStr), " for ")
tmpFile = left(cmdStr, Len(tmpFile))
a = Len(tmpFile) + 6

'Parse out file type
tmpType = GetString(a, LCase(cmdStr), " as ")
tmpType = Mid(cmdStr, a, Len(tmpType))
a = a + Len(tmpType) + 3
tmpType = Trim(LCase(tmpType))

'Parse out file handle
tmpHandle = Trim(Right(cmdStr, Len(cmdStr) - a))

'Check to see if handle is already in use
For a = 1 To fileHandle.itemCount
    If fileHandle.Item(a) = tmpHandle Then
        ErrorMsg "File handle already in use: " & tmpHandle
        Exit Sub
    End If
Next a

'Get the next free file number
tmpFileNum = FreeFile

'Add new file handle and file number
fileHandle.Add tmpHandle
fileNumber.Add tmpFileNum

'Evaluate file path string expression
tmpFile = EvalExpression(tmpFile)

'Open file
Select Case tmpType
    Case "input"
        'Check to see if file exists
        If Dir(tmpFile) = "" Then
            ErrorMsg "File does not exists: " & tmpFile
            Exit Sub
        End If
        fileType.Add FT_INPUT
        Open tmpFile For Input As tmpFileNum
    Case "output"
        fileType.Add FT_OUTPUT
        Open tmpFile For Output As tmpFileNum
    Case "append"
        fileType.Add FT_APPEND
        Open tmpFile For Append As tmpFileNum
    Case "binary"
        fileType.Add FT_BINARY
        Open tmpFile For Binary As tmpFileNum
End Select


End Sub

Public Sub Cmd_Close(ByVal cmdStr As String)

Dim a As Integer

cmdStr = Trim(cmdStr)

For a = 1 To fileHandle.itemCount
    If fileHandle.Item(a) = cmdStr Then
        Close fileNumber.Item(a)
        fileHandle.Remove a
        fileNumber.Remove a
        fileType.Remove a
        Exit Sub
    End If
Next a

ErrorMsg "File handle does not exist: " & cmdStr

End Sub
Public Sub TimerProc(ByVal handle As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)

Dim subObj As SubProgClass
Dim tmpName As String

For a = 1 To timerID.itemCount
    If idEvent = timerID.Item(a) Then
        CallSubProg timerSubIdx.Item(a), timerSubType.Item(a)
    End If
Next a

End Sub


