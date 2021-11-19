VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompilerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cmdList As New ArrayClass
Public txtLneNum As New ArrayClass
Public lineNum As Integer


Public tmpVarName As New ArrayClass
Public tmpVarType As New ArrayClass

Public arrayName As New ArrayClass
Public arrayType As New ArrayClass

Public labelName As New ArrayClass


Public subProg_LabelName As New ArrayClass

Public subProg_VarName As New ArrayClass
Public subProg_VarType As New ArrayClass

Public subProg_ArrayName As New ArrayClass
Public subProg_ArrayType As New ArrayClass

Public inSubProg As Boolean
Public subProgType As Integer

Public subName As New ArrayClass
Public subParamType As New ArrayClass

Public funcName As New ArrayClass
Public funcType As New ArrayClass
Public funcParamType As New ArrayClass



Public Sub Cmd_AddItem(ByVal cmdStr As String)

Dim paramType As Integer
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If (params.itemCount < 2) Or (params.itemCount > 3) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

EvalExpression params.Item(2)
If compileError Then Exit Sub

If params.itemCount = 2 Then Exit Sub

paramType = EvalExpression(params.Item(3))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_BackColor(ByVal cmdStr As String)

Dim paramType As Integer
Dim params As New ArrayClass

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_Box(ByVal cmdStr As String)

Dim paramType, n As Integer
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 5 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 5
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_ButtonImg(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub

Public Sub Cmd_Circle(ByVal cmdStr As String)

Dim paramType, n As Integer
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 4 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 4
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_Clear(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_CloseSound(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_ExitFunction()

If Not inSubProg Or subProgType <> SP_FUNC Then
    ErrorMsg "EXIT FUNCTION cannot be used outside of a FUNCTION"
End If

End Sub

Public Sub Cmd_ExitSub()

If Not inSubProg Or subProgType <> SP_SUB Then
    ErrorMsg "EXIT SUB cannot be used outside of a SUB"
End If

End Sub


Public Sub Cmd_GetFiles(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

If Right(params.Item(2), 2) = ")" Then
    params.Item(2) = Trim(Left(params.Item(2), Len(params.Item(2)) - 2))
End If

SetValue params.Item(2) & "0)", Chr(34) & Chr(34)

End Sub


Public Sub Cmd_GetDirs(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

If Right(params.Item(2), 2) = ")" Then
    params.Item(2) = Trim(Left(params.Item(2), Len(params.Item(2)) - 2))
End If

SetValue params.Item(2) & "0)", Chr(34) & Chr(34)

End Sub



Public Sub Cmd_GetMouseXY(ByVal cmdStr As String)

Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    SetValue params.Item(n), "0"
    If compileError Then Exit Sub
Next n

End Sub

Public Sub Cmd_Kill(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_Mid(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount < 3 Or params.itemCount > 4 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

SetValue params.Item(1), Chr(34) & Chr(34)
If compileError Then Exit Sub

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

If params.itemCount = 3 Then
    paramType = EvalExpression(params.Item(3))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Else
    paramType = EvalExpression(params.Item(3))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
    paramType = EvalExpression(params.Item(4))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then ErrorMsg "Type mismatch"
End If

End Sub

Public Sub Cmd_MkDir(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_RmDir(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_SetState(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error'"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_SpriteBGPos(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

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


Public Sub Cmd_Name(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

params.Add GetString(1, LCase(cmdStr), " as ")
params.Item(1) = Mid(cmdStr, 1, Len(params.Item(1)))

If Len(params.Item(1)) = Len(cmdStr) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

params.Add Mid(cmdStr, Len(params.Item(1)) + 5)

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub



Public Sub Cmd_OpenSound(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_PauseSound(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_PlaySound(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_StopSound(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_DelFrame(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

paramType = EvalExpression(params.Item(3))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_DelItem(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_DrawImg(ByVal cmdStr As String)

Dim paramType, n As Integer
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 4 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

For n = 3 To 4
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Syntax error"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_DrawText(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 4 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub

For n = 3 To 4
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub

Public Sub Cmd_ForeColor(ByVal cmdStr As String)

Dim paramType As Integer
Dim params As New ArrayClass

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_GetImg(ByVal cmdStr As String)

Dim paramType, n As Integer
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 6 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

For n = 3 To 6
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Syntax error"
        Exit Sub
    End If
Next n

End Sub
Public Sub Cmd_Enable(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_GetSel(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 3
    SetValue params.Item(n), "0"
    If compileError Then Exit Sub
Next n

End Sub



Public Sub Func_Collide(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Wrong number of arguments in COLLIDE()"
    Exit Sub
End If

For n = 1 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch in parameter " & n & " of COLLIDE()"
        Exit Sub
    End If
Next n

End Sub

Public Sub Func_EOF(ByVal paramStr As String)

paramStr = Trim(paramStr)

If Left(paramStr, 1) <> "#" Then
    ErrorMsg "File handle does not exist: " & paramStr
End If

End Sub


Public Sub Func_FileOpen(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Wrong number of arguments in FILEOPEN()"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch in parameter " & n & " of FILEOPEN()"
        Exit Sub
    End If
Next n

End Sub


Public Sub Func_FileSave(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Wrong number of arguments in FILESAVE()"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch in parameter " & n & " of FILEOPEN()"
        Exit Sub
    End If
Next n

End Sub



Public Sub Func_GetSoundPos(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of GETSOUNDPOS()"

End Sub


Public Sub Func_GetSoundLen(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of GETSOUNDLEN()"

End Sub

Public Sub Func_GetState(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of GETSTATE()"

End Sub

Public Sub Func_hBmp(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of HBMP()"

End Sub


Public Sub Func_Input(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Wrong number of arguments in INPUT()"
    Exit Sub
End If

If Left(params.Item(1), 1) <> "#" Then
    ErrorMsg "File handle does not exist: " & params.Item(1)
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 2 of INPUT()"

End Sub

Public Sub Func_InputBox(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If (params.itemCount < 2) Or (params.itemCount > 3) Then
    ErrorMsg "Wrong number of arguments in INPUTBOX()"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch in parameter " & n & " of INPUTBOX()"
        Exit Sub
    End If
Next n

If params.itemCount = 2 Then Exit Sub

EvalExpression (params.Item(3))
If compileError Then Exit Sub

End Sub

Public Sub Cmd_Line(ByVal cmdStr As String)

Dim paramType, n As Integer
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 5 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 5
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_LoadImg(ByVal cmdStr As String)

Dim paramType, n As Integer
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub



Public Sub Cmd_Refresh(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_SetClipboardText(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

paramType = EvalExpression(cmdStr)

End Sub

Public Sub Cmd_SetFont(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 4 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

For n = 3 To 4
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub

Public Sub Cmd_SetPixel(ByVal cmdStr As String)

Dim paramType, n As Integer
Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 4 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 4
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub



Public Sub Cmd_SetSel(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub

Public Sub Cmd_GetXY(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 3
    SetValue params.Item(n), "0"
    If compileError Then Exit Sub
Next n

End Sub
Public Sub Cmd_Hide(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_Menu(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n, a, b As Integer
Dim subStr As String

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount < 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

If params.itemCount = 2 Then Exit Sub

n = 2
While n < params.itemCount
    n = n + 1
    If params.Item(n) <> "|" Then
        paramType = EvalExpression(params.Item(n))
        If compileError Then Exit Sub
        If paramType <> DT_STRING Then
            ErrorMsg "Type mismatch"
            Exit Sub
        End If
        If n = params.itemCount Then
            ErrorMsg "Syntax error"
            Exit Sub
        End If
        If Not SubProgExists(params.Item(n + 1)) Then
            ErrorMsg "Undefined function or sub: '" & params.Item(n + 1) & "'"
            Exit Sub
        End If
        n = n + 1
    End If
Wend

End Sub




Public Sub Cmd_SetItem(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(3))
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_SetSelIdx(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub



Public Sub Cmd_SetSelText(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

EvalExpression params.Item(2)

End Sub


Public Sub Cmd_SetText(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)
    
If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_SetXY(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_SetSize(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub

Public Sub Cmd_GetSize(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

For n = 2 To 3
    SetValue params.Item(n), "0"
    If compileError Then Exit Sub
Next n

End Sub



Public Sub Cmd_Seek(ByVal cmdStr As String)

Dim handleStr, seekStr As String
Dim a, tmpType As Integer

cmdStr = Trim(cmdStr)

handleStr = GetString(1, cmdStr, ",")
seekStr = Mid(cmdStr, Len(handleStr) + 2)
handleStr = Trim(handleStr)

tmpType = EvalExpression(seekStr)
If compileError Then Exit Sub
If tmpType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

If Left(handleStr, 1) <> "#" Then
    ErrorMsg "Bad file handle: " & handleStr
End If

End Sub
Public Sub AddFuncDef(ByVal cmdStr As String)

'Add function name, type, and parameter types to
'the function definition collections

Dim tmpName, tmpParam, paramStr, varStr, typeStr As String

cmdStr = Trim(cmdStr)

'Get the function name
tmpName = GetString(1, cmdStr, "(")
For a = 1 To Len(tmpName)
    If (Asc(UCase(Mid(tmpName, a, 1))) < 65 Or Asc(UCase(Mid(tmpName, a, 1))) > 90) And _
    (Not (IsNumeric(Mid(tmpName, a, 1)))) And (Mid(tmpName, a, 1) <> "_") Then
        ErrorMsg "Illegal function name"
        Exit Sub
    End If
Next a
If Len(tmpName) = Len(cmdStr) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

'Get the parameter string
paramStr = GetString(Len(tmpName) + 1, LCase(cmdStr), " as ")
paramStr = Mid(cmdStr, Len(tmpName) + 1, Len(paramStr))
If Len(paramStr) = (Len(cmdStr) - (Len(tmpName) + 1)) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

'Get the function type
typeStr = Right(cmdStr, Len(cmdStr) - (Len(tmpName) + Len(paramStr)))

If Right(Trim(paramStr), 1) <> ")" Then
    ErrorMsg "Syntax error"
    Exit Sub
End If
paramStr = GetString(2, paramStr, ")")

typeStr = Trim(LCase(Right(typeStr, Len(typeStr) - 4)))
If typeStr = "string" Then
    funcType.Add DT_STRING
ElseIf typeStr = "number" Then
    funcType.Add DT_NUMBER
Else
    ErrorMsg "Illegal function type: " & typeStr
    Exit Sub
End If

funcName.Add tmpName & "("
funcParamType.Add New ArrayClass

b = 1
While b <= Len(paramStr)
    tmpParam = GetString(b, paramStr, ",")
    b = Len(tmpParam) + b + 1
    varStr = GetString(1, LCase(tmpParam), " as ")
    varStr = Left(tmpParam, Len(varStr))
    If Len(varStr) = Len(tmpParam) Then
        ErrorMsg "Syntax error"
        Exit Sub
    End If
    typeStr = LCase(Trim(Right(tmpParam, Len(tmpParam) - (Len(varStr) + 4))))
    If typeStr = "number" Then
        funcParamType.Item(funcParamType.itemCount).Add DT_NUMBER
    ElseIf typeStr = "string" Then
        funcParamType.Item(funcParamType.itemCount).Add DT_STRING
    Else
        ErrorMsg "Illegal parameter type: " & typeStr
        Exit Sub
    End If
Wend


End Sub
Public Sub AddSubDef(ByVal cmdStr As String)

'Add sub name and parameter types to
'the sub definition collections

Dim tmpName, tmpParam, varStr, typeStr As String

    cmdStr = Trim(cmdStr)

tmpName = GetString(1, cmdStr, " ")
For a = 1 To Len(tmpName)
    If (Asc(UCase(Mid(tmpName, a, 1))) < 65 Or Asc(UCase(Mid(tmpName, a, 1))) > 90) And _
    (Not (IsNumeric(Mid(tmpName, a, 1)))) And (Mid(tmpName, a, 1) <> "_") Then
        ErrorMsg "Illegal sub name"
        Exit Sub
    End If
Next a

subName.Add Trim(tmpName)
subParamType.Add New ArrayClass

b = Len(tmpName) + 2
While b <= Len(cmdStr)
    tmpParam = GetString(b, cmdStr, ",")
    b = Len(tmpParam) + b + 1
    varStr = GetString(1, LCase(tmpParam), " as ")
    varStr = Left(tmpParam, Len(varStr))
    If Len(varStr) = Len(tmpParam) Then
        ErrorMsg "Syntax error"
        Exit Sub
    End If
    typeStr = LCase(Trim(Right(tmpParam, Len(tmpParam) - (Len(varStr) + 4))))
    If typeStr = "number" Then
        subParamType.Item(subParamType.itemCount).Add DT_NUMBER
    ElseIf typeStr = "string" Then
        subParamType.Item(subParamType.itemCount).Add DT_STRING
    Else
        ErrorMsg "Illegal parameter type: " & typeStr
        Exit Sub
    End If
Wend


End Sub

Public Sub CallUserFunc(ByVal cmdStr As String)

Dim tmpName, tmpParam, paramStr As String
Dim tmpParamList As New ArrayClass

  For n = 1 To funcName.itemCount
    If Left(cmdStr, Len(funcName.Item(n))) = funcName.Item(n) Then
        tmpName = GetString(1, cmdStr, "(")
        paramStr = GetString(Len(tmpName) + 2, cmdStr, ")")
        b = 1
        While b <= Len(paramStr)
            tmpParam = GetString(b, paramStr, ",")
            b = Len(tmpParam) + b + 1
            tmpParamList.Add EvalExpression(tmpParam)
            If compileError Then Exit Sub
        Wend
        If tmpParamList.itemCount <> funcParamType.Item(n).itemCount Then
            ErrorMsg "Wrong number of arguments for function '" & tmpName & "'"
            Exit Sub
        End If
        For a = 1 To tmpParamList.itemCount
            If tmpParamList.Item(a) <> funcParamType.Item(n).Item(a) Then
                ErrorMsg "Type mismatch in argument " & a & " of call to function '" & tmpName & "'"
                Exit Sub
            End If
        Next a
        Exit Sub
    End If
  Next n

End Sub
Public Sub Cmd_Call(ByVal cmdStr As String)

Dim tmpName, tmpParam As String
Dim tmpParamList As New ArrayClass

cmdStr = Trim(cmdStr)

tmpName = GetString(1, cmdStr, " ")

For n = 1 To subName.itemCount
    If tmpName = subName.Item(n) Then
        b = Len(tmpName) + 2
        While b <= Len(cmdStr)
            tmpParam = GetString(b, cmdStr, ",")
            b = Len(tmpParam) + b + 1
            tmpParamList.Add EvalExpression(tmpParam)
            If compileError Then Exit Sub
        Wend
        If tmpParamList.itemCount <> subParamType.Item(n).itemCount Then
            ErrorMsg "Wrong number of arguments for sub '" & tmpName & "'"
            Exit Sub
        End If
        For a = 1 To tmpParamList.itemCount
            If tmpParamList.Item(a) <> subParamType.Item(n).Item(a) Then
                ErrorMsg "Type mismatch in argument " & a & " of call to sub '" & tmpName & "'"
                Exit Sub
            End If
        Next a
        Exit Sub
    End If
Next n
    
ErrorMsg "Sub '" & tmpName & "' not defined"

End Sub

Public Sub Cmd_Control(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim tmpParam As String
Dim paramType As Integer
Dim b As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

b = 1
While b <= Len(cmdStr)
    tmpParam = GetString(b, cmdStr, ",")
    b = Len(tmpParam) + b + 1
    params.Add Trim(tmpParam)
Wend

If params.itemCount <> 8 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For b = 1 To 3
    paramType = EvalExpression(params.Item(b))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next b

tmpParam = LCase(params.Item(4))
If Not ((tmpParam = "button") Or (tmpParam = "statictext") Or _
    (tmpParam = "texteditor") Or (tmpParam = "textbox") Or _
    (tmpParam = "listbox") Or (tmpParam = "combobox") Or _
    (tmpParam = "drawbox") Or (tmpParam = "picbutton") Or _
    (tmpParam = "checkbox") Or (tmpParam = "radiobutton") Or _
    (tmpParam = "groupbox")) Then
        ErrorMsg "Illegal control type: '" & tmpParam & "'"
        Exit Sub
End If

For b = 5 To 8
    paramType = EvalExpression(params.Item(b))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next b


End Sub
Public Sub Cmd_CloseWindow(ByVal cmdStr As String)

Dim paramType As Integer

paramType = EvalExpression(Trim(cmdStr))
If compileError Then Exit Sub

If paramType <> DT_STRING Then ErrorMsg "Type mismatch"


End Sub
Public Sub Cmd_Data(ByVal cmdStr As String)

Dim dataItems As New ArrayClass
Dim a As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, dataItems

For a = 1 To dataItems.itemCount
    EvalExpression dataItems.Item(a)
    If compileError Then Exit Sub
Next a

End Sub

Public Sub Cmd_Event(ByVal cmdStr As String)

Dim paramStr As String
Dim paramType, b, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramStr = GetString(1, cmdStr, ",")
If Len(paramStr) = Len(cmdStr) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If
paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

b = Len(paramStr) + 2

paramStr = GetString(b, cmdStr, ",")
If (b - 1) + Len(paramStr) = Len(cmdStr) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

b = b + Len(paramStr) + 1

paramStr = Trim(Mid(cmdStr, b))
If Not SubProgExists(paramStr) Then ErrorMsg "Undefined function or sub: '" & paramStr & "'"

End Sub

Public Sub Cmd_Message(ByVal cmdStr As String)

Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

EvalExpression (params.Item(1))
If compileError Then Exit Sub

EvalExpression (params.Item(2))

End Sub


Public Sub Cmd_Question(ByVal cmdStr As String)

Dim params As New ArrayClass

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

EvalExpression (params.Item(1))
If compileError Then Exit Sub

EvalExpression (params.Item(2))
If compileError Then Exit Sub

SetValue params.Item(3), Chr(34) & Chr(34)

End Sub
Public Sub Cmd_On(ByVal cmdStr As String)

Dim valueStr, labelStr, tmpLabel As String
Dim valueType As Integer
Dim a, b As Integer

    cmdStr = Trim(cmdStr)

'Get the value if it is an ON...GOTO command
valueStr = GetString(1, LCase(cmdStr), " goto ")
valueStr = Left(cmdStr, Len(valueStr))
b = Len(valueStr) + 6

'If not, then get the value if it is an ON...GOSUB command
If Len(valueStr) = Len(cmdStr) Then
    valueStr = GetString(1, LCase(cmdStr), " gosub ")
    valueStr = Left(cmdStr, Len(valueStr))
    b = Len(valueStr) + 7
'If it is neither, give an error
    If Len(valueStr) = Len(cmdStr) Then
        ErrorMsg "Syntax error"
        Exit Sub
    End If
End If

'Evalute the value
valueType = EvalExpression(valueStr)
If compileError Then Exit Sub
If valueType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

'Get list of branch labels
labelStr = Trim(Right(cmdStr, Len(cmdStr) - b))

'Section out and check each branch label
b = 1
While b <= Len(labelStr)
    tmpLabel = GetString(b, labelStr, ",")
    b = Len(tmpLabel) + b + 1
    tmpLabel = Trim(tmpLabel)
    If inSubProg Then
        For a = 1 To labelName.itemCount
            If tmpLabel = labelName.Item(a) Then
                Exit For
            ElseIf a = labelName.itemCount Then
                ErrorMsg "Branch label '" & tmpLabel & "' does not exist"
                Exit Sub
            End If
        Next a
    Else
        For a = 1 To subProg_LabelName.itemCount
            If tmpLabel = subProg_LabelName.Item(a) Then
                Exit For
            ElseIf a = subProg_LabelName.itemCount Then
                ErrorMsg "Branch label '" & tmpLabel & "' does not exist"
                Exit Sub
            End If
        Next a
    End If
Wend


End Sub
Public Sub Cmd_OnError(ByVal cmdStr As String)

    cmdStr = Trim(cmdStr)

CompileCmd cmdStr


End Sub
Public Sub Cmd_Open(ByVal cmdStr As String)

Dim tmpFile, tmpMode, tmpHandle As String
Dim tmpType As Integer
Dim a As Integer

cmdStr = Trim(cmdStr)

tmpFile = GetString(1, LCase(cmdStr), " for ")
tmpFile = Left(cmdStr, Len(tmpFile))

If Len(tmpFile) = Len(cmdStr) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

tmpType = EvalExpression(tmpFile)
If compileError Then Exit Sub
If tmpType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

a = Len(tmpFile) + 6

tmpMode = GetString(a, LCase(cmdStr), " as ")
tmpMode = Mid(cmdStr, a, Len(tmpMode))

a = a - 1

If Len(tmpMode) + a = Len(cmdStr) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

If Trim(LCase(tmpMode)) <> "input" And Trim(LCase(tmpMode)) <> "output" And _
Trim(LCase(tmpMode)) <> "append" And Trim(LCase(tmpMode)) <> "binary" And _
Trim(LCase(tmpMode)) <> "random" Then
    ErrorMsg "Bad file mode: " & tmpMode
    Exit Sub
End If

a = a + Len(tmpMode) + 4

If Trim(LCase(tmpMode)) = "random" Then

Else
    tmpHandle = Right(cmdStr, Len(cmdStr) - a)
End If

If Left(Trim(tmpHandle), 1) <> "#" Then
    ErrorMsg "Bad file handle: " & tmpHandle
    Exit Sub
End If


End Sub
Public Sub Cmd_Read(ByVal cmdStr As String)

Dim varList As New ArrayClass
Dim a As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, varList

For a = 1 To varList.itemCount
    SetValue (varList.Item(a))
    If compileError Then Exit Sub
Next a

End Sub
Public Sub Cmd_Restore(ByVal cmdStr As String)

Dim tmpType As Integer

cmdStr = Trim(cmdStr)

tmpType = EvalExpression(cmdStr)
If compileError Then Exit Sub

If tmpType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

End Sub
Public Sub Cmd_Run(ByVal cmdStr As String)

Dim paramStr As String
Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramStr = GetString(1, cmdStr, ",")
paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

If Len(paramStr) = Len(cmdStr) Then Exit Sub

paramStr = Trim(LCase(Mid(cmdStr, Len(paramStr) + 2)))
If Not ((paramStr = "normal") Or (paramStr = "hide") Or _
    (paramStr = "minimized") Or (paramStr = "maximized")) Then
        ErrorMsg "Illegal startup type"
End If


End Sub

Public Sub Cmd_Show(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_LineSize(ByVal cmdStr As String)

Dim paramType As Integer
Dim params As New ArrayClass

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub

Public Sub Cmd_Sprite(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 5 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

For n = 3 To 4
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

paramType = EvalExpression(params.Item(5))
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_SpriteBG(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub

Public Sub Cmd_SpriteHide(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_SpritePlay(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_SpritePause(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub

Public Sub Cmd_SpritePos(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 4 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

For n = 3 To 4
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_SpriteRate(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

paramType = EvalExpression(params.Item(3))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

End Sub


Public Sub Cmd_SpriteRotate(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_SpriteShow(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_SpriteSize(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 4 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

For n = 3 To 4
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_Stick(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_StopTimer(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub
Public Sub Cmd_Swap(ByVal cmdStr As String)

Dim var1, var2 As String
Dim type1, type2 As Integer
Dim a As Integer

    cmdStr = Trim(cmdStr)

'Seperate the two variables/arrays
var1 = GetString(1, cmdStr, ",")
If (Len(var1) = Len(cmdStr)) Or (Len(var1) + 1 = Len(cmdStr)) Then
    ErrorMsg "Missing second parameter of SWAP command"
    Exit Sub
End If
a = Len(var1) + 1
var2 = Right(cmdStr, Len(cmdStr) - a)

'Trim the spaces from them
var1 = Trim(var1)
var2 = Trim(var2)

'Get the type of each
type1 = EvalExpression(var1): If compileError Then Exit Sub
type2 = EvalExpression(var2): If compileError Then Exit Sub

'Type check the two variables/arrays
If type2 = DT_STRING Then
    SetValue var1, (Chr(34) & Chr(34))
    If compileError Then Exit Sub
ElseIf type2 = DT_NUMBER Then
    SetValue var1, "0"
    If compileError Then Exit Sub
End If
If type1 = DT_STRING Then
    SetValue var2, (Chr(34) & Chr(34))
    If compileError Then Exit Sub
ElseIf type1 = DT_NUMBER Then
    SetValue var2, "0"
    If compileError Then Exit Sub
End If


End Sub
Public Sub Cmd_TextColor(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub

If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
End If


End Sub

Public Sub Cmd_BGColor(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub

If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
End If


End Sub

Public Sub Cmd_Timer(ByVal cmdStr As String)

Dim subStr As String
Dim paramStr As String
Dim paramType, b As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramStr = GetString(1, cmdStr, ",")
If Len(paramStr) = Len(cmdStr) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If
paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

b = Len(paramStr) + 2

paramStr = GetString(b, cmdStr, ",")
If (b - 1) + Len(paramStr) = Len(cmdStr) Then
    ErrorMsg "Syntax error"
    Exit Sub
End If
paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

b = b + Len(paramStr) + 1

subStr = Trim(Mid(cmdStr, b))

For n = 1 To subName.itemCount
    If subStr = subName.Item(n) Then
        Exit Sub
    End If
Next n

For n = 1 To funcName.itemCount
    If (subStr & "()" = funcName.Item(n) & ")") Or (subStr = Left(funcName.Item(n), Len(funcName.Item(n)) - 1)) Then
        Exit Sub
    End If
Next n

ErrorMsg "Undefined function or sub: '" & subStr & "'"


End Sub

Public Sub Cmd_UnloadImg(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_Window(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim tmpParam As String
Dim paramType As Integer
Dim b As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

b = 1
While b <= Len(cmdStr)
    tmpParam = GetString(b, cmdStr, ",")
    b = Len(tmpParam) + b + 1
    params.Add Trim(tmpParam)
Wend

If params.itemCount <> 7 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For b = 1 To 2
    paramType = EvalExpression(params.Item(b))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next b

tmpParam = LCase(params.Item(3))
If Not ((tmpParam = "normal") Or (tmpParam = "dialog") Or _
    (tmpParam = "popup") Or (tmpParam = "dialog_modal")) Then
        ErrorMsg "Illegal window type"
        Exit Sub
End If

For b = 4 To 7
    paramType = EvalExpression(params.Item(b))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next b


End Sub
Public Sub Cmd_Disable(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub

Public Sub DefineSysVars()

Cmd_Var "ErrorMsg as string"
Cmd_Var "Red as number"
Cmd_Var "Yellow as number"
Cmd_Var "Orange as number"
Cmd_Var "Blue as number"
Cmd_Var "Green as number"
Cmd_Var "Purple as number"
Cmd_Var "Black as number"
Cmd_Var "White as number"
Cmd_Var "Brown as number"
Cmd_Var "ButtonFace as number"
Cmd_Var "DefPath as string"
Cmd_Var "CommandLine as string"
Cmd_Var "True as number"
Cmd_Var "False as number"
Cmd_Var "ScreenWidth as number"
Cmd_Var "ScreenHeight as number"
Cmd_Var "Inkey as string"
Cmd_Var "EnterKey as string"
Cmd_Var "RightKey as string"
Cmd_Var "LeftKey as string"
Cmd_Var "DownKey as string"
Cmd_Var "UpKey as string"
Cmd_Var "SpaceKey as string"

End Sub

Public Sub Cmd_DelSprite(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Cmd_DrawSprites(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_AddFrame(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

For n = 1 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch"
        Exit Sub
    End If
Next n

End Sub


Public Sub Func_GetItem(ByVal paramStr As String)

Dim params As New ArrayClass
Dim winName As String
Dim paramType, itemIdx As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Wrong number of arguments in GETITEM()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch in parameter 1 of GETITEM()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 2 of GETITEM()"

End Sub


Public Sub Func_GetLineText(ByVal paramStr As String)

Dim paramType As Integer
Dim params As New ArrayClass

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Wrong number of arguments in GETLINETEXT()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch in parameter 1 of GETLINETEXT()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 2 of GETLINETEXT()"

End Sub

Public Sub Func_GetSelIdx(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of GETSELIDX()"

End Sub


Public Sub Func_ItemCount(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of ITEMCOUNT()"

End Sub

Public Sub Func_LOF(ByVal paramStr As String)

paramStr = Trim(paramStr)

If Left(paramStr, 1) <> "#" Then
    ErrorMsg "File handle does not exist: " & paramStr
End If

End Sub


Public Sub Func_LineCount(ByVal paramStr As String)

Dim paramType As String

paramType = EvalExpression(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of LINECOUNT()"

End Sub



Public Sub Func_Replace(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Wrong number of arguments in REPLACE()"
    Exit Sub
End If

For n = 1 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch in parameter " & n & " of REPLACE()"
        Exit Sub
    End If
Next n

End Sub

Public Sub Func_RGB(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 3 Then
    ErrorMsg "Wrong number of arguments in RGB()"
    Exit Sub
End If

For n = 1 To 3
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch in parameter " & n & " of RGB()"
        Exit Sub
    End If
Next n

End Sub

Public Sub Func_String(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Wrong number of arguments in STRING()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch in parameter 1 of STRING()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 2 of STRING()"

End Sub

Public Sub Func_Word(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount < 2 Or params.itemCount > 3 Then
    ErrorMsg "Wrong number of arguments in WORD()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch in parameter 1 of WORD()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch in parameter 1 of WORD()"
    Exit Sub
End If

If params.itemCount = 2 Then Exit Sub

paramType = EvalExpression(params.Item(3))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch in parameter 1 of WORD()"
    Exit Sub
End If

End Sub

Public Function NameCheck(ByVal cmdStr As String) As Boolean

NameCheck = False

cmdStr = Trim(UCase(cmdStr))

If (Not (Asc(Left(cmdStr, 1)) >= 65 And Asc(Left(cmdStr, 1)) <= 90)) Then
  NameCheck = True
  Exit Function
End If

For n = 2 To Len(cmdStr)
  If (Asc(Mid(cmdStr, n, 1)) < 65 Or Asc(Mid(cmdStr, n, 1)) > 90) And _
  (Not (IsNumeric(Mid(cmdStr, n, 1)))) And Mid(cmdStr, n, 1) <> "_" Then
    NameCheck = True
    Exit Function
  End If
Next n

cmdStr = LCase(cmdStr)

If (cmdStr = "print") Or (cmdStr = "input") Or (cmdStr = "let") Or (cmdStr = "var") _
    Or (cmdStr = "if") Or (cmdStr = "else") Or (cmdStr = "elseif") Or (cmdStr = "end") _
    Or (cmdStr = "while") Or (cmdStr = "wend") Or (cmdStr = "for") Or (cmdStr = "next") _
    Or (cmdStr = "goto") Or (cmdStr = "gosub") Or (cmdStr = "return") Or (cmdStr = "consoltitle") _
    Or (cmdStr = "array") Or (cmdStr = "redim") Or (cmdStr = "bindvar") Or (cmdStr = "unbindvar") _
    Or (cmdStr = "sub") Or (cmdStr = "call") Or (cmdStr = "function") Or (cmdStr = "dim") _
    Or (cmdStr = "open") Or (cmdStr = "close") Or (cmdStr = "onerror") Or (cmdStr = "rem") _
    Or (cmdStr = "swap") Or (cmdStr = "on") Or (cmdStr = "data") Or (cmdStr = "read") _
    Or (cmdStr = "restore") Or (cmdStr = "select") Or (cmdStr = "case") Or (cmdStr = "textcolor") _
    Or (cmdStr = "bgcolor") Or (cmdStr = "pause") Or (cmdStr = "timer") Or (cmdStr = "stoptimer") _
    Or (cmdStr = "window") Or (cmdStr = "closewindow") Or (cmdStr = "event") Or (cmdStr = "message") _
    Or (cmdStr = "error") Or (cmdStr = "question") Or (cmdStr = "run") Or (cmdStr = "control") Then
      NameCheck = True
End If


End Function
Public Function CompileBlock(ByVal startLne As Integer, ByVal endLne As Integer) As Boolean

'Checks to see if the command at startLne is the beginning of a block. If so, it
'sends the block's code to the appropriate function and returns TRUE. Otherwise,
'it returns FALSE.

Dim ifNum, whileNum, forNum, selectNum As Integer
Dim n As Integer

CompileBlock = True

If LCase(Left(cmdList.Item(startLne), 3)) = "if " Then
    ifNum = 0
    'If the IF statement is multi-lined
    If LCase(Right(cmdList.Item(startLne), 5)) = " then" Then
        For n = startLne To endLne
          If (LCase(Trim(cmdList.Item(n))) = "end if" And ifNum = 0) Then
              CompileIfBlock startLne, n - 1, False
              n = endLne
              If compileError Then Exit Function
          ElseIf n = endLne Then
              ErrorMsg "IF block without END IF"
              Exit Function
          ElseIf (LCase(Left(LTrim(cmdList.Item(n)), 3)) = "if ") And (n > startLne) Then
              'Only adds it to the IF count if it is multi-lined (has an END IF)
              If LCase(Right(cmdList.Item(n), 5)) = " then" Then
                ifNum = ifNum + 1
              End If
          ElseIf LCase(Trim(cmdList.Item(n))) = "end if" Then
              ifNum = ifNum - 1
          End If
        Next n
    'If the IF statement is one-lined
    Else
        CompileIfBlock startLne, 0, True
        If compileError Then Exit Function
    End If
ElseIf LCase(Left(cmdList.Item(startLne), 6)) = "while " Then
    whileNum = 0
    For n = startLne To endLne
        If (LCase(Trim(cmdList.Item(n))) = "wend" And whileNum = 0) Then
            CompileWhileBlock startLne, n - 1
            n = endLne
            If compileError Then Exit Function
        ElseIf n = endLne Then
            ErrorMsg "WHILE block without WEND"
            Exit Function
        ElseIf (LCase(Left(LTrim(cmdList.Item(n)), 6)) = "while ") And (n > startLne) Then
            whileNum = whileNum + 1
        ElseIf LCase(Trim(cmdList.Item(n))) = "wend" Then
            whileNum = whileNum - 1
        End If
    Next n
ElseIf LCase(Left(cmdList.Item(startLne), 4)) = "for " Then
    forNum = 0
    For n = startLne To endLne
        If (LCase(Left(cmdList.Item(n), 5)) = "next " And forNum = 0) Then
            CompileForBlock startLne, n - 1
            n = endLne
            If compileError Then Exit Function
        ElseIf n = endLne Then
            ErrorMsg "FOR block without NEXT"
            Exit Function
        ElseIf (LCase(Left(LTrim(cmdList.Item(n)), 4)) = "for ") And (n > startLne) Then
            forNum = forNum + 1
        ElseIf LCase(Left(cmdList.Item(n), 5)) = "next " Then
            forNum = forNum - 1
        End If
    Next n
ElseIf LCase(Left(cmdList.Item(startLne), 12)) = "select case " Then
    selectNum = 0
    For n = startLne To endLne
        If (LCase(Trim(cmdList.Item(n))) = "end select" And selectNum = 0) Then
            CompileSelectBlock startLne, n - 1
            n = endLne
            If compileError Then Exit Function
        ElseIf n = endLne Then
            ErrorMsg "SELECT CASE block without END SELECT"
            Exit Function
        ElseIf (LCase(Left(LTrim(cmdList.Item(n)), 12)) = "select case ") And (n > startLne) Then
            selectNum = selectNum + 1
        ElseIf LCase(Trim(cmdList.Item(n))) = "end select" Then
            selectNum = selectNum - 1
        End If
    Next n
Else
    CompileBlock = False
End If

End Function
Public Sub CompileFuncBlock(ByVal startLne As Integer, ByVal endLne As Integer)

'Set inSubProg flag, add parameters and function name to
'the temporary sub-program variable collections, scan for
'local branch labels, and compile the function block

Dim tmpName, tmpParam, paramStr, typeStr, cmdStr As String
Dim b As Integer

With compileDlg
    .Status.Caption = "Compiling line " & txtLneNum.Item(lineNum) & " of " & txtLneNum.itemCount
    .progress.Value = lineNum
End With

    inSubProg = True
    subProgType = SP_FUNC

'Cut out the "FUNCTION " portion of the command string
cmdStr = Trim(Right(cmdList.Item(startLne), Len(cmdList.Item(startLne)) - 9))

tmpName = GetString(1, cmdStr, "(")

paramStr = GetString(Len(tmpName) + 1, LCase(cmdStr), " as ")
paramStr = Mid(cmdStr, Len(tmpName) + 1, Len(paramStr))

typeStr = Right(cmdStr, Len(cmdStr) - (Len(tmpName) + Len(paramStr)))

paramStr = GetString(2, paramStr, ")")

Cmd_Var tmpName & typeStr
If compileError Then Exit Sub

b = 1
While b <= Len(paramStr)
    tmpParam = GetString(b, paramStr, ",")
    b = Len(tmpParam) + b + 1
    Cmd_Var tmpParam
    If compileError Then Exit Sub
Wend

'Check for branch label definitions
compileDlg.Status.Caption = "Scanning for local branch labels..."
For lineNum = startLne + 1 To endLne
    If (Left(cmdList.Item(lineNum), 1) = "@") Then
      For b = 2 To Len(cmdList.Item(lineNum))
          If (Asc(UCase(Mid(cmdList.Item(lineNum), b, 1))) < 65 Or Asc(UCase(Mid(cmdList.Item(lineNum), b, 1))) > 90) And _
          (Not (IsNumeric(Mid(cmdList.Item(lineNum), b, 1)))) Then
              ErrorMsg "Illegal branch label name"
              Exit Sub
          End If
      Next b
      For a = 1 To subProg_LabelName.itemCount
        If cmdList.Item(lineNum) = subProg_LabelName.Item(a) Then
            ErrorMsg "Local branch label already defined"
            Exit Sub
        End If
      Next a
      subProg_LabelName.Add cmdList.Item(lineNum)
    End If
Next lineNum

For lineNum = startLne + 1 To endLne
    If Not CompileBlock(lineNum, endLne) Then
        CompileCmd cmdList.Item(lineNum)
    End If
    If compileError Then Exit Sub
Next lineNum

'Clean up temporary local variables, arrays, and labels
subProg_LabelName.Clear
subProg_VarName.Clear
subProg_VarType.Clear
subProg_ArrayName.Clear
subProg_ArrayType.Clear

inSubProg = False
lineNum = endLne + 1


End Sub
Public Sub CompileSelectBlock(ByVal startLne As Integer, ByVal endLne As Integer)

Dim tmpExp As String
Dim elseNum As Integer


With compileDlg
    .Status.Caption = "Compiling line " & txtLneNum.Item(lineNum) & " of " & txtLneNum.itemCount
    .progress.Value = lineNum
End With

tmpExp = Trim(Mid(cmdList.Item(startLne), 13))
EvalExpression (tmpExp)
If compileError Then Exit Sub

For lineNum = startLne + 1 To endLne
      If LCase(Trim(cmdList.Item(lineNum))) = "case else" Then
        elseNum = elseNum + 1
        If elseNum > 1 Then
            ErrorMsg "Too many CASE ELSE statements in SELECT CASE block"
            Exit Sub
        End If
      ElseIf LCase(Left(cmdList.Item(lineNum), 5)) = "case " Then
        tmpExp = Mid(cmdList.Item(lineNum), 6)
        EvalExpression (tmpExp)
        If compileError Then Exit Sub
      Else
        If Not CompileBlock(lineNum, endLne) Then
            CompileCmd cmdList.Item(lineNum)
        End If
        If compileError Then Exit Sub
    End If
Next lineNum
    
lineNum = endLne + 1


End Sub

Public Sub CompileSubBlock(ByVal startLne As Integer, ByVal endLne As Integer)

'Set inSubProg flag, add parameters to the temporary sub-program
'variable collections, scan for local branch labels, and
'compile the sub block

Dim tmpName, tmpParam, cmdStr As String
Dim b As Integer

With compileDlg
    .Status.Caption = "Compiling line " & txtLneNum.Item(lineNum) & " of " & txtLneNum.itemCount
    .progress.Value = lineNum
End With

    inSubProg = True
    subProgType = SP_SUB

'Cut out the "SUB " portion of the command string
cmdStr = Trim(Right(cmdList.Item(startLne), Len(cmdList.Item(startLne)) - 4))

tmpName = GetString(1, cmdStr, " ")

b = Len(tmpName) + 2
While b <= Len(cmdStr)
    tmpParam = GetString(b, cmdStr, ",")
    b = Len(tmpParam) + b + 1
    Cmd_Var tmpParam
    If compileError Then Exit Sub
Wend

'Check for branch label definitions
compileDlg.Status.Caption = "Scanning for local branch labels..."
For lineNum = startLne + 1 To endLne
    If (Left(cmdList.Item(lineNum), 1) = "@") Then
      For b = 2 To Len(cmdList.Item(lineNum))
          If (Asc(UCase(Mid(cmdList.Item(lineNum), b, 1))) < 65 Or Asc(UCase(Mid(cmdList.Item(lineNum), b, 1))) > 90) And _
          (Not (IsNumeric(Mid(cmdList.Item(lineNum), b, 1)))) Then
              ErrorMsg "Illegal branch label name"
              Exit Sub
          End If
      Next b
      For a = 1 To subProg_LabelName.itemCount
        If cmdList.Item(lineNum) = subProg_LabelName.Item(a) Then
            ErrorMsg "Local branch label already defined"
            Exit Sub
        End If
      Next a
      subProg_LabelName.Add cmdList.Item(lineNum)
    End If
Next lineNum

For lineNum = startLne + 1 To endLne
    If Not CompileBlock(lineNum, endLne) Then
        CompileCmd cmdList.Item(lineNum)
    End If
    If compileError Then Exit Sub
Next lineNum

'Clean up temporary local variables, arrays, and labels
subProg_LabelName.Clear
subProg_VarName.Clear
subProg_VarType.Clear
subProg_ArrayName.Clear
subProg_ArrayType.Clear

inSubProg = False
lineNum = endLne + 1


End Sub



Public Sub ErrorMsg(ByVal msgStr As String)

compileDlg.Status.Caption = "Compile error on line " & txtLneNum.Item(lineNum) & ": " & msgStr
editorWin.StatusBar.SimpleText = "Compile error."

SelectLine editorWin.editor, txtLneNum.Item(lineNum)

compileError = True

End Sub
Public Sub Cmd_BindVar(ByVal cmdStr As String)

Dim str, varStr As String
Dim tmpType As Integer

  cmdStr = Trim(cmdStr)
  varStr = GetString(1, LCase(cmdStr), " to ")
  varStr = Mid(varStr, 1, Len(varStr))
  If Len(varStr) = Len(cmdStr) Then ErrorMsg "Syntax error": Exit Sub

If inSubProg Then
  If subProg_VarName.itemCount = 0 Then ErrorMsg "Undefined variable": Exit Sub
  For n = 1 To subProg_VarName.itemCount
    If subProg_VarName.Item(n) = Trim(varStr) Then
        tmpType = subProg_VarType.Item(n)
        Exit For
    ElseIf subProg_VarName.Item(n) <> Trim(varStr) And n = subProg_VarName.itemCount Then
        ErrorMsg "Undefined variable": Exit Sub
    End If
  Next n
  varStr = Trim(Right(cmdStr, Len(cmdStr) - (Len(varStr) + 4)))
  For n = 1 To subProg_VarName.itemCount
    If subProg_VarName.Item(n) = varStr Then
        If tmpType <> subProg_VarType.Item(n) Then
            ErrorMsg "Variable type mismatch": Exit Sub
        End If
        Exit Sub
    End If
  Next n
Else
  If tmpVarName.itemCount = 0 Then ErrorMsg "Undefined variable": Exit Sub
  For n = 1 To tmpVarName.itemCount
    If tmpVarName.Item(n) = Trim(varStr) Then
        tmpType = tmpVarType.Item(n)
        Exit For
    ElseIf tmpVarName.Item(n) <> Trim(varStr) And n = tmpVarName.itemCount Then
        ErrorMsg "Undefined variable": Exit Sub
    End If
  Next n
  varStr = Trim(Right(cmdStr, Len(cmdStr) - (Len(varStr) + 4)))
  For n = 1 To tmpVarName.itemCount
    If tmpVarName.Item(n) = varStr Then
        If tmpType <> tmpVarType.Item(n) Then
            ErrorMsg "Variable type mismatch": Exit Sub
        End If
        Exit Sub
    End If
  Next n
End If

  ErrorMsg "Undefined variable": Exit Sub

End Sub

Public Sub Cmd_Array(ByVal cmdStr As String)

Dim str, varStr, tmpName As String

  cmdStr = Trim(cmdStr)
  varStr = GetString(1, cmdStr, " ")
  If varStr = "" Then ErrorMsg "Syntax error": Exit Sub
  If (Not (Asc(UCase(Left(varStr, 1))) >= 65 And Asc(UCase(Left(varStr, 1))) <= 90)) Then
    ErrorMsg "Illegal array name": Exit Sub
  End If
  For n = 2 To Len(GetString(1, varStr, "("))
    If (Asc(UCase(Mid(varStr, n, 1))) < 65 Or Asc(UCase(Mid(varStr, n, 1))) > 90) And _
    (Not (IsNumeric(Mid(varStr, n, 1)))) And Mid(varStr, n, 1) <> "_" Then
      ErrorMsg "Illegal array name": Exit Sub
    End If
  Next n

'*** Read array subscripts ***
  idxStr = GetString(Len(GetString(1, varStr, "(")) + 2, varStr, ")")
  If Trim(idxStr) = "" Then ErrorMsg "Missing index value in array declaration": Exit Sub
  firstIdx = GetString(1, idxStr, ",")
  idxType = EvalExpression(firstIdx): If compileError Then Exit Sub
  If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
  If Len(firstIdx) < Len(idxStr) Then
      secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
      idxType = EvalExpression(secondIdx): If compileError Then Exit Sub
      If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
  End If
'*****************************
  
  tmpName = Left(varStr, Len(GetString(1, varStr, "(")) + 1)

  If inSubProg Then
    For a = 1 To subProg_ArrayName.itemCount
      If tmpName = subProg_ArrayName.Item(a) Then
          ErrorMsg "Array already exists"
          Exit Sub
      End If
    Next a
  Else
    For a = 1 To arrayName.itemCount
      If tmpName = arrayName.Item(a) Then
          ErrorMsg "Array already exists"
          Exit Sub
      End If
    Next a
  End If
  
  str = Trim(Right(cmdStr, Len(cmdStr) - Len(varStr)))
  If LCase(Left(str, 3)) <> "as " Then ErrorMsg "Syntax error": Exit Sub
  
  If LCase(Trim(Right(str, Len(str) - 3))) = "string" Then
    If inSubProg Then
      subProg_ArrayType.Add DT_STRING
    Else
      arrayType.Add DT_STRING
    End If
  ElseIf LCase(Trim(Right(str, Len(str) - 3))) = "number" Then
    If inSubProg Then
      subProg_ArrayType.Add DT_NUMBER
    Else
      arrayType.Add DT_NUMBER
    End If
  Else
    ErrorMsg "Illegal array type": Exit Sub
  End If
  
  If inSubProg Then
    subProg_ArrayName.Add tmpName
  Else
    arrayName.Add tmpName
  End If
  
  
End Sub

Public Sub Cmd_Close(ByVal cmdStr As String)

cmdStr = Trim(cmdStr)

    If Left(cmdStr, 1) <> "#" Then
        ErrorMsg "Bad file handle: " & cmdStr
    End If

End Sub

Public Sub Cmd_ConsolTitle(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  ret = EvalExpression(cmdStr)

End Sub

Public Sub Cmd_GoTo(ByVal cmdStr As String)

'This sub is used for both GOTO and GOSUB, since
'they both do the same thing in the compiler

  cmdStr = Trim(cmdStr)
  
  If inSubProg Then
    For n = 1 To subProg_LabelName.itemCount
      If cmdStr = subProg_LabelName.Item(n) Then Exit Sub
    Next n
  Else
    For n = 1 To labelName.itemCount
      If cmdStr = labelName.Item(n) Then Exit Sub
    Next n
  End If
  
  ErrorMsg "Undefined branch label: " & cmdStr: Exit Sub

End Sub

Public Sub Cmd_Input(ByVal cmdStr As String)

Dim varStr, expStr, tmpHandle As String
Dim b As Integer

  cmdStr = Trim(cmdStr)
  
If Left(cmdStr, 1) = "#" Then
    tmpHandle = GetString(1, cmdStr, ",")
    If (Len(tmpHandle) = Len(cmdStr)) Or (Len(tmpHandle) + 1 = Len(cmdStr)) Then
        ErrorMsg "Syntax error"
        Exit Sub
    End If
    b = Len(tmpHandle) + 2
    While b <= Len(cmdStr)
        varStr = GetString(b, cmdStr, ",")
        b = b + Len(varStr) + 1
        SetValue varStr
        If compileError Then Exit Sub
    Wend

Else
  expStr = GetString(1, cmdStr, ",")
'*** If there is a prompt in the command ***
  If Len(expStr) < Len(cmdStr) Then
    EvalExpression (expStr)
    If compileError Then Exit Sub
    expStr = Mid(cmdStr, Len(expStr) + 2, Len(cmdStr) - (Len(expStr) + 1))
  End If
'*******************************************
  SetValue expStr: If compileError Then Exit Sub

End If


End Sub

Public Sub Cmd_Let(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  varStr = GetString(1, cmdStr, "=")
  If varStr = tmpCmdStr Then ErrorMsg "Syntax error": Exit Sub
  expStr = Mid(cmdStr, Len(varStr) + 2, Len(cmdStr) - (Len(varStr) + 1))
  varStr = Trim(varStr)
  SetValue varStr, expStr

End Sub

Public Sub Cmd_Print(ByVal cmdStr As String)

Dim tmpHandle, tmpExp As String
Dim b As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = ";" Then
    cmdStr = Left(cmdStr, Len(cmdStr) - 1)
End If

If Left(cmdStr, 1) = "#" Then
    tmpHandle = GetString(1, cmdStr, ",")
    If (Len(tmpHandle) = Len(cmdStr)) Or (Len(tmpHandle) + 1 = Len(cmdStr)) Then
        ErrorMsg "Syntax error"
        Exit Sub
    End If
    tmpExp = Right(cmdStr, Len(cmdStr) - (Len(tmpHandle) + 1))
    EvalExpression (tmpExp)
Else
    EvalExpression (cmdStr)
End If


End Sub



Public Sub Cmd_ReDim(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  
  For n = 1 To subProg_ArrayName.itemCount
    If Left(cmdStr, Len(subProg_ArrayName.Item(n))) = subProg_ArrayName.Item(n) And Right(cmdStr, 1) = ")" Then
    '*** Read array subscripts ***
        idxStr = GetString(Len(GetString(1, subProg_ArrayName.Item(n), "(")) + 2, cmdStr, ")")
        firstIdx = GetString(1, idxStr, ",")
        idxType = EvalExpression(firstIdx): If compileError Then Exit Sub
        If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
        If Len(firstIdx) < Len(idxStr) Then
            secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
            idxType = EvalExpression(secondIdx): If compileError Then Exit Sub
            If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
        End If
    '*****************************
        Exit Sub
    End If
  Next n

  For n = 1 To arrayName.itemCount
    If Left(cmdStr, Len(arrayName.Item(n))) = arrayName.Item(n) And Right(cmdStr, 1) = ")" Then
    '*** Read array subscripts ***
        idxStr = GetString(Len(GetString(1, arrayName.Item(n), "(")) + 2, cmdStr, ")")
        firstIdx = GetString(1, idxStr, ",")
        idxType = EvalExpression(firstIdx): If compileError Then Exit Sub
        If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
        If Len(firstIdx) < Len(idxStr) Then
            secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
            idxType = EvalExpression(secondIdx): If compileError Then Exit Sub
            If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
        End If
    '*****************************
        Exit Sub
    End If
  Next n
  
  ErrorMsg "Array doesn't exist": Exit Sub

End Sub



Public Sub Cmd_Var(ByVal cmdStr As String)

Dim tmpVar, varStr, typeStr, typeVal As String

  cmdStr = Trim(cmdStr)

'Get variable name(s)
varStr = GetString(1, LCase(cmdStr), " as ")
varStr = Left(cmdStr, Len(varStr))

'Syntax check
If Len(varStr) = Len(cmdStr) Then
  ErrorMsg "Syntax error": Exit Sub
End If

'Get variable type string
typeStr = LCase(Trim(Right(cmdStr, Len(cmdStr) - (Len(varStr) + 4))))

'Get type value
If typeStr = "number" Then
    typeVal = DT_NUMBER
ElseIf typeStr = "string" Then
    typeVal = DT_STRING
Else
    ErrorMsg "Illegal variable type: " & typeStr
    Exit Sub
End If

b = 1
While b <= Len(varStr)
    tmpVar = GetString(b, varStr, ",")
    b = Len(tmpVar) + b + 1
    tmpVar = Trim(tmpVar)
    If NameCheck(tmpVar) Then
        ErrorMsg "Illegal variable name"
        Exit Sub
    End If
    If inSubProg Then
        For n = 1 To subProg_VarName.itemCount
            If tmpVar = subProg_VarName.Item(n) Then
                ErrorMsg "Variable '" & tmpVar & "' already exists"
                Exit Sub
            End If
        Next n
        subProg_VarName.Add tmpVar
        subProg_VarType.Add typeVal
    Else
        For n = 1 To tmpVarName.itemCount
            If tmpVar = tmpVarName.Item(n) Then
                ErrorMsg "Variable '" & tmpVar & "' already exists"
                Exit Sub
            End If
        Next n
        tmpVarName.Add tmpVar
        tmpVarType.Add typeVal
    End If
Wend

End Sub


Public Sub Compile(ByVal code As String)

Dim tmpLine, tmpStr As String
Dim a, c, b, n As Integer

c = 0

'Seperate the code string into individual statements
For a = 1 To Len(code)
  If Mid(code, a, 2) = vbCrLf Or a = Len(code) Then
    c = c + 1
    If a = Len(code) Then tmpLine = tmpLine & Right(code, 1)
    tmpLine = Trim(tmpLine)
    If (tmpLine <> "") And (Left(tmpLine, 1) <> "'") And (LCase(Left(tmpLine, 4)) <> "rem ") Then
        b = 1
        While b <= Len(tmpLine)
            tmpStr = GetString(b, tmpLine, ":")
            b = Len(tmpStr) + b + 1
            cmdList.Add Trim(tmpStr)
            txtLneNum.Add c
        Wend
        tmpLine = "": a = a + 1
    Else
        tmpLine = "": a = a + 1
    End If
  Else
    tmpLine = tmpLine & Mid(code, a, 1)
  End If
Next a

'Cut out any comments that may be at the end of a line
For a = 1 To cmdList.itemCount
  cmdList.Item(a) = Trim(GetString(1, cmdList.Item(a), "'"))
    tmpStr = Trim(GetString(1, LCase(cmdList.Item(a) & " "), " rem "))
  cmdList.Item(a) = Left(cmdList.Item(a), Len(tmpStr))
Next a

'Append lines to any line ending with a "_"
For a = cmdList.itemCount To 1 Step -1
    If Right(cmdList.Item(a), 1) = "_" Then
        'Make sure there is a line of code to append
        If a < cmdList.itemCount Then
            cmdList.Item(a) = Left(cmdList.Item(a), Len(cmdList.Item(a)) - 1)
            cmdList.Item(a) = cmdList.Item(a) & cmdList.Item(a + 1)
            cmdList.Remove a + 1
            txtLneNum.Remove a + 1
        End If
    End If
Next a


'Set up the compiler dialog box
compileDlg.progress.Min = 0
compileDlg.progress.Max = cmdList.itemCount
compileDlg.Show

'Check statement count if unregistered
If Not isMbr Then
    If cmdList.itemCount > 150 Then
        lineNum = 150
        ErrorMsg "This unregistered copy only allows 150 statements to be compiled"
        Exit Sub
    End If
End If

'Check for branch label definitions
compileDlg.Status.Caption = "Scanning for branch labels..."
For lineNum = 1 To cmdList.itemCount
    If LCase(Left(cmdList.Item(lineNum), 4)) = "sub " Or _
    LCase(Left(cmdList.Item(lineNum), 9)) = "function " Then
        inSubProg = True
    ElseIf LCase(Trim(cmdList.Item(lineNum))) = "end sub" Or _
    LCase(Trim(cmdList.Item(lineNum))) = "end function" Then
        inSubProg = False
    End If
    If (Left(cmdList.Item(lineNum), 1) = "@") And (Not inSubProg) Then
      For b = 2 To Len(cmdList.Item(lineNum))
          If (Asc(UCase(Mid(cmdList.Item(lineNum), b, 1))) < 65 Or Asc(UCase(Mid(cmdList.Item(lineNum), b, 1))) > 90) And _
          (Not (IsNumeric(Mid(cmdList.Item(lineNum), b, 1)))) Then
              ErrorMsg "Illegal branch label name"
              Exit Sub
          End If
      Next b
      For a = 1 To labelName.itemCount
        If cmdList.Item(lineNum) = labelName.Item(a) Then
            ErrorMsg "Local branch label already defined"
            Exit Sub
        End If
      Next a
      labelName.Add cmdList.Item(lineNum)
    End If
Next lineNum
inSubProg = False

'Check for sub definitions
compileDlg.Status.Caption = "Scanning for subs..."
For lineNum = 1 To cmdList.itemCount
    If LCase(Left(cmdList.Item(lineNum), 4)) = "sub " Then
        AddSubDef Right(cmdList.Item(lineNum), Len(cmdList.Item(lineNum)) - 4)
        If compileError Then Exit Sub
    End If
Next lineNum

'Check for function definitions
compileDlg.Status.Caption = "Scanning for functions..."
For lineNum = 1 To cmdList.itemCount
    If LCase(Left(cmdList.Item(lineNum), 9)) = "function " Then
        AddFuncDef Right(cmdList.Item(lineNum), Len(cmdList.Item(lineNum)) - 9)
        If compileError Then Exit Sub
    End If
Next lineNum

'Define system variables
DefineSysVars

'Loop through and compile each statement
For lineNum = 1 To cmdList.itemCount
    If LCase(Left(cmdList.Item(lineNum), 4)) = "sub " Then
        For n = lineNum To cmdList.itemCount
            If (LCase(Trim(cmdList.Item(n))) = "end sub") Then
                CompileSubBlock lineNum, n - 1
                n = cmdList.itemCount
                If compileError Then Exit Sub
            ElseIf n = cmdList.itemCount Then
                ErrorMsg "SUB block without END SUB"
                Exit Sub
            ElseIf (LCase(Left(LTrim(cmdList.Item(n)), 4)) = "sub ") And (n > lineNum) Then
                ErrorMsg "Cannot create nested subs"
                Exit Sub
            ElseIf LCase(Left(LTrim(cmdList.Item(n)), 9)) = "function " Then
                ErrorMsg "Cannot define a function within a sub"
                Exit Sub
            End If
        Next n
    ElseIf LCase(Left(cmdList.Item(lineNum), 9)) = "function " Then
        For n = lineNum To cmdList.itemCount
            If (LCase(Trim(cmdList.Item(n))) = "end function") Then
                CompileFuncBlock lineNum, n - 1
                n = cmdList.itemCount
                If compileError Then Exit Sub
            ElseIf n = cmdList.itemCount Then
                ErrorMsg "FUNCTION block without END FUNCTION"
                Exit Sub
            ElseIf (LCase(Left(LTrim(cmdList.Item(n)), 9)) = "function ") And (n > lineNum) Then
                ErrorMsg "Cannot create nested functions"
                Exit Sub
            ElseIf LCase(Left(LTrim(cmdList.Item(n)), 4)) = "sub " Then
                ErrorMsg "Cannot define a sub within a function"
                Exit Sub
            End If
        Next n
    Else
        If Not CompileBlock(lineNum, cmdList.itemCount) Then
            CompileCmd cmdList.Item(lineNum)
        End If
        If compileError Then Exit Sub
    End If
Next lineNum

Unload compileDlg

End Sub


Public Sub CompileCmd(ByVal cmdStr As String)

Dim tmpCmdStr, str As String

With compileDlg
    .Status.Caption = "Compiling line " & txtLneNum.Item(lineNum) & " of " & txtLneNum.itemCount
    .progress.Value = lineNum
End With

cmdStr = Trim(cmdStr)

'When a command is found, fill tmpCmdStr with whatever comes
'after it and call the appropriate command evaluator

If LCase(Left(cmdStr, 6)) = "print " Then
  Cmd_Print Mid(cmdStr, 7)
ElseIf LCase(cmdStr) = "print" Then
  'NOOP
ElseIf LCase(Left(cmdStr, 6)) = "input " Then
  Cmd_Input Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 4)) = "let " Then
  Cmd_Let Mid(cmdStr, 5)
ElseIf LCase(Left(cmdStr, 4)) = "var " Then
  Cmd_Var Mid(cmdStr, 5)
ElseIf LCase(cmdStr) = "cls" Then
  'NOOP
ElseIf LCase(Left(cmdStr, 5)) = "goto " Then
  Cmd_GoTo Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 6)) = "gosub " Then
  Cmd_GoTo Mid(cmdStr, 7)
ElseIf LCase(cmdStr) = "return" Then
  'NOOP
ElseIf LCase(cmdStr) = "end" Then
  'NOOP
ElseIf LCase(Left(cmdStr, 12)) = "consoltitle " Then
  Cmd_ConsolTitle Mid(cmdStr, 13)
ElseIf LCase(Left(cmdStr, 6)) = "array " Then
  Cmd_Array Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 6)) = "redim " Then
  Cmd_ReDim Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 8)) = "bindvar " Then
  Cmd_BindVar Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 10)) = "unbindvar " Then
  Cmd_UnbindVar Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 5)) = "open " Then
  Cmd_Open Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 6)) = "close " Then
  Cmd_Close Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 5)) = "call " Then
  Cmd_Call Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 4)) = "dim " Then
  Cmd_Array Mid(cmdStr, 5)
ElseIf LCase(Left(cmdStr, 8)) = "onerror " Then
  Cmd_OnError Mid(cmdStr, 9)
ElseIf LCase(cmdStr) = "onerror" Then
  'NOOP
ElseIf LCase(Left(cmdStr, 5)) = "swap " Then
  Cmd_Swap Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 3)) = "on " Then
  Cmd_On Mid(cmdStr, 4)
ElseIf LCase(Left(cmdStr, 5)) = "read " Then
  Cmd_Read Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 5)) = "data " Then
  Cmd_Data Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 8)) = "restore " Then
  Cmd_Restore Mid(cmdStr, 9)
ElseIf LCase(cmdStr) = "restore" Then
  'NOOP
ElseIf LCase(Left(cmdStr, 10)) = "textcolor " Then
  Cmd_TextColor Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 8)) = "bgcolor " Then
  Cmd_BGColor Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 6)) = "timer " Then
  Cmd_Timer Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 10)) = "stoptimer " Then
  Cmd_StopTimer Mid(cmdStr, 11)
ElseIf LCase(Trim(cmdStr)) = "pause" Then
  'NOOP
ElseIf LCase(Left(cmdStr, 8)) = "message " Then
  Cmd_Message Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 9)) = "question " Then
  Cmd_Question Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 6)) = "error " Then
  Cmd_Message Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 4)) = "run " Then
  Cmd_Run Mid(cmdStr, 5)
ElseIf LCase(Left(cmdStr, 5)) = "seek " Then
  Cmd_Seek Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 7)) = "window " Then
  Cmd_Window Mid(cmdStr, 8)
ElseIf LCase(Left(cmdStr, 12)) = "closewindow " Then
  Cmd_CloseWindow Mid(cmdStr, 13)
ElseIf LCase(Left(cmdStr, 6)) = "event " Then
  Cmd_Event Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 8)) = "control " Then
  Cmd_Control Mid(cmdStr, 9)
ElseIf LCase(Trim(cmdStr)) = "showconsol" Then
  'NOOP
ElseIf LCase(Trim(cmdStr)) = "hideconsol" Then
  'NOOP
ElseIf LCase(Left(cmdStr, 7)) = "enable " Then
  Cmd_Enable Mid(cmdStr, 8)
ElseIf LCase(Left(cmdStr, 8)) = "disable " Then
  Cmd_Disable Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 5)) = "show " Then
  Cmd_Show Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 5)) = "hide " Then
  Cmd_Hide Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 8)) = "getsize " Then
  Cmd_GetSize Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 6)) = "getxy " Then
  Cmd_GetXY Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 8)) = "settext " Then
  Cmd_SetText Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 8)) = "setsize " Then
  Cmd_SetSize Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 6)) = "setxy " Then
  Cmd_SetXY Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 5)) = "menu " Then
  Cmd_Menu Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 8)) = "additem " Then
  Cmd_AddItem Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 8)) = "delitem " Then
  Cmd_DelItem Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 10)) = "setselidx " Then
  Cmd_SetSelIdx Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 8)) = "setitem " Then
  Cmd_SetItem Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 7)) = "getsel " Then
  Cmd_GetSel Mid(cmdStr, 8)
ElseIf LCase(Left(cmdStr, 7)) = "setsel " Then
  Cmd_SetSel Mid(cmdStr, 8)
ElseIf LCase(Left(cmdStr, 11)) = "setseltext " Then
  Cmd_SetSelText Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 17)) = "setclipboardtext " Then
  Cmd_SetClipboardText Mid(cmdStr, 18)
ElseIf LCase(Left(cmdStr, 5)) = "line " Then
  Cmd_Line Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 4)) = "box " Then
  Cmd_Box Mid(cmdStr, 5)
ElseIf LCase(Left(cmdStr, 7)) = "circle " Then
  Cmd_Circle Mid(cmdStr, 8)
ElseIf LCase(Left(cmdStr, 9)) = "linesize " Then
  Cmd_LineSize Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 8)) = "loadimg " Then
  Cmd_LoadImg Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 10)) = "unloadimg " Then
  Cmd_UnloadImg Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 8)) = "drawimg " Then
  Cmd_DrawImg Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 7)) = "getimg " Then
  Cmd_GetImg Mid(cmdStr, 8)
ElseIf LCase(Left(cmdStr, 6)) = "stick " Then
  Cmd_Stick Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 8)) = "refresh " Then
  Cmd_Refresh Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 6)) = "clear " Then
  Cmd_Clear Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 10)) = "backcolor " Then
  Cmd_BackColor Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 10)) = "forecolor " Then
  Cmd_ForeColor Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 9)) = "setpixel " Then
  Cmd_SetPixel Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 9)) = "drawtext " Then
  Cmd_DrawText Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 8)) = "setfont " Then
  Cmd_SetFont Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 7)) = "sprite " Then
  Cmd_Sprite Mid(cmdStr, 8)
ElseIf LCase(Left(cmdStr, 10)) = "delsprite " Then
  Cmd_DelSprite Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 12)) = "drawsprites " Then
  Cmd_DrawSprites Mid(cmdStr, 13)
ElseIf LCase(Left(cmdStr, 9)) = "addframe " Then
  Cmd_AddFrame Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 9)) = "delframe " Then
  Cmd_DelFrame Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 10)) = "spritepos " Then
  Cmd_SpritePos Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 11)) = "spritesize " Then
  Cmd_SpriteSize Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 11)) = "spriterate " Then
  Cmd_SpriteRate Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 13)) = "spriterotate " Then
  Cmd_SpriteRotate Mid(cmdStr, 14)
ElseIf LCase(Left(cmdStr, 11)) = "spriteshow " Then
  Cmd_SpriteShow Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 11)) = "spritehide " Then
  Cmd_SpriteHide Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 11)) = "spriteplay " Then
  Cmd_SpritePlay Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 12)) = "spritepause " Then
  Cmd_SpritePause Mid(cmdStr, 13)
ElseIf LCase(Left(cmdStr, 9)) = "spritebg " Then
  Cmd_SpriteBG Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 7)) = "runcmd " Then
  Cmd_RunCmd Mid(cmdStr, 8)
ElseIf LCase(Left(cmdStr, 11)) = "closesound " Then
  Cmd_CloseSound Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 10)) = "opensound " Then
  Cmd_OpenSound Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 11)) = "pausesound " Then
  Cmd_PauseSound Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 10)) = "playsound " Then
  Cmd_PlaySound Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 10)) = "stopsound " Then
  Cmd_StopSound Mid(cmdStr, 11)
ElseIf LCase(Trim(cmdStr)) = "beep" Then
  'NOOP
ElseIf LCase(Left(cmdStr, 9)) = "getfiles " Then
  Cmd_GetFiles Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 8)) = "getdirs " Then
  Cmd_GetDirs Mid(cmdStr, 9)
ElseIf LCase(Left(cmdStr, 5)) = "name " Then
  Cmd_Name Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 9)) = "setstate " Then
  Cmd_SetState Mid(cmdStr, 10)
ElseIf LCase(Left(cmdStr, 10)) = "buttonimg " Then
  Cmd_ButtonImg Mid(cmdStr, 11)
ElseIf LCase(Left(cmdStr, 6)) = "mkdir " Then
  Cmd_MkDir Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 6)) = "rmdir " Then
  Cmd_RmDir Mid(cmdStr, 7)
ElseIf LCase(Left(cmdStr, 11)) = "getmousexy " Then
  Cmd_GetMouseXY Mid(cmdStr, 12)
ElseIf LCase(Left(cmdStr, 4)) = "mid " Then
  Cmd_Mid Mid(cmdStr, 5)
ElseIf LCase(Left(cmdStr, 5)) = "kill " Then
  Cmd_Kill Mid(cmdStr, 6)
ElseIf LCase(Left(cmdStr, 12)) = "spritebgpos " Then
  Cmd_SpriteBGPos Mid(cmdStr, 13)
ElseIf LCase(Left(cmdStr, 12)) = "setsoundpos " Then
  Cmd_SetSoundPos Mid(cmdStr, 13)
ElseIf LCase(Trim(cmdStr)) = "exit sub" Then
  Cmd_ExitSub
ElseIf LCase(Trim(cmdStr)) = "exit function" Then
  Cmd_ExitFunction

ElseIf LCase(Left(cmdStr, 4)) = "sub " Then
  ErrorMsg "SUB cannot be defined within another block"
ElseIf LCase(Left(cmdStr, 9)) = "function " Then
  ErrorMsg "FUNCTION cannot be defined within another block"

'**** If any of these statements got through to this sub, then  ****
'**** it does not have a corrisponding starting statement with  ****
'**** it (i.e. NEXT without FOR).                               ****
ElseIf LCase(Left(cmdStr, 5)) = "next " Then
  ErrorMsg "NEXT without FOR"
ElseIf LCase(cmdStr) = "wend" Then
  ErrorMsg "WEND without WHILE"
ElseIf LCase(cmdStr) = "end if" Then
  ErrorMsg "END IF without IF"
ElseIf LCase(Left(cmdStr & " ", 5)) = "else " Then
  ErrorMsg "ELSE without IF"
ElseIf LCase(Left(cmdStr, 7)) = "elseif " Then
  ErrorMsg "ELSEIF without IF"
ElseIf LCase(cmdStr) = "end sub" Then
  ErrorMsg "END SUB without SUB"
ElseIf LCase(cmdStr) = "end function" Then
  ErrorMsg "END FUNCTION without FUNCTION"
ElseIf LCase(cmdStr) = "end select" Then
  ErrorMsg "END SELECT without SELECT CASE"
ElseIf LCase(cmdStr) = "case else" Then
  ErrorMsg "CASE ELSE without SELECT CASE"
ElseIf LCase(Left(cmdStr, 5)) = "case " Then
  ErrorMsg "CASE without SELECT CASE"
'******************************************************************
'******************************************************************

'This does nothing, but is here to keep the compiler from
'giving an error when hitting a branch label line
ElseIf Left(cmdStr, 1) = "@" Then
  'NOOP

'This checks to see if a varriable or array is being assigned
'a value, in which case the line is passed to the Let command
Else
  'If there isn't an equal sign in the command
  If Len(GetString(1, cmdStr, "=")) = Len(cmdStr) Then
    ErrorMsg "Syntax error"
  'If there is an equal sign in the command
  Else
    Cmd_Let cmdStr
  End If
End If

End Sub


Public Sub Cmd_SetSoundPos(ByVal cmdStr As String)

Dim params As New ArrayClass
Dim paramType As Integer

cmdStr = Trim(cmdStr)

If Right(cmdStr, 1) = "," Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

ParseParams cmdStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Syntax error"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch"

End Sub


Public Sub Cmd_RunCmd(ByVal cmdStr As String)

Dim paramType As Integer

cmdStr = Trim(cmdStr)

paramType = EvalExpression(cmdStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch"

End Sub


Public Sub CompileForBlock(ByVal startLne As Integer, ByVal endLne As Integer)

Dim str, expStr, varStr, tmpLine As String
Dim expType As Integer


With compileDlg
    .Status.Caption = "Compiling line " & txtLneNum.Item(lineNum) & " of " & txtLneNum.itemCount
    .progress.Value = lineNum
End With

If Len(Trim(cmdList.Item(startLne))) = 3 Then ErrorMsg "Syntax error": Exit Sub

varStr = GetString(5, cmdList.Item(startLne), "=")
If Len(varStr) + 3 = Len(cmdList.Item(startLne)) Then ErrorMsg "Syntax error": Exit Sub
If Trim(varStr) = "" Then ErrorMsg "FOR statement without variable": Exit Sub

SetValue varStr, "0"
If compileError Then Exit Sub

a = Len(varStr) + 6

expStr = GetString(a, LCase(cmdList.Item(startLne) & " "), " to ")
expStr = Mid(cmdList.Item(startLne), a, Len(expStr))
If a + Len(expStr) = Len(cmdList.Item(startLne)) Then ErrorMsg "FOR without TO": Exit Sub
If Len(Trim(expStr)) = "" Then ErrorMsg "Missing first expression in FOR statement": Exit Sub
expType = EvalExpression(expStr): If compileError Then Exit Sub
If expType <> DT_NUMBER Then ErrorMsg "First FOR expression must be numeric": Exit Sub

a = a + Len(expStr) + 3

If a > Len(cmdList.Item(startLne)) Then ErrorMsg "Missing second expression in FOR statement": Exit Sub
expStr = GetString(a, LCase(cmdList.Item(startLne) & " "), " step ")
expStr = Mid(cmdList.Item(startLne), a, Len(expStr))
expType = EvalExpression(expStr): If compileError Then Exit Sub
If a + Len(expStr) < Len(cmdList.Item(startLne)) Then
    a = a + Len(expStr) + 5
    If a > Len(cmdList.Item(startLne)) Then ErrorMsg "Missing STEP value": Exit Sub
    expStr = Right(cmdList.Item(startLne), Len(cmdList.Item(startLne)) - a)
    expType = EvalExpression(expStr): If compileError Then Exit Sub
    If expType <> DT_NUMBER Then ErrorMsg "STEP value must be numeric": Exit Sub
End If

For lineNum = startLne + 1 To endLne
    If Not CompileBlock(lineNum, endLne) Then
        CompileCmd cmdList.Item(lineNum)
    End If
    If compileError Then Exit Sub
Next lineNum
    
lineNum = endLne + 1

End Sub
Public Sub CompileIfBlock(ByVal startLne As Integer, ByVal endLne As Integer, ByVal oneLine As Boolean)

Dim str, expStr, tmpLine As String
Dim expType, elseNum

With compileDlg
    .Status.Caption = "Compiling line " & txtLneNum.Item(lineNum) & " of " & txtLneNum.itemCount
    .progress.Value = lineNum
End With

If Len(Trim(cmdList.Item(startLne))) = 2 Then ErrorMsg "Syntax error": Exit Sub

expStr = GetString(4, LCase(cmdList.Item(startLne)) & " ", " then ")
expStr = Mid(cmdList.Item(startLne), 4, Len(expStr))
If Len(expStr) + 3 = Len(cmdList.Item(startLne)) Then ErrorMsg "IF statement without THEN": Exit Sub
If Trim(expStr) = "" Then ErrorMsg "IF statement without expression": Exit Sub

expType = EvalExpression(expStr): If compileError Then Exit Sub
If expType <> DT_NUMBER Then ErrorMsg "Expression in IF statement must be numeric": Exit Sub

If oneLine Then
    str = LTrim(Right(cmdList.Item(startLne), Len(cmdList.Item(startLne)) - (Len(expStr) + 8)))
    CompileCmd str
    If compileError Then Exit Sub
Else
    elseNum = 0
    For lineNum = startLne + 1 To endLne
        If LCase(Left(cmdList.Item(lineNum), 7)) = "elseif " Then
            expStr = GetString(8, LCase(cmdList.Item(lineNum)) & " ", " then ")
            expStr = Mid(cmdList.Item(lineNum), 8, Len(expStr))
            If Len(expStr) + 7 = Len(cmdList.Item(lineNum)) Then ErrorMsg "ELSEIF statement without THEN": Exit Sub
            If Trim(expStr) = "" Then ErrorMsg "ELSEIF statement without expression": Exit Sub
            expType = EvalExpression(expStr): If compileError Then Exit Sub
            If expType <> DT_NUMBER Then ErrorMsg "Expression in ELSEIF statement must be numeric": Exit Sub
            If Len(expStr) + 12 < Len(cmdList.Item(lineNum)) Then
                 str = LTrim(Right(cmdList.Item(lineNum), Len(cmdList.Item(lineNum)) - (Len(expStr) + 12)))
                 CompileCmd str: If compileError Then Exit Sub
            End If
        ElseIf LCase(cmdList.Item(lineNum)) = "else" Then
            elseNum = elseNum + 1
            If elseNum > 1 Then
                 ErrorMsg "Too many ELSE statements in IF block"
                 Exit Sub
            End If
        Else
            If Not CompileBlock(lineNum, endLne) Then
                CompileCmd cmdList.Item(lineNum)
            End If
            If compileError Then Exit Sub
        End If
    Next lineNum
    lineNum = endLne + 1
End If

End Sub


Public Sub CompileWhileBlock(ByVal startLne As Integer, ByVal endLne As Integer)

Dim str, expStr, tmpLine As String
Dim expType As Integer

With compileDlg
    .Status.Caption = "Compiling line " & txtLneNum.Item(lineNum) & " of " & txtLneNum.itemCount
    .progress.Value = lineNum
End With

If Len(Trim(cmdList.Item(startLne))) = 5 Then ErrorMsg "Syntax error": Exit Sub

expStr = Right(cmdList.Item(startLne), Len(cmdList.Item(startLne)) - 6)
expType = EvalExpression(expStr): If compileError Then Exit Sub
If expType <> DT_NUMBER Then ErrorMsg "WHILE expression must be numeric": Exit Sub

For lineNum = startLne + 1 To endLne
    If Not CompileBlock(lineNum, endLne) Then
        CompileCmd cmdList.Item(lineNum)
    End If
    If compileError Then Exit Sub
Next lineNum
    
lineNum = endLne + 1

End Sub


Private Function EvalExpression(ByVal expStr As String) As Integer

'This function separates the operands from the operators in
'the given expression, determines the type of each operand,
'tests each operand set to determine if the operation for each
'operand set is correct, and returns the data type of the
'whole expression.
'
'If an operand is a string literal or a number value (begins
'with a quotation mark, negative sign, decimal, or number),
'determine its data type by making it a string if it begins
'with a quotation mark, or a number value if it begins with
'a negative sign, decimal, or a number.
'
'If an operand is a variable name, then make the data type of
'that operand the same type as the variable its referencing.
'
'If an operand is a function call, then the operand's data
'type will be determined by what function list it's in
'(string or number), and the EvalFunction() member function
'will be called to evaluate the parameters of the function.
'
'If a set of parentheses are encountered that do not belong
'to a function call, then the entire expression from within
'the parentheses will be retrieved using the GetString function,
'evaluated with the EvalExpression function, and a new operand
'will be added with the returned data type.
'
'If an operand is non of these, then it will flag a
'compile error.

Dim operandType As New ArrayClass
Dim operator As New ArrayClass
Dim tmpExp As String
Dim inString As Boolean

If Trim(expStr) = "" Then Exit Function

'parse out each operand
For I = 1 To Len(expStr)

  If Mid(expStr, I, 1) = Chr(34) Then
    If inString = False Then inString = True Else inString = False
  End If
  
  If Mid(expStr, I, 1) = "(" And inString = False Then
    temp = GetString(I + 1, expStr, ")")
    If I + Len(temp) = Len(expStr) Then
        ErrorMsg "Syntax error"
        Exit Function
    End If
    tmpExp = tmpExp & "(" & temp & ")"
    I = I + Len(temp) + 1
  
  ElseIf (Mid(expStr, I, 1) = "+" Or Mid(expStr, I, 1) = "%" _
  Or Mid(expStr, I, 1) = "*" Or Mid(expStr, I, 1) = "/" Or Mid(expStr, I, 1) = "^") _
  And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add OT_NUMERIC
    
  ElseIf (Mid(expStr, I, 2) = ">=" Or Mid(expStr, I, 2) = "<=" Or Mid(expStr, I, 2) = "<>") _
  And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add OT_COMPARISON: I = I + 1
    
  ElseIf (Mid(expStr, I, 1) = "=" Or Mid(expStr, I, 1) = "<" Or Mid(expStr, I, 1) = ">") _
  And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add OT_COMPARISON

  ElseIf (Mid(expStr, I, 1) = "&") And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add OT_STRING
  
  ElseIf (UCase(Mid(expStr, I, 5)) = " AND " Or UCase(Mid(expStr, I, 5)) = " XOR ") _
  And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add OT_NUMERIC: I = I + 4
  
  ElseIf (UCase(Mid(expStr, I, 4)) = " OR ") And inString = False Then
    GoSub AddType: tmpExp = ""
    operator.Add OT_NUMERIC: I = I + 3
  
  ElseIf Mid(expStr, I, 1) = "-" And inString = False Then
    If Len(Trim(tmpExp)) = 0 Then
      tmpExp = tmpExp & "-"
    Else
      GoSub AddType: tmpExp = ""
      operator.Add OT_NUMERIC
    End If
    
  Else
    tmpExp = tmpExp & Mid(expStr, I, 1)
  End If
  If compileError Then Exit Function
Next

    GoSub AddType

For a = operator.itemCount To 1 Step -1
  If operator.Item(a) = OT_NUMERIC Then
    If operandType.Item(a) = DT_NUMBER And operandType.Item(a + 1) = DT_NUMBER Then
        operandType.Remove (a)
        operator.Remove (a)
    Else
        ErrorMsg "Type mismatch": Exit Function
    End If
  End If
Next a

For b = operator.itemCount To 1 Step -1
  If operator.Item(b) = OT_STRING Then
    operandType.Remove (b + 1)
    operator.Remove (b)
    '*** FUCK FUCK FUCK FUCK ***
    operandType.Item(b) = DT_STRING
    '***************************
  End If
Next b

For b = operator.itemCount To 1 Step -1
  If operator.Item(b) = OT_COMPARISON Then
    operandType.Remove (b + 1)
    operator.Remove (b)
    '*** FUCK FUCK FUCK FUCK ***
    operandType.Item(b) = DT_NUMBER
    '***************************
  End If
Next b

EvalExpression = operandType.Item(1)

Exit Function


'routine to add an operand's type to the stack
AddType:
    tmpExp = Trim(tmpExp)
    If IsNumeric(tmpExp) Then
          operandType.Add DT_NUMBER
    ElseIf (Left(tmpExp, 1) = Chr(34)) And (Right(tmpExp, 1) = Chr(34)) Then
      operandType.Add DT_STRING
    ElseIf (Left(tmpExp, 1) = "(") And (Right(tmpExp, 1) = ")") Then
      tmpExp = Mid(tmpExp, 2, Len(tmpExp) - 2)
      operandType.Add EvalExpression(tmpExp)
    Else
      For a = 1 To subProg_VarName.itemCount
        If tmpExp = subProg_VarName.Item(a) Then
            operandType.Add subProg_VarType.Item(a)
            Return
        End If
      Next a
      For a = 1 To subProg_ArrayName.itemCount
        If Left(tmpExp, Len(subProg_ArrayName.Item(a))) = subProg_ArrayName.Item(a) And (Right(tmpExp, 1) = ")") Then
            operandType.Add subProg_ArrayType.Item(a)
        '*** Read array subscripts ***
            idxStr = GetString(Len(subProg_ArrayName.Item(a)) + 1, tmpExp, ")")
            firstIdx = GetString(1, idxStr, ",")
            idxType = EvalExpression(firstIdx): If compileError Then Exit Function
            If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Function
            If Len(firstIdx) < Len(idxStr) Then
                secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
                idxType = EvalExpression(secondIdx): If compileError Then Exit Function
                If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Function
            End If
        '*****************************
            Return
        End If
      Next a
      For a = 1 To tmpVarName.itemCount
        If tmpExp = tmpVarName.Item(a) Then
            operandType.Add tmpVarType.Item(a)
            Return
        End If
      Next a
      For a = 1 To arrayName.itemCount
        If Left(tmpExp, Len(arrayName.Item(a))) = arrayName.Item(a) And (Right(tmpExp, 1) = ")") Then
            operandType.Add arrayType.Item(a)
        '*** Read array subscripts ***
            idxStr = GetString(Len(arrayName.Item(a)) + 1, tmpExp, ")")
            firstIdx = GetString(1, idxStr, ",")
            idxType = EvalExpression(firstIdx): If compileError Then Exit Function
            If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Function
            If Len(firstIdx) < Len(idxStr) Then
                secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
                idxType = EvalExpression(secondIdx): If compileError Then Exit Function
                If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Function
            End If
        '*****************************
            Return
        End If
      Next a
      For b = 1 To strFunc.itemCount
        If LCase(Left(tmpExp, Len(strFunc.Item(b)))) = LCase(strFunc.Item(b)) And (Right(tmpExp, 1) = ")") Then
          If Right(tmpExp, 1) = ")" Then
            operandType.Add DT_STRING
            EvalFunction (tmpExp): If compileError Then Exit Function
            Return
          Else
            ErrorMsg "Syntax error": Exit Function
          End If
        End If
      Next b
      For c = 1 To numFunc.itemCount
        If LCase(Left(tmpExp, Len(numFunc.Item(c)))) = LCase(numFunc.Item(c)) Then
          If Right(tmpExp, 1) = ")" Then
            operandType.Add DT_NUMBER
            EvalFunction (tmpExp): If compileError Then Exit Function
            Return
          Else
            ErrorMsg "Syntax error": Exit Function
          End If
        End If
      Next c
      For c = 1 To funcName.itemCount
        If Left(tmpExp, Len(funcName.Item(c))) = funcName.Item(c) Then
          If Right(tmpExp, 1) = ")" Then
            operandType.Add funcType.Item(c)
            CallUserFunc (tmpExp): If compileError Then Exit Function
            Return
          Else
            ErrorMsg "Syntax error": Exit Function
          End If
        End If
      Next c
      ErrorMsg "Undefined variable or function: '" & tmpExp & "'": Exit Function
    End If
Return

End Function
Private Sub EvalFunction(ByVal funcStr As String)

'This sub evaluates the parameters of the given function
'call using the EvalExpression function for each parameter,
'making sure each one has the correct data type.

funcStr = Trim(funcStr)

If LCase(Left(funcStr, 4)) = "abs(" Then
    Func_Abs GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "asc(" Then
    Func_Asc GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "not(" Then
    Func_Not GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "int(" Then
    Func_Int GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "len(" Then
    Func_Len GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "rnd(" Then
    'NOOP
ElseIf LCase(Left(funcStr, 4)) = "val(" Then
    Func_Val GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "chr(" Then
    Func_Chr GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "str(" Then
    Func_Str GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 6)) = "upper(" Then
    Func_Upper GetString(7, funcStr, ")")
ElseIf LCase(Left(funcStr, 6)) = "lower(" Then
    Func_Lower GetString(7, funcStr, ")")
ElseIf LCase(Left(funcStr, 5)) = "trim(" Then
    Func_Trim GetString(6, funcStr, ")")
ElseIf LCase(Left(funcStr, 5)) = "left(" Then
    Func_Left GetString(6, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "mid(" Then
    Func_Mid GetString(5, funcStr, ")")
ElseIf LCase(Left(funcStr, 6)) = "right(" Then
    Func_Right GetString(7, funcStr, ")")
ElseIf LCase(Left(funcStr, 6)) = "instr(" Then
    Func_InStr GetString(7, funcStr, ")")
ElseIf LCase(Left(funcStr, 4)) = "loc(" Then
    Func_Loc (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 8)) = "gettext(" Then
    Func_GetText (GetString(9, funcStr, ")"))
ElseIf LCase(Left(funcStr, 5)) = "hwnd(" Then
    Func_hWnd (GetString(6, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "hdc(" Then
    Func_hDC (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 10)) = "getselidx(" Then
    Func_GetSelIdx (GetString(11, funcStr, ")"))
ElseIf LCase(Left(funcStr, 8)) = "getitem(" Then
    Func_GetItem (GetString(9, funcStr, ")"))
ElseIf LCase(Left(funcStr, 10)) = "itemcount(" Then
    Func_ItemCount (GetString(11, funcStr, ")"))
ElseIf LCase(Left(funcStr, 11)) = "getseltext(" Then
    Func_GetSelText (GetString(12, funcStr, ")"))
ElseIf LCase(Left(funcStr, 10)) = "linecount(" Then
    Func_LineCount (GetString(11, funcStr, ")"))
ElseIf LCase(Left(funcStr, 12)) = "getlinetext(" Then
    Func_GetLineText (GetString(13, funcStr, ")"))
ElseIf LCase(Left(funcStr, 17)) = "getclipboardtext(" Then
    'NOOP
ElseIf LCase(Left(funcStr, 9)) = "inputbox(" Then
    Func_InputBox (GetString(10, funcStr, ")"))
ElseIf LCase(Left(funcStr, 5)) = "date(" Then
    'NOOP
ElseIf LCase(Left(funcStr, 5)) = "time(" Then
    'NOOP
ElseIf LCase(Left(funcStr, 9)) = "getstate(" Then
    Func_GetState (GetString(10, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "min(" Then
    Func_Min (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "max(" Then
    Func_Max (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "sqr(" Then
    Func_Sqr (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 6)) = "space(" Then
    Func_Space (GetString(7, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "lof(" Then
    Func_LOF (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "eof(" Then
    Func_EOF (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 5)) = "hbmp(" Then
    Func_hBmp (GetString(6, funcStr, ")"))
ElseIf LCase(Left(funcStr, 9)) = "fileopen(" Then
    Func_FileOpen (GetString(10, funcStr, ")"))
ElseIf LCase(Left(funcStr, 9)) = "filesave(" Then
    Func_FileSave (GetString(10, funcStr, ")"))
ElseIf LCase(Left(funcStr, 8)) = "collide(" Then
    Func_Collide (GetString(9, funcStr, ")"))
ElseIf LCase(Left(funcStr, 6)) = "input(" Then
    Func_Input (GetString(7, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "rgb(" Then
    Func_RGB (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 12)) = "getsoundpos(" Then
    Func_GetSoundPos (GetString(13, funcStr, ")"))
ElseIf LCase(Left(funcStr, 12)) = "getsoundlen(" Then
    Func_GetSoundLen (GetString(13, funcStr, ")"))
ElseIf LCase(Left(funcStr, 8)) = "replace(" Then
    Func_Replace (GetString(9, funcStr, ")"))
ElseIf LCase(Left(funcStr, 7)) = "string(" Then
    Func_String (GetString(8, funcStr, ")"))
ElseIf LCase(Left(funcStr, 5)) = "word(" Then
    Func_Word (GetString(6, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "sin(" Then
    Func_Sin (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "cos(" Then
    Func_Cos (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "tan(" Then
    Func_Tan (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "log(" Then
    Func_Log (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "exp(" Then
    Func_Exp (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "atn(" Then
    Func_Atn (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "hex(" Then
    Func_Hex (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "oct(" Then
    Func_Oct (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 4)) = "sgn(" Then
    Func_Sgn (GetString(5, funcStr, ")"))
ElseIf LCase(Left(funcStr, 6)) = "round(" Then
    Func_Round (GetString(7, funcStr, ")"))
End If

End Sub


Public Sub Func_Sin(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of SIN()"

End Sub


Public Sub Func_Cos(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of COS()"

End Sub


Public Sub Func_Tan(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of TAN()"

End Sub


Public Sub Func_Log(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of LOG()"

End Sub


Public Sub Func_Exp(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of EXP()"

End Sub


Public Sub Func_Atn(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of ATN()"

End Sub


Public Sub Func_Hex(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of HEX()"

End Sub


Public Sub Func_Oct(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of OCT()"

End Sub


Public Sub Func_Round(ByVal paramStr As String)

Dim paramType As Integer
Dim params As New ArrayClass

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount < 1 Or params.itemCount > 2 Then
    ErrorMsg "Wrong number of arguments in ROUND()"
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch in parameter 1 of ROUND()"
    Exit Sub
End If

If params.itemCount = 1 Then Exit Sub

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 2 of ROUND()"

End Sub


Public Sub Func_Sgn(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of SGN()"

End Sub


Public Sub Func_Min(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Wrong number of arguments in MIN()"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch in parameter " & n & " of MIN()"
        Exit Sub
    End If
Next n

End Sub


Public Sub Func_Max(ByVal paramStr As String)

Dim params As New ArrayClass
Dim paramType, n As Integer

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount <> 2 Then
    ErrorMsg "Wrong number of arguments in MAX()"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch in parameter " & n & " of MAX()"
        Exit Sub
    End If
Next n

End Sub







Public Sub Func_Sqr(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of SQR()"

End Sub


Public Sub Func_Space(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter 1 of SPACE()"

End Sub



Public Sub Func_GetSelText(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of GETSELTEXT()"

End Sub



Public Sub Func_Abs(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter of 1 ABS()": Exit Sub

End Sub


Public Sub Func_Asc(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter of 1 ASC()": Exit Sub

End Sub



Public Sub Func_Chr(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter of 1 CHR()": Exit Sub

End Sub

Public Sub Func_InStr(ByVal paramStr As String)

Dim paramType, n As Integer
Dim params As New ArrayClass

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount < 2 Or params.itemCount > 3 Then
    ErrorMsg "Wrong number of arguments in INSTR()"
    Exit Sub
End If

For n = 1 To 2
    paramType = EvalExpression(params.Item(n))
    If compileError Then Exit Sub
    If paramType <> DT_STRING Then
        ErrorMsg "Type mismatch in parameter " & n & " of INSTR()"
        Exit Sub
    End If
Next n

If params.itemCount = 3 Then
    paramType = EvalExpression(params.Item(3))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch in parameter 3 of INSTR()"
        Exit Sub
    End If
End If

End Sub

Public Sub Func_Int(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter of 1 INT()": Exit Sub

End Sub

Public Sub Func_Left(ByVal paramStr As String)

paramStr = Trim(paramStr)

param = GetString(1, paramStr, ",")
paramType = EvalExpression(param): If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter of 1 LEFT()": Exit Sub
param = Right(paramStr, Len(paramStr) - (Len(param) + 1))
paramType = EvalExpression(param): If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter of 2 LEFT()": Exit Sub

End Sub

Public Sub Func_Len(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter of 1 LEN()": Exit Sub

End Sub


Public Sub Func_Loc(ByVal paramStr As String)

paramStr = Trim(paramStr)

If Left(paramStr, 1) <> "#" Then
    ErrorMsg "File handle does not exist: " & paramStr
End If

End Sub
Public Sub Func_Lower(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter of 1 LOWER()": Exit Sub

End Sub

Public Sub Func_Mid(ByVal paramStr As String)

Dim paramType As Integer
Dim params As New ArrayClass

paramStr = Trim(paramStr)

ParseParams paramStr, params

If params.itemCount < 2 Or params.itemCount > 3 Then
    ErrorMsg "Wrong number of arguments in MID()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(1))
If compileError Then Exit Sub
If paramType <> DT_STRING Then
    ErrorMsg "Type mismatch in parameter 1 of MID()"
    Exit Sub
End If

paramType = EvalExpression(params.Item(2))
If compileError Then Exit Sub
If paramType <> DT_NUMBER Then
    ErrorMsg "Type mismatch in parameter 2 of MID()"
    Exit Sub
End If

If params.itemCount = 3 Then
    paramType = EvalExpression(params.Item(3))
    If compileError Then Exit Sub
    If paramType <> DT_NUMBER Then
        ErrorMsg "Type mismatch in parameter 3 of MID()"
        Exit Sub
    End If
End If

End Sub
Public Sub Func_Not(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter of 1 NOT()": Exit Sub

End Sub


Public Sub Func_Right(ByVal paramStr As String)

paramStr = Trim(paramStr)

param = GetString(1, paramStr, ",")
paramType = EvalExpression(param): If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter of 1 LEFT()": Exit Sub
param = Right(paramStr, Len(paramStr) - (Len(param) + 1))
paramType = EvalExpression(param): If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter of 2 LEFT()": Exit Sub

End Sub

Public Sub Func_GetText(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of GETTEXT()"

End Sub







Public Sub Func_hDC(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of HDC()"

End Sub


Public Sub Func_hWnd(ByVal paramStr As String)

Dim paramType As Integer

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr)
If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter 1 of HWND()"

End Sub
Public Sub Func_Str(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_NUMBER Then ErrorMsg "Type mismatch in parameter of 1 STR()": Exit Sub

End Sub


Public Sub Func_Trim(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter of 1 TRIM()": Exit Sub

End Sub

Public Sub Func_Upper(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter of 1 UPPER()": Exit Sub

End Sub


Public Sub Func_Val(ByVal paramStr As String)

paramStr = Trim(paramStr)

paramType = EvalExpression(paramStr): If compileError Then Exit Sub
If paramType <> DT_STRING Then ErrorMsg "Type mismatch in parameter of 1 VAL()": Exit Sub

End Sub


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

Public Sub SetValue(ByVal varStr As String, Optional ByVal expStr As String)

Dim expType As Integer

varStr = Trim(varStr)

For n = 1 To subProg_VarName.itemCount
    If subProg_VarName.Item(n) = varStr Then
        If expStr <> Empty Then
            expType = EvalExpression(expStr): If compileError Then Exit Sub
            If expType <> subProg_VarType.Item(n) Then ErrorMsg "Type mismatch": Exit Sub
        End If
        Exit Sub
    End If
Next n
For n = 1 To subProg_ArrayName.itemCount
    If Left(varStr, Len(subProg_ArrayName.Item(n))) = subProg_ArrayName.Item(n) Then
      If Right(varStr, 1) = ")" Then
        If expStr <> Empty Then
            expType = EvalExpression(expStr): If compileError Then Exit Sub
            If expType <> subProg_ArrayType.Item(n) Then ErrorMsg "Type mismatch": Exit Sub
        End If
    '*** Read array subscripts ***
        idxStr = GetString(Len(subProg_ArrayName.Item(n)) + 1, varStr, ")")
        firstIdx = GetString(1, idxStr, ",")
        idxType = EvalExpression(firstIdx): If compileError Then Exit Sub
        If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
        If Len(firstIdx) < Len(idxStr) Then
            secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
            idxType = EvalExpression(secondIdx): If compileError Then Exit Sub
            If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
        End If
    '*****************************
        Exit Sub
      Else
        ErrorMsg "Syntax error": Exit Sub
      End If
    End If
Next n

For n = 1 To tmpVarName.itemCount
    If tmpVarName.Item(n) = varStr Then
        If expStr <> Empty Then
            expType = EvalExpression(expStr): If compileError Then Exit Sub
            If expType <> tmpVarType.Item(n) Then ErrorMsg "Type mismatch": Exit Sub
        End If
        Exit Sub
    End If
Next n
For n = 1 To arrayName.itemCount
    If Left(varStr, Len(arrayName.Item(n))) = arrayName.Item(n) Then
      If Right(varStr, 1) = ")" Then
        If expStr <> Empty Then
            expType = EvalExpression(expStr): If compileError Then Exit Sub
            If expType <> arrayType.Item(n) Then ErrorMsg "Type mismatch": Exit Sub
        End If
    '*** Read array subscripts ***
        idxStr = GetString(Len(arrayName.Item(n)) + 1, varStr, ")")
        firstIdx = GetString(1, idxStr, ",")
        idxType = EvalExpression(firstIdx): If compileError Then Exit Sub
        If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
        If Len(firstIdx) < Len(idxStr) Then
            secondIdx = Right(idxStr, Len(idxStr) - (Len(firstIdx) + 1))
            idxType = EvalExpression(secondIdx): If compileError Then Exit Sub
            If idxType <> DT_NUMBER Then ErrorMsg "Array index must be numeric": Exit Sub
        End If
    '*****************************
        Exit Sub
      Else
        ErrorMsg "Syntax error": Exit Sub
      End If
    End If
Next n

ErrorMsg "Undefined variable or array name: " & varStr: Exit Sub

End Sub
Public Sub Cmd_UnbindVar(ByVal cmdStr As String)

  cmdStr = Trim(cmdStr)
  varStr = GetString(1, LCase(cmdStr), " from ")
  varStr = Mid(varStr, 1, Len(varStr))
  If Len(varStr) = Len(cmdStr) Then ErrorMsg "Syntax error": Exit Sub
  
If inSubProg Then
  If subProg_VarName.itemCount = 0 Then ErrorMsg "Undefined variable": Exit Sub
  For n = 1 To subProg_VarName.itemCount
    If subProg_VarName.Item(n) = Trim(varStr) Then
        Exit For
    ElseIf subProg_VarName.Item(n) <> Trim(varStr) And n = subProg_VarName.itemCount Then
        ErrorMsg "Undefined variable": Exit Sub
    End If
  Next n
  varStr = Trim(Right(cmdStr, Len(cmdStr) - (Len(varStr) + 6)))
  For n = 1 To subProg_VarName.itemCount
    If subProg_VarName.Item(n) = varStr Then
        Exit Sub
    End If
  Next n
Else
  If tmpVarName.itemCount = 0 Then ErrorMsg "Undefined variable": Exit Sub
  For n = 1 To tmpVarName.itemCount
    If tmpVarName.Item(n) = Trim(varStr) Then
        Exit For
    ElseIf tmpVarName.Item(n) <> Trim(varStr) And n = tmpVarName.itemCount Then
        ErrorMsg "Undefined variable": Exit Sub
    End If
  Next n
  varStr = Trim(Right(cmdStr, Len(cmdStr) - (Len(varStr) + 6)))
  For n = 1 To tmpVarName.itemCount
    If tmpVarName.Item(n) = varStr Then
        Exit Sub
    End If
  Next n
End If
  
  ErrorMsg "Undefined variable": Exit Sub

End Sub

Public Function SubProgExists(ByVal spName As String) As Boolean

spName = Trim(spName)

SubProgExists = False

For n = 1 To subName.itemCount
    If spName = subName.Item(n) Then
        SubProgExists = True
        Exit Function
    End If
Next n
For n = 1 To funcName.itemCount
    If (spName & "()" = funcName.Item(n) & ")") Or (spName = Left(funcName.Item(n), Len(funcName.Item(n)) - 1)) Then
        SubProgExists = True
        Exit Function
    End If
Next n

End Function

Private Sub Class_Initialize()

compileError = False
inSubProg = False

End Sub


