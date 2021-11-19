Attribute VB_Name = "CompileCode"
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Data Types
Public Const DT_STRING = 1
Public Const DT_NUMBER = 2

'Operator Types
Public Const OT_NUMERIC = 1
Public Const OT_STRING = 2
Public Const OT_COMPARISON = 3

'Sub Program Types
Public Const SP_SUB = 1
Public Const SP_FUNC = 2

Public strFunc As New ArrayClass
Public numFunc As New ArrayClass

Public compileError As Boolean

Public ToolProg As New ArrayClass
Public ToolName As New ArrayClass

Public titleBar As String

Public isMbr As Boolean
Public rFile As String


Public Function CmpMbr(ByVal nParam As String, ByVal pParam As String) As Boolean

'Compare the user name as password to see if they are valid

Dim eName, tmpChar As String

For n = 1 To Len(nParam) Step 3
    tmpChar = Chr(Asc(Mid(nParam, n, 1)) + 20 - n)
    If Asc(tmpChar) > 32 And Asc(tmpChar) < 127 Then
        eName = eName & tmpChar
    End If
Next n

If eName = pParam Then CmpMbr = True Else CmpMbr = False

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
numFunc.Add "rgb("
numFunc.Add "getsoundpos("
numFunc.Add "getsoundlen("
numFunc.Add "sin("
numFunc.Add "cos("
numFunc.Add "tan("
numFunc.Add "log("
numFunc.Add "exp("
numFunc.Add "atn("
numFunc.Add "round("
numFunc.Add "sgn("

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
strFunc.Add "replace("
strFunc.Add "string("
strFunc.Add "word("
strFunc.Add "hex("
strFunc.Add "oct("

End Sub
Public Sub LoadMbr()

Dim eFile, fileData As String
Dim upArr() As String
Dim pathStr As String
Dim pathLen As Long

isMbr = False

'eFile = "0(KGZ`WcliSKrmoajZq2e4663y€{"
eFile = "d%X'))&lsn"
For n = 1 To Len(eFile)
    rFile = rFile & Chr(Asc(Mid(eFile, n, 1)) + 15 - n)
Next n

pathStr = Space(1024)
pathLen = GetSystemDirectory(pathStr, 1024)
pathStr = Left(pathStr, pathLen)
If Right(pathStr, 1) <> "\" Then pathStr = pathStr & "\"

rFile = pathStr & rFile

If Dir(rFile) = "" Then Exit Sub
Open rFile For Binary As #1
    fileData = Input(LOF(1), 1)
Close #1
If Trim(fileData) = "" Then Exit Sub

fileData = DTask(fileData)
upArr = Split(fileData, vbCrLf)

If CmpMbr(upArr(0), upArr(1)) Then isMbr = True

End Sub

Public Sub WriteRunFile(code As ArrayClass, ByVal fName As String)

'Takes an array of compiled code lines, encrypts them, and
'writes them into the default runtime file for execution
'by the runtime engine

'If Dir(App.Path & "\runtime.mbr") <> "" Then Kill App.Path & "\runtime.mbr"

Dim listStr, fileData As String
Dim nsName As String
Dim fileStr As String
Dim dat As String
Dim n As Long

'Add nag screen code if unregistered
If Not isMbr Then
    For n = 1 To 8
        Randomize
        nsName = nsName & Chr(Int((Rnd * 25) + 65))
    Next n
    code.Add "window " & Chr(34) & nsName & Chr(34) & ", " & Chr(34) & "Unregistered" & Chr(34) & ", dialog, 200, 200, 400, 200", 1
    code.Add "control " & Chr(34) & "nsLabel" & Chr(34) & ", " & Chr(34) & nsName & Chr(34) & ", " & Chr(34) & "This program was made with an unregistered copy of MicroByte." & Chr(34) & ", statictext, 90, 60, 250, 100", 2
    code.Add "control " & Chr(34) & "nsBtn" & Chr(34) & ", " & Chr(34) & nsName & Chr(34) & ", " & Chr(34) & "OK" & Chr(34) & ", button, 150, 120, 72, 25", 3
    code.Add "event " & Chr(34) & "nsBtn" & Chr(34) & ", " & Chr(34) & "click" & Chr(34) & ", NsBtnClick", 4
    code.Add "sub NsBtnClick", 5
    code.Add "closewindow " & Chr(34) & nsName & Chr(34), 6
    code.Add "end sub", 7
End If

'Load each command into the string
For n = 1 To code.itemCount
    listStr = listStr & code.Item(n) & vbCrLf
Next n

'Encrypt the file data
fileData = ETask(listStr)

'Read the EXE file
Open fName For Binary As #1
    fileStr = Space(LOF(1))
    Get #1, , fileStr
    'fileStr = Input(LOF(1), 1) 'Too sloooooooooowwwww
Close #1

'Get the position of the extra data (if present)
n = InStr(1, fileStr, "lib")
If n = 0 Then n = Len(fileStr) Else n = n - 1

'Extract the actual EXE data from the extra data
dat = Mid(fileStr, 1, n)

'Append the new extra data to the EXE data
fileStr = dat & "lib" & fileData

'Delete the old EXE file
Kill fName

'Write the EXE data and new extra data to a new EXE file
Open fName For Binary As #1
    Put #1, , fileStr
Close #1

End Sub


Public Function ETask(ByVal dat As String) As String

Dim eData, eKey, tmpDat As String

Randomize
eKey = Chr(Int((Rnd * 10) + 10))

'Add key and beginning padding
For n = 1 To 10
    Randomize
    eData = eData & Chr(Int((Rnd * 127) + 1))
Next n
eData = eData & eKey
For n = 1 To 8
    Randomize
    eData = eData & Chr(Int((Rnd * 127) + 1))
Next n

'Add encrypted data
For n = 1 To Len(dat)
    tmpDat = tmpDat & Chr(Asc(Mid(dat, n, 1)) + Asc(eKey))
Next n
For n = 1 To Len(tmpDat) Step 2
    Mid(tmpDat, n, 1) = Chr(Asc(Mid(tmpDat, n, 1)) - 3)
Next n
eData = eData & tmpDat

'Add end padding
For n = 1 To 10
    eData = eData & Chr(Int((Rnd * 127) + 1))
Next n

ETask = eData

End Function


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


Public Sub LoadOptions()

If Dir(App.Path & "\options.dat") = "" Then Exit Sub

Dim bcolor, fcolor As Long

Open App.Path & "\options.dat" For Input As #1
    Input #1, bcolor: editorWin.editor.BackColor = bcolor
    Input #1, fcolor: editorWin.editor.SelColor = fcolor
Close #1


End Sub
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
Public Sub SelectLine(txtControl As Object, ByVal theLine As Integer)

'This sub selects the given line from a VB textbox control

Dim lneCount As Integer
Dim lneTxt As String

lneCount = 0

For a = 1 To Len(txtControl.Text)
    If Mid(txtControl.Text, a, 2) = vbCrLf Then
        lneCount = lneCount + 1
        If lneCount = theLine Then
            txtControl.SelStart = (a - 1) - Len(lneTxt)
            txtControl.SelLength = Len(lneTxt)
            Exit Sub
        Else
            a = a + 1
            lneTxt = ""
        End If
    Else
        lneTxt = lneTxt & Mid(txtControl.Text, a, 1)
    End If
    If a = Len(txtControl.Text) Then
        lneCount = lneCount + 1
        If lneCount = theLine Then
            txtControl.SelStart = a - Len(lneTxt)
            txtControl.SelLength = Len(lneTxt)
            Exit Sub
        End If
    End If
Next a

End Sub
