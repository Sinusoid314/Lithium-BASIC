'EXAMPLE #19 - Text Editor/Box Function Test

var startPos, endPos as number

window "win", "Text Editor Test", normal, 100, 100, 500, 350
control "edit1", "win", "", texteditor, 100, 70, 200, 200
control "button1", "win", "Get Selected Text", button, 100, 20, 120, 25
control "button2", "win", "Get Line Count", button, 230, 20, 110, 25
control "button3", "win", "Get Line 3 Text", button, 350, 20, 110, 25
control "button4", "win", "Set Selected Text", button, 350, 70, 110, 25

event "button1", "click", getSelText
event "button2", "click", getLineNum
event "button3", "click", getLineText
event "button4", "click", setSelText

pause


sub getSelText
    message GetSelText("edit1"), ""
end sub


sub getLineNum
    message LineCount("edit1"), ""
end sub


sub getLineText
    message GetLineText("edit1", 3), ""
end sub


sub setSelText
    SetSelText "edit1", "****"
end sub
