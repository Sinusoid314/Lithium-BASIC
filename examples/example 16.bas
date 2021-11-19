'EXAMPLE #16 - Window Size and Position Test

var left, top, width, height as number

window "win1", "Window 1", normal, 100, 100, 600, 400

control "btn1", "win1", "Get Position", button, 100, 100, 120, 25
control "btn2", "win1", "Get Size", button, 100, 130, 120, 25
control "btn3", "win1", "Set Position", button, 100, 200, 120, 25
control "btn4", "win1", "Set Size", button, 100, 230, 120, 25

control "left", "win1", "100", textbox, 300, 150, 50, 20 
control "top", "win1", "100", textbox, 370, 150, 50, 20 
control "width", "win1", "600", textbox, 300, 180, 50, 20 
control "height", "win1", "400", textbox, 370, 180, 50, 20 

event "btn1", "click", getPos
event "btn2", "click", getSize
event "btn3", "click", setPos
event "btn4", "click", setSize

pause


sub getPos
  GetXY "win1", left, top
  SetText "left", left
  SetText "top", top
end sub

sub getSize
  GetSize "win1", width, height
  SetText "width", width
  SetText "height", height
end sub

sub setPos
  left = val(GetText("left"))
  top = val(GetText("top"))
  SetXY "win1", left, top
end sub

sub setSize
  width = val(GetText("width"))
  height = val(GetText("height"))
  SetSize "win1", width, height
end sub
