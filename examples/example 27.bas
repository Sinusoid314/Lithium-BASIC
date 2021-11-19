'EXAMPLE #27 - PicButton, RadioButton, and CheckBox Control Test

loadimg "get", DefPath & "\examples\get.bmp"
loadimg "set", DefPath & "\examples\set.bmp"

window "win", "Control Test", normal, 100, 100, 600, 400
control "group1", "win", "ZEE CONTROLS:", groupbox, 10, 10, 500, 300
control "pbtn1", "win", "", picbutton, 100, 100, 72, 25
control "pbtn2", "win", "", picbutton, 200, 100, 72, 25
control "rbtn1", "win", "Radio Button", radiobutton, 100, 150, 100, 15
control "cbox1", "win", "Check Box", checkbox, 100, 200, 100, 15

buttonimg "pbtn1", "get"
buttonimg "pbtn2", "set"

event "pbtn1", "click", PBtn1Click
event "pbtn2", "click", PBtn2Click

pause


sub PBtn1Click
    message GetState("rbtn1"), ""
    message GetState("cbox1"), ""
end sub


sub PBtn2Click
    SetState "rbtn1", False
    SetState "cbox1", True
end sub
