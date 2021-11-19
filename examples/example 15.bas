'EXAMPLE #15 - SetText and GetText() Test

window "win1", "Window 1", normal, 100, 100, 600, 400

control "btn1", "win1", "Change Caption", button, 260, 100, 120, 25
control "text1", "win1", "", textbox, 50, 100, 200, 20

event "btn1", "click", changeText

pause


sub changeText
  var text as string
  text = GetText("text1")
  SetText "win1", text
end sub