'EXAMPLE #14 - GUI Control Test

window "win1", "Window 1", normal, 100, 100, 600, 400

control "btn1", "win1", "Button 1", button, 10, 10, 72, 25
control "txt1", "win1", "", texteditor, 10, 50, 250, 200
control "txt2", "win1", "", textbox, 10, 300, 200, 20
control "lst1", "win1", "", listbox, 300, 10, 100, 200
control "cmbox1", "win1", "", combobox, 430, 20, 150, 200
control "drwbox1", "win1", "", drawbox, 300, 220, 200, 150
control "stctxt1", "win1", "GUI Controls Test", statictext, 420, 100, 200, 20

event "btn1", "click", Close

additem "cmbox1", "Item 1"

pause


sub Close
  var answer as string
  question "Are you sure you want to quit?", "Test GUI", answer
  if answer = "yes" then end
end sub