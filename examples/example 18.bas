'EXAMPLE #18 - Listbox/Combobox Function Test

var text as string
var idx as number

window "win", "Listbox Test", normal, 100, 100, 500, 350
control "list1", "win", "", listbox, 100, 80, 150, 200
control "button1", "win", "Get Selected Item", button, 100, 20, 120, 25
control "button2", "win", "Get Item Count", button, 100, 50, 100, 25

event "button1", "click", getSel
event "button2", "click", getCount

AddItem "list1", "One"
AddItem "list1", "Two"
AddItem "list1", "Three"

pause


sub getSel
    idx = GetSelIdx("list1")
    text = GetItem("list1", idx)
    message text, ""
end sub


sub getCount
    message ItemCount("list1"), ""
end sub
