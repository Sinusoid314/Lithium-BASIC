'EXAMPLE #17 - Menu Test

window "win1", "Window 1", normal, 100, 100, 600, 400

menu "win1", "&File", "&New", New, "&Open", Open, "&Save", Save, |, "E&xit", Exit
menu "win1", "&Edit", "Cu&t", Cut, "&Copy", Copy, "&Paste", Paste

pause


sub New
  message "File -> New", "Menu Test"
end sub

sub Open
  message "File -> Open", "Menu Test"
end sub

sub Save
  message "File -> Save", "Menu Test"
end sub

sub Cut
  message "Edit -> Cut", "Menu Test"
end sub

sub Copy
  message "Edit -> Copy", "Menu Test"
end sub

sub Paste
  message "Edit -> Paste", "Menu Test"
end sub

sub Exit
  end
end sub