'Lithium Notepad v1.0
'written in Lithium BASIC v1.0
'by Andrew Sturges (andrew@britcoms.com)

'Set the default path (if running from the editor)
DefPath = DefPath & "\examples\notepad\"

'Define the global variables
var winLeft, winTop, winWidth, winHeight as number
var editText, filename as string

'Load images
loadimg "newImg", DefPath & "new.jpg"
loadimg "openImg", DefPath & "open.jpg"
loadimg "saveImg", DefPath & "save.jpg"
loadimg "cutImg", DefPath & "cut.jpg"
loadimg "copyImg", DefPath & "copy.jpg"
loadimg "pasteImg", DefPath & "paste.jpg"
loadimg "helpImg", DefPath & "help.jpg"

'Set the window's size
winWidth = ScreenWidth - 300
winHeight = ScreenHeight - 200

'Make sure the window is centered
winLeft = (ScreenWidth-winWidth)/2
winTop = (ScreenHeight-winHeight)/2

'Open the window
window "win", "Lithium Notepad", normal, winLeft, winTop, winWidth, winHeight

'Create the controls
control "editor", "win", "", texteditor, 0, 30, winWidth-8, winHeight-76
control "newBtn", "win", "", picbutton, 10, 5, 25, 23
control "openBtn", "win", "", picbutton, 35, 5, 25, 23
control "saveBtn", "win", "", picbutton, 60, 5, 25, 23
control "cutBtn", "win", "", picbutton, 95, 5, 25, 23
control "copyBtn", "win", "", picbutton, 120, 5, 25, 23
control "pasteBtn", "win", "", picbutton, 145, 5, 25, 23
control "helpBtn", "win", "", picbutton, 180, 5, 25, 23

'Create the menus
menu "win", "&File", "&New", NewFile, "&Open", OpenFile, "&Save", SaveFile, _
    "Save &As...", SaveAs, |, "E&xit", Exit
menu "win", "&Edit", "Cu&t", Cut, "&Copy", Copy, "&Paste", Paste
menu "win", "&Help", "&About Lithium Notepad...", About

'Set button images
buttonimg "newBtn", "newImg"
buttonimg "openBtn", "openImg"
buttonimg "saveBtn", "saveImg"
buttonimg "cutBtn", "cutImg"
buttonimg "copyBtn", "copyImg"
buttonimg "pasteBtn", "pasteImg"
buttonimg "helpBtn", "helpImg"

'Set up the window and control events
event "win", "close", Exit
event "win", "resize", Resize
event "newBtn", "click", NewFile
event "openBtn", "click", OpenFile
event "saveBtn", "click", SaveFile
event "cutBtn", "click", Cut
event "copyBtn", "click", Copy
event "pasteBtn", "click", Paste
event "helpBtn", "click", About

pause


'Clear the text editor
sub NewFile
    var a, tmpText as string
    tmpText = gettext("editor")
    if editText <> tmpText then
        question "Save changes made to '" & filename & "'?", "Lithium Notepad", a
        if a = "yes" then call SaveFile
    end if
    settext "win", "Lithium Notepad"
    settext "editor", ""
    filename = ""
    editText = ""
end sub


'Open a text file
sub OpenFile
    var a, tmpText as string
    tmpText = gettext("editor")
    if editText <> tmpText then
        question "Save changes made to '" & filename & "'?", "Lithium Notepad", a
        if a = "yes" then call SaveFile
    end if
    filename = fileopen("Open...","Text file (*.txt) | *.txt")
    if filename <> "" then
        open filename for input as #file
            editText = input(#file, lof(#file))
        close #file
        settext "editor", editText
        settext "win", "Lithium Notepad - [" & filename & "]"
    end if
end sub


'Save the current text file
sub SaveFile
    var tmpText as string
    tmpText = gettext("editor")
    if filename = "" then
        filename = filesave("Save...", "Text file (*.txt) | *.txt")
        if filename <> "" then
            open filename for output as #file
                print #file, tmpText;
            close #file
            settext "win", "Lithium Notepad - [" & filename & "]"
            editText = tmpText
        end if
    else
        open filename for output as #file
            print #file, tmpText;
        close #file
        editText = tmpText
    end if
end sub


'Save the text to any file
sub SaveAs
    var tmpText, tmpFile as string
    tmpText = gettext("editor")
    tmpFile = filesave("Save As...", "Text file (*.txt) | *.txt")
    if tmpFile <> "" then
        open tmpFile for output as #file
            print #file, tmpText;
        close #file
    end if
end sub


'Close the program
sub Exit
    var a, tmpText as string
    tmpText = gettext("editor")
    if tmpText <> editText then
        question "Save changes made to '" & filename & "'?", "Lithium Notepad", a
        if a = "yes" then call SaveFile
    end if
    unloadimg "newImg"
    unloadimg "openImg"
    unloadimg "saveImg"
    unloadimg "cutImg"
    unloadimg "copyImg"
    unloadimg "pasteImg"
    unloadimg "helpImg"
    end
end sub


'Cut any selected text onto the clipboard
sub Cut
    var selText as string
    selText = getseltext("editor")
    setclipboardtext selText
    setseltext "editor", ""
end sub


'Copy any selected text onto the clipboard
sub Copy
    var selText as string
    selText = getseltext("editor")
    setclipboardtext selText
end sub


'Paste text from the clipboard onto the text editor
sub Paste
    var selText as string
    selText = getclipboardtext()
    setseltext "editor", selText
end sub


'Resize the controls to fit the window
sub Resize
    getsize "win", winWidth, winHeight
    setsize "editor", winWidth-8, winHeight-76
end sub


'Display info about Lithium Notepad
sub About
    loadimg "aboutImg", DefPath & "about.jpg"
    getxy "win", winLeft, winTop
    window "aboutWin", "About Lithium Notepad...", dialog_modal, winLeft+80, winTop+80, 350, 200
    control "aboutBtn1", "aboutWin", "OK", button, 130, 130, 100, 30
    backcolor "aboutWin", ButtonFace
    drawimg "aboutWin", "aboutImg", 10, 10
    drawtext "aboutWin", "Lithium Notepad v1.0", 100, 10
    drawtext "aboutWin", "Written in the Lithium BASIC", 100, 40
    drawtext "aboutWin", "programming language.", 100, 55
    drawtext "aboutWin", "© 2003 SirCodezAlot Software", 100, 90
    stick "aboutWin"
    event "aboutWin", "close", CloseAbout
    event "aboutBtn1", "click", CloseAbout
end sub

sub CloseAbout
    unloadimg "aboutImg"
    closewindow "aboutWin"
end sub
