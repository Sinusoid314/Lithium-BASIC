'Lithium Music Player v1.0
'Written in the Lithium BASIC programming langauge
'by Andrew Sturges (andrew@britcoms.com)

'Set the default path (if running from the editor)
DefPath = DefPath & "\examples\music player\"

'Define the global variables
var winLeft, winTop, winWidth, winHeight as number
var sName as string

'Load the images
loadimg "playImg", DefPath & "play.jpg"
loadimg "stopImg", DefPath & "stop.jpg"
loadimg "openImg", DefPath & "open.jpg"

'Set the window's size
winWidth = 360
winHeight = 150

'Make sure the window is centered
winLeft = (ScreenWidth-winWidth)/2
winTop = (ScreenHeight-winHeight)/2

'Open the window
window "win", "Lithium Music Player", normal, winLeft, winTop, winWidth, winHeight

'Create the controls
control "playBtn", "win", "", picbutton, 15, 10, 100, 30
control "stopBtn", "win", "", picbutton, 125, 10, 100, 30
control "openBtn", "win", "", picbutton, 235, 10, 100, 30
control "status", "win", "", statictext, 10, 95, 340, 20

backcolor "win", ButtonFace
drawtext "win", "Now Playing:", 10, 60
stick "win"

'Set the button images
buttonimg "playBtn", "playImg"
buttonimg "stopBtn", "stopImg"
buttonimg "openBtn", "openImg"

'Set up the window and control events
event "win", "close", Exit
event "openBtn", "click", OpenFile
event "playBtn", "click", Play
event "stopBtn", "click", Stop

pause


'Play the sound file
sub Play
    enable "stopBtn"
    disable "playBtn"
    playsound "sound"
    timer "scroller", 300, ScrollName
end sub


'Stop playing the sound file
sub Stop
    enable "playBtn"
    disable "stopBtn"
    stopsound "sound"
    stoptimer "scroller"
end sub


'Open a sound file to play
sub OpenFile
    var filename as string
    filename = fileopen("Open a sound file...", "Sound files (*.wav,*.mid) | *.wav; *.mid")
    if filename <> "" then
        settext "status", space(4) & FileTitle(filename) & space(4)
        closesound "sound"
        opensound "sound", filename
        call Play
    end if
end sub


sub Exit
    closesound "sound"
    unloadimg "playImg"
    unloadimg "stopImg"
    unloadimg "openImg"
    end
end sub


sub ScrollName
    var tmpChar, tmpText as string
    tmpText = gettext("status")
    tmpChar = left(tmpText, 1)
    tmpText = mid(tmpText, 2) & tmpChar
    settext "status", tmpText
end sub


function FileTitle(pathStr as string) as string
    var newStr, tmpChar as string
    var n as number
    n=len(pathStr)
    tmpChar = mid(pathStr, n, 1)
    while tmpChar <> "\"
        newStr = tmpChar & newStr
        n = n - 1
        tmpChar = mid(pathStr, n, 1)
    wend
    n = len(newStr)
    tmpChar = ""
    while tmpChar <> "."
        tmpChar = mid(newStr, n, 1)
        n = n - 1
    wend
    newStr = left(newStr, n)
    FileTitle = newStr
end function
