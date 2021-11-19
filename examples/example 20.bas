'EXAMPLE #20 - Graphic Functions Test

window "win", "Graphics Test", normal, 100, 100, 600, 400
control "draw1", "win", "", drawbox, 100, 50, 300, 300
control "btn1", "win", "Refresh", button, 450, 100, 80, 25
control "btn2", "win", "Clear", button, 450, 150, 80, 25

event "btn1", "click", RefreshDraw
event "btn2", "click", ClearWin

loadimg "img1", DefPath & "\examples\image1.bmp"

backcolor "draw1", Orange
forecolor "draw1", Blue
linesize "draw1", 5

box "draw1", 10, 10, 100, 100
circle "draw1", 60, 60, 50
line "draw1", 60, 60, 110, 110
setpixel "draw1", 100, 20, Brown

drawtext "draw1", "ALL YOUR BASE ARE BELONG TO US", 100, 100

stick "draw1"

drawimg "draw1", "img1", 10, 10

getimg "draw1", "img2", 50, 50, 50, 50
drawimg "draw1", "img2", 200, 200

unloadimg "img1"
unloadimg "img2"

pause


sub RefreshDraw
    refresh "draw1"
end sub


sub ClearWin
    clear "draw1"
end sub