
window "win", "Color Blocks", normal, 100, 100, 500, 400
control "draw", "win", "", drawbox, 10, 10, 470, 350

event "draw", "leftbuttondown", BtnDown

pause


sub BtnDown
    var x, y, color as number
    color = RandColor()
    backcolor "draw", color
    forecolor "draw", color
    getmousexy x, y
    box "draw", x-5, y-5, 10, 10
end sub


function RandColor() as number
    var r, g, b as number
    r = int(rnd()*255)+1
    g = int(rnd()*255)+1
    b = int(rnd()*255)+1
    RandColor = rgb(r,g,b)
end function
