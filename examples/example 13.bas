'EXAMPLE #13 - Text and background color test

var n as number

showconsol

@start

input "Press enter to change colors...", n
textcolor RandColor()
bgcolor RandColor()
cls

goto @start


'Generate a random color
function RandColor() as number
    var r, g, b as number
    r = int(rnd()*255)+1
    g = int(rnd()*255)+1
    b = int(rnd()*255)+1
    RandColor = rgb(r, g, b)
end function
