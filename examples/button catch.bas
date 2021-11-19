'Button Catch - Try to click on the button before it moves

var moveNum as number

moveNum = 1

window "win", "Catch the Button!", normal, 100, 100, 500, 400
control "btn1", "win", "Click Me", button, 200, 150, 100, 25

event "btn1", "mousemove", MoveButton
event "btn1", "click", CatchButton

pause


sub MoveButton
    var newX, newY, n as number
    for n = 1 to moveNum
        newX = int(rnd()*400)+1
        newY = int(rnd()*300)+1
        setxy "btn1", newX, newY
    next n
end sub


sub CatchButton
    message "Bah! You will pay for this clicking!", "Button Caught"
    moveNum = moveNum + 2
end sub
