'EXAMPLE #21 - Sprite Commands Test

var bgX, bgY as number

bgX = 0
bgY = 0

loadimg "img1", DefPath & "\examples\sprite1-1.bmp"
loadimg "img2", DefPath & "\examples\sprite1-2.bmp"
loadimg "img3", DefPath & "\examples\bgPic.bmp"

window "win", "Sprite Test", normal, 100, 100, 600, 400
control "draw1", "win", "", drawbox, 10, 10, 400, 350
'control "btn1", "win", "Draw Sprites", button, 420, 100, 110, 25
control "btn2", "win", "Remove Frame", button, 420, 70, 110, 25
control "btn3", "win", "Move Sprite", button, 420, 130, 110, 25
control "btn4", "win", "Resize Sprite", button, 420, 160, 110, 25
control "btn5", "win", "Hide Sprite", button, 420, 190, 110, 25
control "btn6", "win", "Pause Sprite", button, 420, 220, 110, 25

event "win", "close", Close
'event "btn1", "click", RedrawSprites
event "btn2", "click", RemoveFrame
event "btn3", "click", MoveSprite
event "btn4", "click", ResizeSprite
event "btn5", "click", HideSprite
event "btn6", "click", PauseSprite

sprite "draw1", "sprite1", 10, 10, "img1"
addframe "draw1", "sprite1", "img2"
spriterate "draw1", "sprite1", 300
spriterotate "draw1", "sprite1", "flip"

spritebg "draw1", "img3"
spritebgpos "draw1", bgX, bgY

timer "timer1", 100, RedrawSprites

pause


sub RedrawSprites
    drawsprites "draw1"
    bgX = bgX + 10
    bgY = bgY + 5
    if bgX > 1000 then bgX = 0
    if bgY > 1000 then bgY = 0
    spritebgpos "draw1", bgX, bgY
end sub


sub RemoveFrame
    delframe "draw1", "sprite1", 1
end sub


sub MoveSprite
    spritepos "draw1", "sprite1", 100, 50
end sub


sub ResizeSprite
    spritesize "draw1", "sprite1", 200, 50
end sub


sub HideSprite
    spritehide "draw1", "sprite1"
end sub


sub PauseSprite
    spritepause "draw1", "sprite1"
end sub


sub Close
    delsprite "draw1", "sprite1"
    unloadimg "img1"
    unloadimg "img2"
    unloadimg "img3"
    end
end sub
