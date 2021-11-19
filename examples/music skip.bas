'Music Skip - Jump around to different parts of a song

var v as string
var musicLen as number

showconsol

OpenSound "mid", DefPath & "\examples\the guess who - no time.mid"
PlaySound "mid"

musicLen = GetSoundLen("mid")

@skipPrompt
input "Press enter ('q' to quit): ", v
if v = "q" then goto @done
call ResetSound
goto @skipPrompt

@done
CloseSound "mid"

sub ResetSound
    var musicPos as number
    musicPos = int(rnd()*musicLen)+1
    SetSoundPos "mid", musicPos
end sub