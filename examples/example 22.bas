'EXAMPLE #22 - Sound Commands Test

var v as number

showconsol

beep

opensound "midi", DefPath & "\examples\the guess who - no time.mid"

input "Play>> ", v

beep

playsound "midi"

input "Stop>> ", v

closesound "midi"

beep
