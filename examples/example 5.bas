'EXAMPLE #5 - HILO guessing game

var num as number
var guess as number
var guessCount as number
var playAgain as string

consoltitle "Example #5 - HILO Guessing Game"

showconsol

while lower(trim(playAgain)) <> "n"
  num = int(rnd()*100)+1
  guessCount = 0
  while guess <> num
    input "Guess a number from 1 to 100: ", guess
    if guess > num then
      print "Guess lower.": print
    elseif guess < num then
      print "Guess higher.": print
    end if
    guessCount = guessCount + 1
  wend
  print "Correct!"
  print
  print "It took you " & guessCount & " guess(es)."
  print: print
  input "Play agian? (y/n): ", playAgain
  cls
wend

print "Done."
