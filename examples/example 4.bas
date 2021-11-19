'EXAMPLE #4 - Determine if a number is greater
'             than, less than or equal to 10

var num as number
var quit as string

showconsol

print "Example #4"

while lower(quit) <> "n"
  print: print
  input "Enter a numer: ", num
  print
  if num < 10 then 
    print "Number is less than 10"
  elseif num > 10 then
    print "Number is greater than 10"
  else
    print "Number is 10"
  end if
  print
  input "Continue? (y/n): ", quit
wend

print: print
print "Done."
