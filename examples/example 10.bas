'EXAMPLE #10 - Sub and function call test

showconsol

print "Example #10 - JERK Program (originally from C for Dummies)"
print
print

print "He calls me on the phone with nothing to say"
print "Not once, or twice, but three times a day!"
call Jerk 1

print "He insulted my wife, my cat, my mother"
print "He irritates and grates, like no other!"
call Jerk 2

print "He chuckles it off, his big belly a-heavin'"
print "But he won't be laughing when I get even!"
call Jerk 3


sub Jerk jerkNum as number
  var a as number
  print
  for a = 1 to jerkNum
    print "Bill is a jerk."
  next a
  print
end sub
