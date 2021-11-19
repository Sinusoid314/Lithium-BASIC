'EXAMPLE #6 - Determine if a number if less than, greater than, or
'             equal to 0 using GOTOs

var num as number

showconsol

print "Example #6"
print: print

input "Enter an integer: ", num
print

if num < 0 then goto @lessThan
if num = 0 then goto @equalTo
if num > 0 then goto @greaterThan

@lessThan
  print "Less than 0"
  end

@equalTo
  print "Equal to 0"
  end

@greaterThan
  print "Greater than 0"
  end