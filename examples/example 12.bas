'EXAMPLE #12 - Test of ON...GOTO/GOSUB command

var labelNum as number

showconsol

input "Please choose a branch label: ", labelNum
print

On labelNum GoSub @one, @two, @three

print
print "Return"
end


@one
  print "Label One"
return

@two
  print "Label Two"
return

@three
  print "Label Three"
return