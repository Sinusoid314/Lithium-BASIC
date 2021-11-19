'EXAMPLE #31 - Cheese Master Program
'Written by Trish

var num as number

showconsol
@start
print "Are you the cheese master"
print: print

input "Enter number of cheeses in fridge: ", num
print

if num < 5 then goto @cheeser
if num = 5 then goto @gouda
if num > 5 then goto @bri

@cheeser
  print "you do not appreciate cheese at all. It is absolutley pathetic. You need to go to the store right now and buy at least 5 different types of cheeses and write a poem about it...then you will appreciate cheese"
  goto @start

@gouda
  print "Ok i can deal with you. You are not a cheeze wiz..haha..but you enjoy cheese every once in a while. It is about time, however, that you should start writing poems about cheese."
  goto @start

@bri
  print "YES! you are the ulitmate cheese wiz and not the artificle kind..the good stuff.Obviously i am interupting your cheese studies so please continue to be gouda..haha..and have some bri!"
  goto @start