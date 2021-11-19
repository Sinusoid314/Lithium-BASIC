'EXAMPLE #7 - Tax Calculator

var amount as number
var tax as number
var total as number

showconsol

print "Example # 7 - Tax calculator"
print: print

input "Enter an amount: ", amount

tax = int(amount) * .06
total = amount + tax

print
print "Tax is " & tax
print "Total amount is " & total