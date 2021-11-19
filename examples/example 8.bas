'EXAMPLE #8 - Shows how variable binding using BindVar works

var v1 as string
var v2 as string
var v3 as string

showconsol

v1 = "Cheese"
v2 = "Rice"
v3 = "Peppers"

print "Before:"
print "    " & v1
print "    " & v2
print "    " & v3
print

bindvar v2 to v1
bindvar v3 to v1 
unbindvar v2 from v1
unbindvar v3 from v1

v1 = "nil"

print "After:"
print "    " & v1
print "    " & v2
print "    " & v3