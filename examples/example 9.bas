'EXAMPLE #9 - Shows how to handle files in Lithium BASIC

var filename as string
var n as number

showconsol

filename = FileOpen("Open a file...", "Any file (*.*) | *.*")

if filename = "" then end

open filename for input as #file
    for n = 1 to lof(#file)
        print input(#file, 1)
    next n
close #file
