'EXAMPLE #25 - Name Command Test

var v as number

showconsol

open DefPath & "\old.dat" for output as #file
    print #file, "ALL YOUR BASE ARE BELONG TO US"
close #file

input "Rename>>", v

Name DefPath & "\old.dat" As DefPath & "\new.dat"
