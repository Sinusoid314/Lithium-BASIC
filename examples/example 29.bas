'EXAMPLE #29 - Functions Test #2

var n as number

showconsol

loadimg "pic", DefPath & "\examples\get.bmp"
print "HBMP('pic') = " & hbmp("pic")
unloadimg "pic"
print

print "FileOpen - " & FileOpen("Open...","Text Files (*.txt) | *.txt")
print "FileSave - " & FileSave("Save...","Code Files (*.txt,*.bas,*.bak) | *.txt;*.bas;*.bak")
print

open DefPath & "\examples\funcTest.txt" for input as #file
    seek #file, 2
    input #file, n
    print "LOF(#file) = " & lof(#file)
    print "EOF(#file) = " & eof(#file)
close #file
