'EXAMPLE #26 - GetDirs and GetFiles Command Test

array dirList(0) as string
array fileList(0) as string
var n as number
var path as string

showconsol

input "Enter directory: ", path

getdirs path, dirList(
getfiles path, fileList(

print

print dirList(0) & " directorie(s)"
print "---------------------"
for n = 1 to val(dirList(0))
    print dirList(n)
next n

print

print fileList(0) & " file(s)"
print "---------------------"
for n = 1 to val(fileList(0))
    print fileList(n)
next n
