Attribute VB_Name = "RNDMOD"
' Generating random numbers within limits and placing in an array
'   In this case, RandomSize number of random numbers between 45 & 12000
' *************************************************************

minr = 45          ' Lower limit of range of numbers
maxr = 12000          ' Upper limit of range of numbers
randno = minr + Fix(Rnd * (maxr - minr + 1))

 ' ************ A more complete example *****************

Option Base 1
MyLower = Val(RandomStart)          ' The lower limit of the random number range
MyUpper = Val(RandomEnd)           ' Upper limit of the random number range
MySample = Val(RandomSize)         ' The number of random numbers to generate
For r = 1 To MySample
     randno = MyLower + Fix(Rnd * (MyUpper - MyLower + 1))
     ReDim Preserve RandomNumberArray(r)
     RandomNumberArray(r) = LTrim(Str(randno))  ' An array of our random numbers
Next

