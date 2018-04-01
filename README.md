- English1 pass the test in 188sec, it is compact and easy to read, it is written in 54 lines
- English2 pass the test in 133sec, is speed optimized by using minimaly arrays and string functions
- Please propose your algorithm if you have faster

## test method:
 
```
sub test
	dim i as double
	dim numtest as string
	dim start as date
	start=now()
	for i=0 to 999999
	numtest=numtoletter1(i)
	next
	msgbox(datediff("s",start,now()))
end sub
```
