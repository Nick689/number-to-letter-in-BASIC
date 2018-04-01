- English1 complete the test in 156sec
- English2 complete the test in 189sec it use a little more string function that slow down the script
- Please propose your algorithm

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
