- English1 complete the test in 156sec
- English2 complete the test in 189sec, it is more easy to understand, but as it use more string function it is slower than English1
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
