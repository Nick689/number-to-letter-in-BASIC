- English1 complete the test in 182sec
- English2 complete the test in 249sec
- English3 complete the test in 65sec
- English4 complete the test in 331sec
.
There's an performance issue with array in LibreOffice
.
.
.
Please propose your algorithm
.
.
.
## test method:
.
```
sub test
	dim i as double
	dim numtest as string
	dim start as date
	start=now()
	for i=0 to 999999
	numtest=letter1(i)
	next
	msgbox(datediff("s",start,now()))
end sub
```
