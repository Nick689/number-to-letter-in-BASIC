English1 take 182sec to pass the test
English2 take 249sec to pass the test
English3 take 165sec to pass the test
English4 take 331sec to pass the test
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
