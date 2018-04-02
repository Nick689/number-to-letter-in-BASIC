- English1 pass the test in 188sec, it is compact and easy to read, it is written in 55 lines
- English2 pass the test in 156sec, is littlely optimised but stays compact
- English3 pass the test in 133sec, is speed optimized by using minimaly arrays and string functions
- Others languages are based on English2 algorithm
- Please propose your algorithm if you have faster or other language

## test method:
 
```
sub test
	dim i as double
	dim numtest as string
	dim start as date
	start=now()
	for i=0 to 999999
		numtest=english2(i)
	next
	msgbox(datediff("s",start,now()))
end sub
```
