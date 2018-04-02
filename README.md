- English1 pass the test in 188sec, it is compact and easy to read, it is written in 57 lines
- English2 pass the test in 156sec, is littlely optimised but stays compact
- English3 pass the test in 133sec, is speed optimized by using minimaly arrays and string functions
- French pass the test in 190sec, it is based on English2
- Spanish pass the test in 187sec, it is based on English2
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
