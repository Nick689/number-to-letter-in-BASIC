- English1 pass the test in 188sec, it is compact and easy to read, it is written in 53 lines
- English2 pass the test in 139sec, is littlely optimised but stays compact
- French pass the test in 141ec, it is based on English2
- Spanish pass the test in 180sec, it is based on English2
- Please propose your algorithm if you have faster or other language

## Keys to improve Basic program speed:
- loop-test your program like below
- Limit the use of string's functions, calculation deduction can sometime be a faster alternative (compare English1 to English2)
- Limit the use of arrays
- Write IF-THEN-ELSE on one line when possible (no ENDIF at the end)
- When theres not too many instance, write every instance instead of loop
- Declare Variant instead of Array

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
