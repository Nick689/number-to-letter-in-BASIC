- English1.bas pass the test in 182sec, it is compact and easy to read, it is written in 53 lines
- English2.bas pass the test in 141sec, it is optimised for speed
- French.bas pass the test in 143ec, it is based on English2.bas
- Spanish1.bas pass the test in 187sec, it is a mix of English1.bas and English2.bas
- Spanish2.bas pass the test in 149sec, it is based on French.bas
- Please propose your algorithm if you have faster or other language

## Keys to improve Basic program speed:
- loop-test your program like below
- Limit the use of string's functions, calculation deduction can sometime be a faster alternative (compare English1 to English2)
- Limit the use of arrays, prefer the use of SELECT CASE instead
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
