- English1 pass the test in 188sec, it is compact and easy to read, it is written in 53 lines
- English2 pass the test in 150sec, is littlely optimised but stays compact
- English3 pass the test in 100sec, is speed optimized by using minimaly arrays and string functions
- French pass the test in 181sec, it is based on English2
- Spanish pass the test in 180sec, it is based on English2
- Please propose your algorithm if you have faster or other language

## SPEED's keys to fast basic script:
- Limit the use of string's functions, calculation deduction can sometime be a faster alternative (compare English1 to English2)
- Limit the use of arrays
- Declare Variant instead of array
- Write IF-THEN-ELSE on one line when possible (no ENDIF at the end)
- Prefer IF-THEN-ELSE on one line to SELECT-CASE

- List every case with a SELECT CASE instead of using an array (this is extreme optimisation like in English3)

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
