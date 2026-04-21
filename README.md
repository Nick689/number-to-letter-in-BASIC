## Usage
```
sub test
	msgbox(French(123456))
end sub
```

## Result
<img width="922" height="278" alt="Capture d’écran du 2026-04-21 08-28-09" src="https://github.com/user-attachments/assets/e34f20ce-9aae-4223-8c05-bd91d4dc343e" />


## Version choosing
- English1.bas is written in 53 lines, it pass the test in 182sec, it is easy to read
- English2.bas is written in 108 lines, it pass the test in 141sec, it is optimised for speed
- French.bas is based on English2.bas
- Spanish1.bas is written in 79 lines, it pass the test in 187sec, it is based on English1.bas
- Spanish2.bas is written in 162 lines, it pass the test in 149sec, it is based on French.bas
- Please propose your algorithm if you have faster or other language


## Speed testing method
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
