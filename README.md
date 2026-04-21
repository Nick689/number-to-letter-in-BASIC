## Usage
```
sub test
	msgbox(French(123456))
end sub
```
result:<img width="922" height="278" alt="Capture d’écran du 2026-04-21 08-28-09" src="https://github.com/user-attachments/assets/e34f20ce-9aae-4223-8c05-bd91d4dc343e" />


## Performance
- English1.bas pass the test in 182sec, it is compact, easy to read, and written in 53 lines
- English2.bas pass the test in 141sec, it is optimised for speed
- French.bas pass the test in 143ec, it is based on English2.bas
- Spanish1.bas pass the test in 187sec, it is a mix of English1.bas and English2.bas
- Spanish2.bas pass the test in 149sec, it is based on French.bas
- Please propose your algorithm if you have faster or other language

# Test method 
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
