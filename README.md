English1 (182sec) is faster than English2 (249sec)

Please propose your algorithm
.
.
.

## test method:

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
