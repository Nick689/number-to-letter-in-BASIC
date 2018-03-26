global function numtoletter(numValue as double) as string
	dim numString as string
	dim numberTable(12) as integer
	dim tenTable(20) as string
	dim unitTable(10) as string
	dim negative as boolean
	dim leng as integer
	dim outputString as string
	dim i as integer
	unitTable=array("","ONE","TWO","THREE","FOUR","FIVE","SIX","SEVEN","EIGHT","NINE","TEN","ELEVEN","TWELVE","THIRTEEN","FOURTEEN","FIFTEEN","SIXTEEN","SEVENTEEN","EIGHTEEN","NINETEEN")
	tenTable=array("","","TWENTY","THIRTY","FORTY","FIFTY","SIXTY","SEVENTY","EIGHTY","NINETY")
	if numValue>999999999999 then
		numtoletter=""
		exit function
	endif
	if numValue<0 then
		negative=true
		numValue=abs(numValue)
	else
		negative=false
	endif
	numString=cstr(numValue)
	leng=len(numString)
	while leng<12
		numString="0"+numString
		leng=leng+1
	wend
	for i=1 to 12
		numberTable(i-1)=cint(mid(numString,i,1))
		leng=numberTable(i-1)
	next
	if numberTable(0)+numberTable(1)+numberTable(2) then
		if numberTable(0) then
			select case numberTable(1)
			case 0: outputString=unitTable(numberTable(0))+" HUNDRED "+unitTable(numberTable(2))+" BILLION"
			case 1: outputString=unitTable(numberTable(0))+" HUNDRED "+unitTable(10+numberTable(2))+" BILLION"
			case else: outputString=unitTable(numberTable(0))+" HUNDRED "+tenTable(numberTable(1))+" "+unitTable(numberTable(2))+" BILLION"
			end select
		else
			select case numberTable(1)
			case 0: if numberTable(3) then outputString=unitTable(numberTable(3))+" BILLION"
			case 1: outputString=unitTable(10+numberTable(3))+" BILLION"
			case else: outputString=tenTable(numberTable(2))+" "+unitTable(numberTable(3))+" BILLION"
			end select
		endif
	endif
	if numberTable(3)+numberTable(4)+numberTable(5) then
		if outputString<>"" then outputString=outputString+" "
		if numberTable(3) then
			select case numberTable(4)
			case 0:	outputString=outputString+unitTable(numberTable(3))+" HUNDRED "+unitTable(numberTable(5))+" MILLION"
			case 1: outputString=outputString+unitTable(numberTable(3))+" HUNDRED "+unitTable(10+numberTable(5))+" MILLION"
			case else: outputString=outputString+unitTable(numberTable(3))+" HUNDRED "+tenTable(numberTable(4))+" "+unitTable(numberTable(5))+" MILLION"
			end select
		else
			select case numberTable(4)
			case 0: if numberTable(5) then outputString=outputString+unitTable(numberTable(5))+" MILLION"
			case 1: outputString=outputString+unitTable(10+numberTable(5))+" MILLION"
			case else: outputString=outputString+tenTable(numberTable(4))+" "+unitTable(numberTable(5))+" MILLION"
			end select
		endif
	endif
	if numberTable(6)+numberTable(7)+numberTable(8) then
		if outputString<>"" then outputString=outputString+" "
		if numberTable(6) then
			select case numberTable(7)
			case 0: outputString=outputString+unitTable(numberTable(6))+" HUNDRED "+unitTable(numberTable(8))+" THOUSAND"
			case 1: outputString=outputString+unitTable(numberTable(6))+" HUNDRED "+unitTable(10+numberTable(8))+" THOUSAND"
			case else: outputString=outputString+unitTable(numberTable(6))+" HUNDRED "+tenTable(numberTable(7))+" "+unitTable(numberTable(8))+" THOUSAND"
			end select
		else
			select case numberTable(7)
			case 0: if numberTable(8) then	outputString=outputString+unitTable(numberTable(8))+" THOUSAND"
			case 1: outputString=outputString+unitTable(10+numberTable(8))+" THOUSAND"
			case else: outputString=outputString+tenTable(numberTable(7))+" "+unitTable(numberTable(8))+" THOUSAND"
			end select
		endif
	endif
	if numberTable(9)+numberTable(10)+numberTable(11) then
		if outputString<>"" then outputString=outputString+" "
		if numberTable(9) then
			select case numberTable(10)
			case 0: outputString=outputString+unitTable(numberTable(9))+" HUNDRED "+unitTable(numberTable(11))
			case 1: outputString=outputString+unitTable(numberTable(9))+" HUNDRED "+unitTable(10+numberTable(11))
			case else: outputString=outputString+unitTable(numberTable(9))+" HUNDRED "+tenTable(numberTable(10))+" "+unitTable(numberTable(11))
			end select
		else
			select case numberTable(10)
			case 0: outputString=outputString+unitTable(numberTable(11))
			case 1: outputString=outputString+unitTable(10+numberTable(11))
			case else: outputString=outputString+tenTable(numberTable(10))+" "+unitTable(numberTable(11))
			end select
		endif
	endif
	if numValue=0 then outputString="ZERO"
	if negative then outputString="MINUS "+outputString
	numtoletter=outputString
end function
