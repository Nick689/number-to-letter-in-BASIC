function numtoletter(numValue as double) as string
	dim numString as string
	dim numTable(12) as integer
	dim tensTable(20) as string
	dim unitsTable(10) as string
	dim negative as boolean
	dim leng as integer
	dim outputString as string
	dim i as integer
	unitsTable=array("","ONE","TWO","THREE","FOUR","FIVE","SIX","SEVEN","EIGHT","NINE","TEN","ELEVEN","TWELVE","THIRTEEN","FOURTEEN","FIFTEEN","SIXTEEN","SEVENTEEN","EIGHTEEN","NINETEEN")
	tensTable=array("","","TWENTY","THIRTY","FORTY","FIFTY","SIXTY","SEVENTY","EIGHTY","NINETY")
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
		numTable(i-1)=cint(mid(numString,i,1))
		leng=numTable(i-1)
	next
	if numTable(0)+numTable(1)+numTable(2) then
		if numTable(0) then
			select case numTable(1)
			case 0: outputString=unitsTable(numTable(0))+" HUNDRED "+unitsTable(numTable(2))+" BILLION"
			case 1: outputString=unitsTable(numTable(0))+" HUNDRED "+unitsTable(10+numTable(2))+" BILLION"
			case else: outputString=unitsTable(numTable(0))+" HUNDRED "+tensTable(numTable(1))+" "+unitsTable(numTable(2))+" BILLION"
			end select
		else
			select case numTable(1)
			case 0: if numTable(3) then outputString=unitsTable(numTable(3))+" BILLION"
			case 1: outputString=unitsTable(10+numTable(3))+" BILLION"
			case else: outputString=tensTable(numTable(2))+" "+unitsTable(numTable(3))+" BILLION"
			end select
		endif
	endif
	if numTable(3)+numTable(4)+numTable(5) then
		if outputString<>"" then outputString=outputString+" "
		if numTable(3) then
			select case numTable(4)
			case 0:	outputString=outputString+unitsTable(numTable(3))+" HUNDRED "+unitsTable(numTable(5))+" MILLION"
			case 1: outputString=outputString+unitsTable(numTable(3))+" HUNDRED "+unitsTable(10+numTable(5))+" MILLION"
			case else: outputString=outputString+unitsTable(numTable(3))+" HUNDRED "+tensTable(numTable(4))+" "+unitsTable(numTable(5))+" MILLION"
			end select
		else
			select case numTable(4)
			case 0: if numTable(5) then outputString=outputString+unitsTable(numTable(5))+" MILLION"
			case 1: outputString=outputString+unitsTable(10+numTable(5))+" MILLION"
			case else: outputString=outputString+tensTable(numTable(4))+" "+unitsTable(numTable(5))+" MILLION"
			end select
		endif
	endif
	if numTable(6)+numTable(7)+numTable(8) then
		if outputString<>"" then outputString=outputString+" "
		if numTable(6) then
			select case numTable(7)
			case 0: outputString=outputString+unitsTable(numTable(6))+" HUNDRED "+unitsTable(numTable(8))+" THOUSAND"
			case 1: outputString=outputString+unitsTable(numTable(6))+" HUNDRED "+unitsTable(10+numTable(8))+" THOUSAND"
			case else: outputString=outputString+unitsTable(numTable(6))+" HUNDRED "+tensTable(numTable(7))+" "+unitsTable(numTable(8))+" THOUSAND"
			end select
		else
			select case numTable(7)
			case 0: if numTable(8) then	outputString=outputString+unitsTable(numTable(8))+" THOUSAND"
			case 1: outputString=outputString+unitsTable(10+numTable(8))+" THOUSAND"
			case else: outputString=outputString+tensTable(numTable(7))+" "+unitsTable(numTable(8))+" THOUSAND"
			end select
		endif
	endif
	if numTable(9)+numTable(10)+numTable(11) then
		if outputString<>"" then outputString=outputString+" "
		if numTable(9) then
			select case numTable(10)
			case 0: outputString=outputString+unitsTable(numTable(9))+" HUNDRED "+unitsTable(numTable(11))
			case 1: outputString=outputString+unitsTable(numTable(9))+" HUNDRED "+unitsTable(10+numTable(11))
			case else: outputString=outputString+unitsTable(numTable(9))+" HUNDRED "+tensTable(numTable(10))+" "+unitsTable(numTable(11))
			end select
		else
			select case numTable(10)
			case 0: outputString=outputString+unitsTable(numTable(11))
			case 1: outputString=outputString+unitsTable(10+numTable(11))
			case else: outputString=outputString+tensTable(numTable(10))+" "+unitsTable(numTable(11))
			end select
		endif
	endif
	if numValue=0 then outputString="ZERO"
	if negative then outputString="MINUS "+outputString
	numtoletter=outputString
end function
