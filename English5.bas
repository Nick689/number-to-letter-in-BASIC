function numtoletter(numValue as double) as string' 215s
	dim numString as string
	dim numTable(12) as integer
	dim tensTable(20) as string
	dim unitsTable(10) as string
	dim triplet0 as integer
	dim triplet1 as integer
	dim triplet2 as integer
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
	for i=0 to 11
		numTable(i)=cint(mid(numString,i+1,1))
	next
	if numTable(0)+numTable(1)+numTable(2) then
		triplet0=numTable(0)
		triplet1=numTable(1)
		triplet2=numTable(2)
		gosub routine
		outputString=outputString+" BILLION"
	endif
	if numTable(3)+numTable(4)+numTable(5) then
		triplet0=numTable(3)
		triplet1=numTable(4)
		triplet2=numTable(5)
		gosub routine
		outputString=outputString+" MILLION"
	endif
	if numTable(6)+numTable(7)+numTable(8) then
		triplet0=numTable(6)
		triplet1=numTable(7)
		triplet2=numTable(8)
		gosub routine
		outputString=outputString+" THOUSAND"
	endif
	if numTable(9)+numTable(10)+numTable(11) then
		triplet0=numTable(9)
		triplet1=numTable(10)
		triplet2=numTable(11)
		gosub routine
	endif
	if numValue=0 then outputString="ZERO"
	if negative then outputString="MINUS "+outputString
	numtoletter=outputString
	exit function
routine:
	if outputString<>"" then outputString=outputString+" "
	if triplet0 then
		select case triplet1
		case 0: if triplet2 then
				outputString=outputString+unitsTable(triplet0)+" HUNDRED "+unitsTable(triplet2)
			else
				outputString=outputString+unitsTable(triplet0)+" HUNDRED"
			endif
		case 1:	outputString=outputString+unitsTable(triplet0)+" HUNDRED "+unitsTable(10+triplet2)
		case else: if triplet2 then
				outputString=outputString+unitsTable(triplet0)+" HUNDRED "+tensTable(triplet1)+" "+unitsTable(triplet2)
			else
				outputString=outputString+unitsTable(triplet0)+" HUNDRED "+tensTable(triplet1)
			endif
		end select
	else
		select case triplet1
		case 0: outputString=outputString+unitsTable(triplet2)
		case 1: outputString=outputString+unitsTable(10+triplet2)
		case else: if triplet2 then
				outputString=outputString+tensTable(triplet1)+" "+unitsTable(triplet2)
			else
				outputString=outputString+tensTable(triplet1)
			endif
		end select
	endif
return
end function
