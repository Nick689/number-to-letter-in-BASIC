function numtoletter(numValue as double) as string
	dim numString as string
	dim num01 as integer
	dim num02 as integer
	dim num03 as integer
	dim num04 as integer
	dim num05 as integer
	dim num06 as integer
	dim num07 as integer
	dim num08 as integer
	dim num09 as integer
	dim num10 as integer
	dim num11 as integer
	dim num12 as integer
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
	num01=cint(mid(numString,1,1))
	num02=cint(mid(numString,2,1))
	num03=cint(mid(numString,3,1))
	num04=cint(mid(numString,4,1))
	num05=cint(mid(numString,5,1))
	num06=cint(mid(numString,6,1))
	num07=cint(mid(numString,7,1))
	num08=cint(mid(numString,8,1))
	num09=cint(mid(numString,9,1))
	num10=cint(mid(numString,10,1))
	num11=cint(mid(numString,11,1))
	num12=cint(mid(numString,12,1))
	if num01+num02+num03 then
		if num01 then
			select case num02
			case 0: outputString=unitsTable(num01)+" HUNDRED "+unitsTable(num03)+" BILLION"
			case 1: outputString=unitsTable(num01)+" HUNDRED "+unitsTable(10+num03)+" BILLION"
			case else: outputString=unitsTable(num01)+" HUNDRED "+tensTable(num02)+" "+unitsTable(num03)+" BILLION"
			end select
		else
			select case num02
			case 0: outputString=unitsTable(num03)+" BILLION"
			case 1: outputString=unitsTable(10+num03)+" BILLION"
			case else: outputString=tensTable(num02)+" "+unitsTable(num03)+" BILLION"
			end select
		endif
	endif
	if num04+num05+num06 then
		if outputString<>"" then outputString=outputString+" "
		if num04 then
			select case num05
			case 0:	outputString=outputString+unitsTable(num04)+" HUNDRED "+unitsTable(num06)+" MILLION"
			case 1: outputString=outputString+unitsTable(num04)+" HUNDRED "+unitsTable(10+num06)+" MILLION"
			case else: outputString=outputString+unitsTable(num04)+" HUNDRED "+tensTable(num05)+" "+unitsTable(num06)+" MILLION"
			end select
		else
			select case num05
			case 0: outputString=outputString+unitsTable(num06)+" MILLION"
			case 1: outputString=outputString+unitsTable(10+num06)+" MILLION"
			case else: outputString=outputString+tensTable(num05)+" "+unitsTable(num06)+" MILLION"
			end select
		endif
	endif
	if num07+num08+num09 then
		if outputString<>"" then outputString=outputString+" "
		if num07 then
			select case num08
			case 0: outputString=outputString+unitsTable(num07)+" HUNDRED "+unitsTable(num09)+" THOUSAND"
			case 1: outputString=outputString+unitsTable(num07)+" HUNDRED "+unitsTable(10+num09)+" THOUSAND"
			case else: outputString=outputString+unitsTable(num07)+" HUNDRED "+tensTable(num08)+" "+unitsTable(num09)+" THOUSAND"
			end select
		else
			select case num08
			case 0: outputString=outputString+unitsTable(num09)+" THOUSAND"
			case 1: outputString=outputString+unitsTable(10+num09)+" THOUSAND"
			case else: outputString=outputString+tensTable(num08)+" "+unitsTable(num09)+" THOUSAND"
			end select
		endif
	endif
	if num10+num11+num12 then
		if outputString<>"" then outputString=outputString+" "
		if num10 then
			select case num11
			case 0: outputString=outputString+unitsTable(num10)+" HUNDRED "+unitsTable(num12)
			case 1: outputString=outputString+unitsTable(num10)+" HUNDRED "+unitsTable(10+num12)
			case else: outputString=outputString+unitsTable(num10)+" HUNDRED "+tensTable(num11)+" "+unitsTable(num12)
			end select
		else
			select case num11
			case 0: outputString=outputString+unitsTable(num12)
			case 1: outputString=outputString+unitsTable(10+num12)
			case else: outputString=outputString+tensTable(num11)+" "+unitsTable(num12)
			end select
		endif
	endif
	if numValue=0 then outputString="ZERO"
	if negative then outputString="MINUS "+outputString
	numtoletter=outputString
end function
