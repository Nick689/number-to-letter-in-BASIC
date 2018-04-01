function numtoletter(numValue as double) as string' 189s
	dim numString as string
	dim units as variant
	dim tens as variant
	dim multiple as variant
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	dim negative as boolean
	dim leng as integer
	dim outputString as string
	dim i as integer
	units=array(""," ONE"," TWO"," THREE"," FOUR"," FIVE"," SIX"," SEVEN"," EIGHT"," NINE"," TEN"," ELEVEN"," TWELVE"," THIRTEEN"," FOURTEEN"," FIFTEEN"," SIXTEEN"," SEVENTEEN"," EIGHTEEN"," NINETEEN")
	tens=array("",""," TWENTY"," THIRTY"," FORTY"," FIFTY"," SIXTY"," SEVENTY"," EIGHTY"," NINETY")
	multiple=array(""," BILLION","",""," MILLION","",""," THOUSAND","","","")
	if numValue>999999999999 then
		numtoletter=""
		exit function
	endif
	if numValue=0 then
		outputString="ZERO"
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
	for i=1 to 10 step 3
		triplet1=cint(mid(numString,i,1))
		triplet2=cint(mid(numString,i+1,1))
		triplet3=cint(mid(numString,i+2,1))
		if triplet1+triplet2+triplet3 then
			if triplet1 then outputString=outputString+units(triplet1)+" HUNDRED"
			select case triplet2
			case 0: outputString=outputString+units(triplet3)+multiple(i)
			case 1: outputString=outputString+units(10+triplet3)+multiple(i)
			case else: outputString=outputString+tens(triplet2)+units(triplet3)+multiple(i)
			end select
		endif
	next
	if negative then
		outputString="MINUS"+outputString
	else
		outputString=right(outputString,len(outputString)-1)
	endif
	numtoletter=outputString
end function
