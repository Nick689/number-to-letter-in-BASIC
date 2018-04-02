function english1(inputvalue as double) as string
	dim inputstring as string
	dim outputstring as string
	dim negative as boolean
	dim leng as integer
	dim unit01 as variant
	dim unit11 as variant
	dim tens as variant
	dim multiple as variant
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	dim i as integer
	unit01=array(""," ONE"," TWO"," THREE"," FOUR"," FIVE"," SIX"," SEVEN"," EIGHT"," NINE")
	unit11=array(" TEN"," ELEVEN"," TWELVE"," THIRTEEN"," FOURTEEN"," FIFTEEN"," SIXTEEN"," SEVENTEEN"," EIGHTEEN"," NINETEEN")
	tens=array("",""," TWENTY"," THIRTY"," FORTY"," FIFTY"," SIXTY"," SEVENTY"," EIGHTY"," NINETY")
	multiple=array(""," BILLION","",""," MILLION","",""," THOUSAND","","","")
	if inputvalue>999999999999 then
		english1=""
		exit function
	endif
	if inputvalue=0 then
		outputstring="ZERO"
		exit function
	endif
	if inputvalue<0 then
		negative=true
		inputvalue=abs(inputvalue)
	else
		negative=false
	endif
	inputstring=cstr(inputvalue)
	leng=len(inputstring)
	while leng<12
		inputstring="0"+inputstring
		leng=leng+1
	wend
	for i=1 to 10 step 3
		triplet1=cint(mid(inputstring,i,1))
		triplet2=cint(mid(inputstring,i+1,1))
		triplet3=cint(mid(inputstring,i+2,1))
		if triplet1+triplet2+triplet3 then
			if triplet1 then outputstring=outputstring+unit01(triplet1)+" HUNDRED"
			select case triplet2
			case 0: outputstring=outputstring+unit01(triplet3)+multiple(i)
			case 1: outputstring=outputstring+unit11(triplet3)+multiple(i)
			case else: outputstring=outputstring+tens(triplet2)+unit01(triplet3)+multiple(i)
			end select
		endif
	next
	if negative then
		outputstring="MINUS"+outputstring
	else
		outputstring=mid(outputstring,2,)
	endif
	english1=outputstring
end function
