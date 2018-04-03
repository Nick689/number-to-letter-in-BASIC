function english2(inputvalue as double) as string
	dim inputstring as string
	dim outputstring as string
	dim negative as boolean
	dim leng as integer
	dim unit01 as variant
	dim unit11 as variant
	dim tens as variant
	dim triplet as integer
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	dim i as integer
	unit01=array(""," ONE"," TWO"," THREE"," FOUR"," FIVE"," SIX"," SEVEN"," EIGHT"," NINE")
	unit11=array(" TEN"," ELEVEN"," TWELVE"," THIRTEEN"," FOURTEEN"," FIFTEEN"," SIXTEEN"," SEVENTEEN"," EIGHTEEN"," NINETEEN")
	tens=array("",""," TWENTY"," THIRTY"," FORTY"," FIFTY"," SIXTY"," SEVENTY"," EIGHTY"," NINETY")
	if inputvalue>999999999999 then
		english2=""
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
		triplet=cint(mid(inputstring,i,3))
		if triplet then
			triplet1=triplet\100
			triplet2=(triplet-(triplet1*100))\10
			triplet3=triplet-triplet1*100-triplet2*10
			if triplet then
				if triplet1 then outputstring=outputstring+unit01(triplet1)+" HUNDRED"
				if triplet2=1 then outputstring=outputstring+unit11(triplet3) else outputstring=outputstring+tens(triplet2)+unit01(triplet3)
			endif
			select case i
			case 1: outputstring=outputstring+" BILLION"
			case 4: outputstring=outputstring+" MILLION"
			case 7: outputstring=outputstring+" THOUSAND"
			end select
		endif
	next
	if negative then outputstring="MINUS"+outputstring else outputstring=mid(outputstring,2,)
	english2=outputstring
end function
