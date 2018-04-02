function english3(inputvalue as double) as string
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
	unit01=array(""," ONE"," TWO"," THREE"," FOUR"," FIVE"," SIX"," SEVEN"," EIGHT"," NINE")
	unit11=array(" TEN"," ELEVEN"," TWELVE"," THIRTEEN"," FOURTEEN"," FIFTEEN"," SIXTEEN"," SEVENTEEN"," EIGHTEEN"," NINETEEN")
	if inputvalue>999999999999 then
		english3=""
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
	triplet=cint(mid(inputstring,1,3))
	if triplet then
		triplet1=triplet\100
		triplet2=(triplet-(triplet1*100))\10
		triplet3=triplet-triplet1*100-triplet2*10			
		if triplet1 then outputstring=outputstring+unit01(triplet1)+" HUNDRED"
		select case triplet2
		case 0: outputstring=outputstring+unit01(triplet3)+" BILLION"
		case 1: outputstring=outputstring+unit11(triplet3)+" BILLION"
		case 2: outputstring=outputstring+" TWENTY"+unit01(triplet3)+" BILLION"
		case 3: outputstring=outputstring+" THIRTY"+unit01(triplet3)+" BILLION"
		case 4: outputstring=outputstring+" FORTY"+unit01(triplet3)+" BILLION"
		case 5: outputstring=outputstring+" FIFTY"+unit01(triplet3)+" BILLION"
		case 6: outputstring=outputstring+" SIXTY"+unit01(triplet3)+" BILLION"
		case 7: outputstring=outputstring+" SEVENTY"+unit01(triplet3)+" BILLION"
		case 8: outputstring=outputstring+" EIGHTY"+unit01(triplet3)+" BILLION"
		case 9: outputstring=outputstring+" NINETY"+unit01(triplet3)+" BILLION"
		end select
	endif
	triplet=cint(mid(inputstring,4,3))
	if triplet then
		triplet1=triplet\100
		triplet2=(triplet-(triplet1*100))\10
		triplet3=triplet-triplet1*100-triplet2*10			
		if triplet1 then outputstring=outputstring+unit01(triplet1)+" HUNDRED"
		select case triplet2
		case 0: outputstring=outputstring+unit01(triplet3)+" MILLION"
		case 1: outputstring=outputstring+unit11(triplet3)+" MILLION"
		case 2: outputstring=outputstring+" TWENTY"+unit01(triplet3)+" MILLION"
		case 3: outputstring=outputstring+" THIRTY"+unit01(triplet3)+" MILLION"
		case 4: outputstring=outputstring+" FORTY"+unit01(triplet3)+" MILLION"
		case 5: outputstring=outputstring+" FIFTY"+unit01(triplet3)+" MILLION"
		case 6: outputstring=outputstring+" SIXTY"+unit01(triplet3)+" MILLION"
		case 7: outputstring=outputstring+" SEVENTY"+unit01(triplet3)+" MILLION"
		case 8: outputstring=outputstring+" EIGHTY"+unit01(triplet3)+" MILLION"
		case 9: outputstring=outputstring+" NINETY"+unit01(triplet3)+" MILLION"
		end select
	endif
	triplet=cint(mid(inputstring,7,3))
	if triplet then
		triplet1=triplet\100
		triplet2=(triplet-(triplet1*100))\10
		triplet3=triplet-triplet1*100-triplet2*10			
		if triplet1 then outputstring=outputstring+unit01(triplet1)+" HUNDRED"
		select case triplet2
		case 0: outputstring=outputstring+unit01(triplet3)+" THOUSAND"
		case 1: outputstring=outputstring+unit11(triplet3)+" THOUSAND"
		case 2: outputstring=outputstring+" TWENTY"+unit01(triplet3)+" THOUSAND"
		case 3: outputstring=outputstring+" THIRTY"+unit01(triplet3)+" THOUSAND"
		case 4: outputstring=outputstring+" FORTY"+unit01(triplet3)+" THOUSAND"
		case 5: outputstring=outputstring+" FIFTY"+unit01(triplet3)+" THOUSAND"
		case 6: outputstring=outputstring+" SIXTY"+unit01(triplet3)+" THOUSAND"
		case 7: outputstring=outputstring+" SEVENTY"+unit01(triplet3)+" THOUSAND"
		case 8: outputstring=outputstring+" EIGHTY"+unit01(triplet3)+" THOUSAND"
		case 9: outputstring=outputstring+" NINETY"+unit01(triplet3)+" THOUSAND"
		end select
	endif
	triplet=cint(mid(inputstring,10,3))
	if triplet then
		triplet1=triplet\100
		triplet2=(triplet-(triplet1*100))\10
		triplet3=triplet-triplet1*100-triplet2*10			
		if triplet1 then outputstring=outputstring+unit01(triplet1)+" HUNDRED"
		select case triplet2
		case 0: outputstring=outputstring+unit01(triplet3)
		case 1: outputstring=outputstring+unit11(triplet3)
		case 2: outputstring=outputstring+" TWENTY"+unit01(triplet3)
		case 3: outputstring=outputstring+" THIRTY"+unit01(triplet3)
		case 4: outputstring=outputstring+" FORTY"+unit01(triplet3)
		case 5: outputstring=outputstring+" FIFTY"+unit01(triplet3)
		case 6: outputstring=outputstring+" SIXTY"+unit01(triplet3)
		case 7: outputstring=outputstring+" SEVENTY"+unit01(triplet3)
		case 8: outputstring=outputstring+" EIGHTY"+unit01(triplet3)
		case 9: outputstring=outputstring+" NINETY"+unit01(triplet3)
		end select
	endif
	if negative then
		outputstring="MINUS"+outputstring
	else
		outputstring=mid(outputstring,2,)
	endif
	english3=outputstring
end function
