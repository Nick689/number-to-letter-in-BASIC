function english2(inputvalue as double) as string
	dim inputstring as string
	dim ouputstring as string
	dim negative as boolean
	dim leng as integer
	dim triplet as integer
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	dim i as integer
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
			select case triplet1
			case 1: outputstring=outputstring+" HUNDRED"
			case 2: outputstring=outputstring+" TWO HUNDRED"
			case 3: outputstring=outputstring+" THREE HUNDRED"
			case 4: outputstring=outputstring+" FOUR HUNDRED"
			case 5: outputstring=outputstring+" FIVE HUNDRED"
			case 6: outputstring=outputstring+" SIX HUNDRED"
			case 7: outputstring=outputstring+" SEVEN HUNDRED"
			case 8: outputstring=outputstring+" EIGHT HUNDRED"
			case 9: outputstring=outputstring+" NINE HUNDRED"
			end select
			select case triplet2
			case 0: gosub sub01
			case 1: gosub sub11
			case 2
				outputstring=outputstring+" TWENTY"
				if triplet3 then gosub sub01
			case 3
				outputstring=outputstring+" THIRTY"
				if triplet3 then gosub sub01
			case 4
				outputstring=outputstring+" FORTY"
				if triplet3 then gosub sub01
			case 5
				outputstring=outputstring+" FIFTY"
				if triplet3 then gosub sub01
			case 6
				outputstring=outputstring+" SIXTY"
				if triplet3 then gosub sub01
			case 7
				outputstring=outputstring+" SEVENTY"
				if triplet3 then gosub sub01
			case 8
				outputstring=outputstring+" EIGHTY"
				if triplet3 then gosub sub01
			case 9
				outputstring=outputstring+" NINETY"
				if triplet3 then gosub sub01
			end select
			select case i
			case 1: outputstring=outputstring+" BILLION"
			case 4: outputstring=outputstring+" MILLION"
			case 7: outputstring=outputstring+" THOUSAND"
			end select
		endif
	next
	if negative then outputstring="MINUS"+outputstring else	outputstring=mid(outputstring,2,)
	english2=outputstring
	exit function
sub01:
	select case triplet3
	case 1: outputstring=outputstring+" ONE"
	case 2: outputstring=outputstring+" TWO"
	case 3: outputstring=outputstring+" THREE"
	case 4: outputstring=outputstring+" FOUR"
	case 5: outputstring=outputstring+" FIVE"
	case 6: outputstring=outputstring+" SIX"
	case 7: outputstring=outputstring+" SEVEN"
	case 8: outputstring=outputstring+" EIGHT"
	case 9: outputstring=outputstring+" NINE"
	end select
return
sub11:
	select case triplet3
	case 0: outputstring=outputstring+" TEN"
	case 1: outputstring=outputstring+" ELEVEN"
	case 2: outputstring=outputstring+" TWELVE"
	case 3: outputstring=outputstring+" THIRTEEN"
	case 4: outputstring=outputstring+" FOURTEEN"
	case 5: outputstring=outputstring+" FIFTEEN"
	case 6: outputstring=outputstring+" SIXTEEN"
	case 7: outputstring=outputstring+" SEVENTEEN"
	case 8: outputstring=outputstring+" EIGHTEEN"
	case 9: outputstring=outputstring+" NINETEEN"
	end select
return
end function
