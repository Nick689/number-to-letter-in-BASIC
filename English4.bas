function numtoletter(numValue as double) as string
	dim numString as string
	dim numTable(12) as integer
	dim Table00(10) as string
	dim Table10(10) as string
	dim Table20(10) as string
	dim Table30(10) as string
	dim Table40(10) as string
	dim Table50(10) as string
	dim Table60(10) as string
	dim Table70(10) as string
	dim Table80(10) as string
	dim Table90(10) as string
	dim negative as boolean
	dim leng as integer
	dim outputString as string
	dim i as integer
	table00=array("ZERO","ONE","TWO","THREE","FOUR","FIVE","SIX","SEVEN","EIGHT","NINE")
	table10=array("TEN","ELEVEN","TWELVE","THIRTEEN","FOURTEEN","FIFTEEN","SIXTEEN","SEVENTEEN","EIGHTEEN","NINETEEN")
	table20=array("TWENTY ","TWENTY ONE","TWENTY TWO","TWENTY THREE","TWENTY FOUR","TWENTY FIVE","TWENTY SIX","TWENTY SEVEN","TWENTY EIGHT","TWENTY NINE")
	table30=array("THIRTY","THIRTY ONE","THIRTY TWO","THIRTY THREE","THIRTY FOUR","THIRTY FIVE","THIRTY SIX","THIRTY SEVEN","THIRTY EIGHT","THIRTY NINE")
	table40=array("FORTY","FORTY ONE","FORTY TWO","FORTY THREE","FORTY FOUR","FORTY FIVE","FORTY SIX","FORTY SEVEN","FORTY EIGHT","FORTY NINE")
	table50=array("FIFTY","FIFTY ONE","FIFTY TWO","FIFTY THREE","FIFTY FOUR","FIFTY FIVE","FIFTY SIX","FIFTY SEVEN","FIFTY EIGHT","FIFTY NINE")
	table60=array("SIXTY","SIXTY ONE","SIXTY TWO","SIXTY THREE","SIXTY FOUR","SIXTY FIVE","SIXTY SIX","SIXTY SEVEN","SIXTY EIGHT","SIXTY NINE")
	table70=array("SEVENTY","SEVENTY ONE","SEVENTY TWO","SEVENTY THREE","SEVENTY FOUR","SEVENTY FIVE","SEVENTY SIX","SEVENTY SEVEN","SEVENTY EIGHT","SEVENTY NINE")
	table80=array("EIGHTY","EIGHTY ONE","EIGHTY TWO","EIGHTY THREE","EIGHTY FOUR","EIGHTY FIVE","EIGHTY SIX","EIGHTY SEVEN","EIGHTY EIGHT","EIGHTY NINE")
	table90=array("NINETY","NINETY ONE","NINETY TWO","NINETY THREE","NINETY FOUR","NINETY FIVE","NINETY SIX","NINETY SEVEN","NINETY EIGHT","NINETY NINE")
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
			case 0: outputString=table00(numTable(0))+" HUNDRED "+table00(numTable(2))+" BILLION"
			case 1: outputString=table00(numTable(0))+" HUNDRED "+table10(numTable(2))+" BILLION"
			case 2: outputString=table00(numTable(0))+" HUNDRED "+table20(numTable(2))+" BILLION"
			case 3: outputString=table00(numTable(0))+" HUNDRED "+table30(numTable(2))+" BILLION"
			case 4: outputString=table00(numTable(0))+" HUNDRED "+table40(numTable(2))+" BILLION"
			case 5: outputString=table00(numTable(0))+" HUNDRED "+table50(numTable(2))+" BILLION"
			case 6: outputString=table00(numTable(0))+" HUNDRED "+table60(numTable(2))+" BILLION"
			case 7: outputString=table00(numTable(0))+" HUNDRED "+table70(numTable(2))+" BILLION"
			case 8: outputString=table00(numTable(0))+" HUNDRED "+table80(numTable(2))+" BILLION"
			case 9: outputString=table00(numTable(0))+" HUNDRED "+table90(numTable(2))+" BILLION"
			end select
		else
			select case numTable(1)
			case 0: outputString=table00(numTable(2))+" BILLION"
			case 1: outputString=table10(numTable(2))+" BILLION"
			case 2: outputString=table20(numTable(2))+" BILLION"
			case 3: outputString=table30(numTable(2))+" BILLION"
			case 4: outputString=table40(numTable(2))+" BILLION"
			case 5: outputString=table50(numTable(2))+" BILLION"
			case 6: outputString=table60(numTable(2))+" BILLION"
			case 7: outputString=table70(numTable(2))+" BILLION"
			case 8: outputString=table80(numTable(2))+" BILLION"
			case 9: outputString=table90(numTable(2))+" BILLION"
			end select
		endif
	endif
	if numTable(3)+numTable(4)+numTable(5) then
		if outputString<>"" then outputString=outputString+" "
		if numTable(3) then
			select case numTable(4)
			case 0: outputString=outputString+table00(numTable(3))+" HUNDRED "+table00(numTable(5))+" MILLION"
			case 1: outputString=outputString+table00(numTable(3))+" HUNDRED "+table10(numTable(5))+" MILLION"
			case 2: outputString=outputString+table00(numTable(3))+" HUNDRED "+table20(numTable(5))+" MILLION"
			case 3: outputString=outputString+table00(numTable(3))+" HUNDRED "+table30(numTable(5))+" MILLION"
			case 4: outputString=outputString+table00(numTable(3))+" HUNDRED "+table40(numTable(5))+" MILLION"
			case 5: outputString=outputString+table00(numTable(3))+" HUNDRED "+table50(numTable(5))+" MILLION"
			case 6: outputString=outputString+table00(numTable(3))+" HUNDRED "+table60(numTable(5))+" MILLION"
			case 7: outputString=outputString+table00(numTable(3))+" HUNDRED "+table70(numTable(5))+" MILLION"
			case 8: outputString=outputString+table00(numTable(3))+" HUNDRED "+table80(numTable(5))+" MILLION"
			case 9: outputString=outputString+table00(numTable(3))+" HUNDRED "+table90(numTable(5))+" MILLION"
			end select
		else
			select case numTable(4)
			case 0: outputString=outputString+table00(numTable(5))+" MILLION"
			case 1: outputString=outputString+table10(numTable(5))+" MILLION"
			case 2: outputString=outputString+table20(numTable(5))+" MILLION"
			case 3: outputString=outputString+table30(numTable(5))+" MILLION"
			case 4: outputString=outputString+table40(numTable(5))+" MILLION"
			case 5: outputString=outputString+table50(numTable(5))+" MILLION"
			case 6: outputString=outputString+table60(numTable(5))+" MILLION"
			case 7: outputString=outputString+table70(numTable(5))+" MILLION"
			case 8: outputString=outputString+table80(numTable(5))+" MILLION"
			case 9: outputString=outputString+table90(numTable(5))+" MILLION"
			end select
		endif
	endif
	if numTable(6)+numTable(7)+numTable(8) then
		if outputString<>"" then outputString=outputString+" "
		if numTable(6) then
			select case numTable(7)
			case 0: outputString=outputString+table00(numTable(6))+" HUNDRED "+table00(numTable(8))+" THOUSAND"
			case 1: outputString=outputString+table00(numTable(6))+" HUNDRED "+table10(numTable(8))+" THOUSAND"
			case 2: outputString=outputString+table00(numTable(6))+" HUNDRED "+table20(numTable(8))+" THOUSAND"
			case 3: outputString=outputString+table00(numTable(6))+" HUNDRED "+table30(numTable(8))+" THOUSAND"
			case 4: outputString=outputString+table00(numTable(6))+" HUNDRED "+table40(numTable(8))+" THOUSAND"
			case 5: outputString=outputString+table00(numTable(6))+" HUNDRED "+table50(numTable(8))+" THOUSAND"
			case 6: outputString=outputString+table00(numTable(6))+" HUNDRED "+table60(numTable(8))+" THOUSAND"
			case 7: outputString=outputString+table00(numTable(6))+" HUNDRED "+table70(numTable(8))+" THOUSAND"
			case 8: outputString=outputString+table00(numTable(6))+" HUNDRED "+table80(numTable(8))+" THOUSAND"
			case 9: outputString=outputString+table00(numTable(6))+" HUNDRED "+table90(numTable(8))+" THOUSAND"
			end select
		else
			select case numTable(7)
			case 0: outputString=outputString+table00(numTable(8))+" THOUSAND"
			case 1: outputString=outputString+table10(numTable(8))+" THOUSAND"
			case 2: outputString=outputString+table20(numTable(8))+" THOUSAND"
			case 3: outputString=outputString+table30(numTable(8))+" THOUSAND"
			case 4: outputString=outputString+table40(numTable(8))+" THOUSAND"
			case 5: outputString=outputString+table50(numTable(8))+" THOUSAND"
			case 6: outputString=outputString+table60(numTable(8))+" THOUSAND"
			case 7: outputString=outputString+table70(numTable(8))+" THOUSAND"
			case 8: outputString=outputString+table80(numTable(8))+" THOUSAND"
			case 9: outputString=outputString+table90(numTable(8))+" THOUSAND"
			end select
		endif
	endif
	if numTable(9)+numTable(10)+numTable(11) then
		if outputString<>"" then outputString=outputString+" "
		if numTable(9) then
			select case numTable(10)
			case 0: outputString=outputString+table00(numTable(9))+" HUNDRED "+table00(numTable(11))
			case 1: outputString=outputString+table00(numTable(9))+" HUNDRED "+table10(numTable(11))
			case 2: outputString=outputString+table00(numTable(9))+" HUNDRED "+table20(numTable(11))
			case 3: outputString=outputString+table00(numTable(9))+" HUNDRED "+table30(numTable(11))
			case 4: outputString=outputString+table00(numTable(9))+" HUNDRED "+table40(numTable(11))
			case 5: outputString=outputString+table00(numTable(9))+" HUNDRED "+table50(numTable(11))
			case 6: outputString=outputString+table00(numTable(9))+" HUNDRED "+table60(numTable(11))
			case 7: outputString=outputString+table00(numTable(9))+" HUNDRED "+table70(numTable(11))
			case 8: outputString=outputString+table00(numTable(9))+" HUNDRED "+table80(numTable(11))
			case 9: outputString=outputString+table00(numTable(9))+" HUNDRED "+table90(numTable(11))
			end select
		else
			select case numTable(10)
			case 0: outputString=outputString+table00(numTable(11))
			case 1: outputString=outputString+table10(numTable(11))
			case 2: outputString=outputString+table20(numTable(11))
			case 3: outputString=outputString+table30(numTable(11))
			case 4: outputString=outputString+table40(numTable(11))
			case 5: outputString=outputString+table50(numTable(11))
			case 6: outputString=outputString+table60(numTable(11))
			case 7: outputString=outputString+table70(numTable(11))
			case 8: outputString=outputString+table80(numTable(11))
			case 9: outputString=outputString+table90(numTable(11))
			end select
		endif
	endif
	if numValue=0 then outputString="ZERO"
	if negative then outputString="MINUS "+outputString
	numtoletter=outputString
end function
