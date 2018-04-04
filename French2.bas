function french(inputvalue as double) as string
	dim inputstring as string
	dim ouputstring as string
	dim negative as boolean
	dim leng as integer
	dim unit01 as variant
	dim unit11 as variant
	dim tens as variant
	dim multiple as variant
	dim triplet as integer
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	dim i as integer
	unit01=array(""," UN"," DEUX"," TROIS"," QUATRE"," CINQ"," SIX"," SEPT"," HUIT"," NEUF")
	unit11=array(" DIX"," ONZE"," DOUZE"," TREIZE"," QUATORZE"," QUINZE"," SEIZE"," DIX SEPT"," DIX HUIT"," DIX NEUF")
	tens=array("",""," VINGT"," TRENTE"," QUARANTE"," CINQUANTE"," SOIXANTE"," SOIXANTE DIX"," QUATRE VINGT"," QUATRE VINGT DIX")
	if inputvalue>999999999999 then
		french=""
		exit function
	endif
	if inputvalue=0 then
		outputstring="ZÉRO"
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
		select case triplet
		case 0:
		case 1
			select case i
			case 1: outputstring=outputstring+" UN MILLIARD"
			case 4: outputstring=outputstring+" UN MILLION"
			case 7: outputstring=outputstring+" MILLE"
			end select
		case else
			triplet1=triplet\100
			triplet2=(triplet-(triplet1*100))\10
			triplet3=triplet-triplet1*100-triplet2*10
			select case triplet1
			case 0:
			case 1: outputstring=outputstring+" CENT"
			case else: outputstring=outputstring+unit01(triplet1)+" CENT"
			end select
			select case triplet2
			case 0: if triplet3 then outputstring=outputstring+unit01(triplet3)
			case 1: outputstring=outputstring+unit11(triplet3)
			case 2,3,4,5,6: if triplet3=1 then outputstring=outputstring+tens(triplet2)+" ET UN" else outputstring=outputstring+tens(triplet2)+unit01(triplet3)
			case 7: if triplet3=1 then outputstring=outputstring+" SOIXANTE ET ONZE" else outputstring=outputstring+" SOIXANTE"+unit11(triplet3)
			case 8: if triplet3 then outputstring=outputstring+" QUATRE VINGT"+unit01(triplet3) else outputstring=outputstring+" QUATRE VINGT"
			case 9: if triplet3 then outputstring=outputstring+" QUATRE VINGT"+unit11(triplet3) else outputstring=outputstring+" QUATRE VINGT DIX"
			end select
			select case i
			case 1: outputstring=outputstring+" MILLIARDS"
			case 4: outputstring=outputstring+" MILLIONS"
			case 7: outputstring=outputstring+" MILLE"
			end select
		end select
	next
	if negative then outputstring="MOINS"+outputstring else	outputstring=mid(outputstring,2,)
	french=outputstring
end function
