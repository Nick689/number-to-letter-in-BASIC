function french(inputvalue as double) as string
	dim inputstring as string
	dim ouputstring as string
	dim negative as boolean
	dim leng as integer
	dim units as variant
	dim tens as variant
	dim multiple as variant
	dim triplet as integer
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	dim i as integer
	units=array(""," UN"," DEUX"," TROIS"," QUATRE"," CINQ"," SIX"," SEPT"," HUIT"," NEUF"," DIX"," ONZE"," DOUZE"," TREIZE"," QUATORZE"," QUINZE"," SEIZE"," DIX SEPT"," DIX HUIT"," DIX NEUF")
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
		if triplet then
			triplet1=triplet\100
			triplet2=(triplet-(triplet1*100))\10
			triplet3=triplet-triplet1*100-triplet2*10
			select case triplet1
				case 0:
				case 1: outputstring=outputstring+" CENT"
				case else: outputstring=outputstring+units(triplet1)+" CENT"
			end select
			select case triplet2
				case 0: if triplet3>1 then outputstring=outputstring+units(triplet3)
				case 1: outputstring=outputstring+units(10+triplet3)
				case 2
					select case triplet3
						case 0: outputstring=outputstring+" VINGT"
						case 1: outputstring=outputstring+" VINGT ET"+units(triplet3)
						case else: outputstring=outputstring+" VINGT"+units(triplet3)
					end select
				case 3
					select case triplet3
						case 0: outputstring=outputstring+" TRENTE"
						case 1: outputstring=outputstring+" TRENTE ET"+units(triplet3)
						case else: outputstring=outputstring+" TRENTE"+units(triplet3)
					end select
				case 4
					select case triplet3
						case 0: outputstring=outputstring+" QUARANTE"
						case 1: outputstring=outputstring+" QUARANTE ET"+units(triplet3)
						case else: outputstring=outputstring+" QUARANTE"+units(triplet3)
					end select
				case 5
					select case triplet3
						case 0: outputstring=outputstring+" CINQUANTE"
						case 1: outputstring=outputstring+" CINQUANTE ET"+units(triplet3)
						case else: outputstring=outputstring+" CINQUANTE"+units(triplet3)
					end select
				case 6
					select case triplet3
						case 0: outputstring=outputstring+" SOIXANTE"
						case 1: outputstring=outputstring+" SOIXANTE ET"+units(triplet3)
						case else: outputstring=outputstring+" SOIXANTE"+units(triplet3)
					end select
				case 7
					select case triplet3
						case 0: outputstring=outputstring+" SOIXANTE DIX"
						case 1: outputstring=outputstring+" SOIXANTE ET ONZE"
						case else: outputstring=outputstring+" SOIXANTE"+units(10+triplet3)
					end select
				case 8
					if triplet3 then
						outputstring=outputstring+" QUATRE VINGT"+units(triplet3)
					else
						outputstring=outputstring+" QUATRE VINGT"
					end if
				case 9
					if triplet3 then
						outputstring=outputstring+" QUATRE VINGT"+units(10+triplet3)
					else
						outputstring=outputstring+" QUATRE VINGT DIX"
					end if
			end select
			select case i
				case 1:
				if triplet=1 then
					outputstring=outputstring+" MILLIARD"
				else
					outputstring=outputstring+" MILLIARDS"
				endif
				case 4:
				if triplet=1 then
					outputstring=outputstring+" MILLION"
				else
					outputstring=outputstring+" MILLIONS"
				endif
				case 7: outputstring=outputstring+" MILLE"
			end select
		endif
	next
	if negative then
		outputstring="MOINS"+outputstring
	else
		outputstring=mid(outputstring,2,)
	endif
	french=outputstring
end function
