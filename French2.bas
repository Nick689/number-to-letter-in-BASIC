function french2(inputvalue as double) as string
	dim inputstring as string
	dim ouputstring as string
	dim negative as boolean
	dim leng as integer
	dim triplet as integer
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	if inputvalue>999999999999 then
		french2=""
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
	triplet=cint(mid(inputstring,1,3))
	select case triplet
	case 0:
	case 1: outputstring=outputstring+" UN MILLIARD"
	case else
		gosub sub00
		outputstring=outputstring+" MILLIARDS"
	end select
	triplet=cint(mid(inputstring,4,3))
	select case triplet
	case 0:
	case 1: outputstring=outputstring+" UN MILLION"
	case else
		gosub sub00
		outputstring=outputstring+" MILLIONS"
	end select
	triplet=cint(mid(inputstring,7,3))
	select case triplet
	case 0:
	case 1: outputstring=outputstring+" MILLE"
	case else
		gosub sub00
		outputstring=outputstring+" MILLE"
	end select
	triplet=cint(mid(inputstring,10,3))
	if triplet then gosub sub00
	if negative then outputstring="MOINS"+outputstring else	outputstring=mid(outputstring,2,)
	french2=outputstring
	exit function
sub00:
	triplet1=triplet\100
	triplet2=(triplet-(triplet1*100))\10
	triplet3=triplet-triplet1*100-triplet2*10
	select case triplet1
	case 1: outputstring=outputstring+" CENT"
	case 2: outputstring=outputstring+" DEUX CENT"
	case 3: outputstring=outputstring+" TROIS CENT"
	case 4: outputstring=outputstring+" QUATRE CENT"
	case 5: outputstring=outputstring+" CINQ CENT"
	case 6: outputstring=outputstring+" SIX CENT"
	case 7: outputstring=outputstring+" SEPT CENT"
	case 8: outputstring=outputstring+" HUIT CENT"
	case 9: outputstring=outputstring+" NEUF CENT"
	end select
	select case triplet2
	case 0: gosub sub01
	case 1: gosub sub11
	case 2
		if triplet3=1 then outputstring=outputstring+" VINGT ET" else outputstring=outputstring+" VINGT"
		gosub sub01
	case 3
		if triplet3=1 then outputstring=outputstring+" TRENTE ET" else outputstring=outputstring+" TRENTE"
		gosub sub01
	case 4
		if triplet3=1 then outputstring=outputstring+" QUARANTE ET" else outputstring=outputstring+" QUARANTE"
		gosub sub01
	case 5
		if triplet3=1 then outputstring=outputstring+" CINQUANTE ET" else outputstring=outputstring+" CINQUANTE"
		gosub sub01
	case 6
		if triplet3=1 then outputstring=outputstring+" SOIXANTE ET" else outputstring=outputstring+" SOIXANTE"
		gosub sub01
	case 7
		if triplet3=1 then outputstring=outputstring+" SOIXANTE ET" else outputstring=outputstring+" SOIXANTE"
		gosub sub11
	case 8
		outputstring=outputstring+" QUATRE VINGT"
		gosub sub01
	case 9
		outputstring=outputstring+" QUATRE VINGT"
		gosub sub11
	end select
return
sub01:
	select case triplet3
	case 1: outputstring=outputstring+" UN"
	case 2: outputstring=outputstring+" DEUX"
	case 3: outputstring=outputstring+" TROIS"
	case 4: outputstring=outputstring+" QUATRE"
	case 5: outputstring=outputstring+" CINQ"
	case 6: outputstring=outputstring+" SIX"
	case 7: outputstring=outputstring+" SEPT"
	case 8: outputstring=outputstring+" HUIT"
	case 9: outputstring=outputstring+" NEUF"
	end select
return
sub11:
	select case triplet3
	case 0: outputstring=outputstring+" DIX"
	case 1: outputstring=outputstring+" ONZE"
	case 2: outputstring=outputstring+" DOUZE"
	case 3: outputstring=outputstring+" TREIZE"
	case 4: outputstring=outputstring+" QUATORZE"
	case 5: outputstring=outputstring+" QUINZE"
	case 6: outputstring=outputstring+" SEIZE"
	case 7: outputstring=outputstring+" DIX SEPT"
	case 8: outputstring=outputstring+" DIX HUIT"
	case 9: outputstring=outputstring+" DIX NEUF"
	end select
return
end function
