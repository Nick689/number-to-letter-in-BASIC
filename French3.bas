function french3(inputvalue as double) as string
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
		french3=""
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
			case 0
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
			case 1
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
			case 2
				select case triplet3
				case 0: outputstring=outputstring+" VINGT"
				case 1: outputstring=outputstring+" VINGT ET UN"
				case 2: outputstring=outputstring+" VINGT DEUX"
				case 3: outputstring=outputstring+" VINGT TROIS"
				case 4: outputstring=outputstring+" VINGT QUATRE"
				case 5: outputstring=outputstring+" VINGT CINQ"
				case 6: outputstring=outputstring+" VINGT SIX"
				case 7: outputstring=outputstring+" VINGT SEPT"
				case 8: outputstring=outputstring+" VINGT HUIT"
				case 9: outputstring=outputstring+" VINGT NEUF"
				end select
			case 3
				select case triplet3
				case 0: outputstring=outputstring+" TRENTE"
				case 1: outputstring=outputstring+" TRENTE ET UN"
				case 2: outputstring=outputstring+" TRENTE DEUX"
				case 3: outputstring=outputstring+" TRENTE TROIS"
				case 4: outputstring=outputstring+" TRENTE QUATRE"
				case 5: outputstring=outputstring+" TRENTE CINQ"
				case 6: outputstring=outputstring+" TRENTE SIX"
				case 7: outputstring=outputstring+" TRENTE SEPT"
				case 8: outputstring=outputstring+" TRENTE HUIT"
				case 9: outputstring=outputstring+" TRENTE NEUF"
				end select
			case 4
				select case triplet3
				case 0: outputstring=outputstring+" QUARANTE"
				case 1: outputstring=outputstring+" QUARANTE ET UN"
				case 2: outputstring=outputstring+" QUARANTE DEUX"
				case 3: outputstring=outputstring+" QUARANTE TROIS"
				case 4: outputstring=outputstring+" QUARANTE QUATRE"
				case 5: outputstring=outputstring+" QUARANTE CINQ"
				case 6: outputstring=outputstring+" QUARANTE SIX"
				case 7: outputstring=outputstring+" QUARANTE SEPT"
				case 8: outputstring=outputstring+" QUARANTE HUIT"
				case 9: outputstring=outputstring+" QUARANTE NEUF"
				end select
			case 5
				select case triplet3
				case 1: outputstring=outputstring+" CINQUANTE ET UN"
				case 2: outputstring=outputstring+" CINQUANTE DEUX"
				case 3: outputstring=outputstring+" CINQUANTE TROIS"
				case 4: outputstring=outputstring+" CINQUANTE QUATRE"
				case 5: outputstring=outputstring+" CINQUANTE CINQ"
				case 6: outputstring=outputstring+" CINQUANTE SIX"
				case 7: outputstring=outputstring+" CINQUANTE SEPT"
				case 8: outputstring=outputstring+" CINQUANTE HUIT"
				case 9: outputstring=outputstring+" CINQUANTE NEUF"
				end select
			case 6
				select case triplet3
				case 0: outputstring=outputstring+" SOIXANTE"
				case 1: outputstring=outputstring+" SOIXANTE ET UN"
				case 2: outputstring=outputstring+" SOIXANTE DEUX"
				case 3: outputstring=outputstring+" SOIXANTE TROIS"
				case 4: outputstring=outputstring+" QUATRE"
				case 5: outputstring=outputstring+" SOIXANTE CINQ"
				case 6: outputstring=outputstring+" SOIXANTE SIX"
				case 7: outputstring=outputstring+" SOIXANTE SEPT"
				case 8: outputstring=outputstring+" SOIXANTE HUIT"
				case 9: outputstring=outputstring+" SOIXANTE NEUF"
				end select
			case 7
				select case triplet3
				case 0: outputstring=outputstring+" SOIXANTE DIX"
				case 1: outputstring=outputstring+" SOIXANTE ET ONZE"
				case 2: outputstring=outputstring+" SOIXANTE DOUZE"
				case 3: outputstring=outputstring+" SOIXANTE TREIZE"
				case 4: outputstring=outputstring+" SOIXANTE QUATORZE"
				case 5: outputstring=outputstring+" SOIXANTE QUINZE"
				case 6: outputstring=outputstring+" SOIXANTE SEIZE"
				case 7: outputstring=outputstring+" SOIXANTE DIX SEPT"
				case 8: outputstring=outputstring+" SOIXANTE DIX HUIT"
				case 9: outputstring=outputstring+" SOIXANTE DIX NEUF"
				end select
			case 8
				select case triplet3
				case 0: outputstring=outputstring+" QUATRE VINGT"
				case 1: outputstring=outputstring+" QUATRE VINGT UN"
				case 2: outputstring=outputstring+" QUATRE VINGT DEUX"
				case 3: outputstring=outputstring+" QUATRE VINGT TROIS"
				case 4: outputstring=outputstring+" QUATRE VINGT QUATRE"
				case 5: outputstring=outputstring+" QUATRE VINGT CINQ"
				case 6: outputstring=outputstring+" QUATRE VINGT SIX"
				case 7: outputstring=outputstring+" QUATRE VINGT SEPT"
				case 8: outputstring=outputstring+" QUATRE VINGT HUIT"
				case 9: outputstring=outputstring+" QUATRE VINGT NEUF"
				end select
			case 9
				select case triplet3
				case 0: outputstring=outputstring+" QUATRE VINGT DIX"
				case 1: outputstring=outputstring+" QUATRE VINGT ONZE"
				case 2: outputstring=outputstring+" QUATRE VINGT DOUZE"
				case 3: outputstring=outputstring+" QUATRE VINGT TREIZE"
				case 4: outputstring=outputstring+" QUATRE VINGT QUATORZE"
				case 5: outputstring=outputstring+" QUATRE VINGT QUINZE"
				case 6: outputstring=outputstring+" QUATRE VINGT SEIZE"
				case 7: outputstring=outputstring+" QUATRE VINGT DIX SEPT"
				case 8: outputstring=outputstring+" QUATRE VINGT DIX HUIT"
				case 9: outputstring=outputstring+" QUATRE VINGT DIX NEUF"
				end select
			end select
			select case i
			case 1: outputstring=outputstring+" MILLIARDS"
			case 4: outputstring=outputstring+" MILLIONS"
			case 7: outputstring=outputstring+" MILLE"
			end select
		end select
	next
	if negative then outputstring="MOINS"+outputstring else	outputstring=mid(outputstring,2,)
	french3=outputstring
end function
