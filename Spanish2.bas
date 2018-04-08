function spanish2(inputvalue as double) as string
	dim inputstring as string
	dim ouputstring as string
	dim negative as boolean
	dim leng as integer
	dim triplet as integer
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	if inputvalue>999999999999 then
		spanish2=""
		exit function
	endif
	if inputvalue=0 then
		outputstring="CERO"
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
	case 1: outputstring=outputstring+" MIL"
	case 100: outputstring=outputstring+" CIEN MIL"
	case else
		gosub subxxx
		outputstring=outputstring+" MIL"
	end select
	triplet=cint(mid(inputstring,4,3))
	if outputstring="" then
		select case triplet
		case 0:
		case 1: outputstring=outputstring+" UN MILLION"
		case 100: outputstring=outputstring+" CIEN MILLIONES"
		case else
			gosub subxxx
			outputstring=outputstring+" MILLONES"
		end select
	else
		select case triplet
		case 0: outputstring=outputstring+" MILLONES"
		case 1: outputstring=outputstring+" UN MILLIONES"
		case 100: outputstring=outputstring+" CIEN MILLIONES"
		case else
			gosub subxxx
			outputstring=outputstring+" MILLONES"
		end select
	endif
	triplet=cint(mid(inputstring,7,3))
	select case triplet
	case 0:
	case 1: outputstring=outputstring+" MIL"
	case 100: outputstring=outputstring+" CIEN MIL"
	case else
		gosub subxxx
		outputstring=outputstring+" MIL"
	end select
	triplet=cint(mid(inputstring,10,3))
	select case triplet
	case 0:
	case 100: outputstring=outputstring+" CIEN"
	case else
		gosub subxxx
		if triplet3=1 then outputstring=outputstring+"O"
	end select
	if negative then outputstring="MENOS"+outputstring else	outputstring=mid(outputstring,2,)
	spanish2=outputstring
	exit function
subxxx:
	triplet1=triplet\100
	triplet2=(triplet-(triplet1*100))\10
	triplet3=triplet-triplet1*100-triplet2*10
	select case triplet1
	case 1: outputstring=outputstring+" CIENTO"
	case 2: outputstring=outputstring+" DOSCIENTOS"
	case 3: outputstring=outputstring+" TRESCIENTOS"
	case 4: outputstring=outputstring+" CUATROCIENTOS"
	case 5: outputstring=outputstring+" QUINIENTOS"
	case 6: outputstring=outputstring+" SEISCIENTOS"
	case 7: outputstring=outputstring+" SETECIENTOS"
	case 8: outputstring=outputstring+" OCHOCIENTOS"
	case 9: outputstring=outputstring+" NOVECIENTOS"
	end select
	select case triplet2
	case 0: gosub sub01
	case 1: gosub sub11
	case else: if triplet3 then gosub tensx else gosub tens0
	end select
return
sub01:
	select case triplet3
	case 1: outputstring=outputstring+" UN"
	case 2: outputstring=outputstring+" DOS"
	case 3: outputstring=outputstring+" TRES"
	case 4: outputstring=outputstring+" CUATRO"
	case 5: outputstring=outputstring+" CINCO"
	case 6: outputstring=outputstring+" SEIS"
	case 7: outputstring=outputstring+" SIETE"
	case 8: outputstring=outputstring+" OCHO"
	case 9: outputstring=outputstring+" NUEVE"
	end select
return
sub11:
	select case triplet3
	case 0: outputstring=outputstring+" DIEZ"
	case 1: outputstring=outputstring+" ONCE"
	case 2: outputstring=outputstring+" DOCE"
	case 3: outputstring=outputstring+" TRECE"
	case 4: outputstring=outputstring+" CATORCE"
	case 5: outputstring=outputstring+" QUINCE"
	case 6: outputstring=outputstring+" DIECISEIS"
	case 7: outputstring=outputstring+" DIECISIETE"
	case 8: outputstring=outputstring+" DIECIOCHO"
	case 9: outputstring=outputstring+" DIECINUEVE"
	end select
return
tens0:
	select case triplet2
	case 2: outputstring=outputstring+" VEINTE"
	case 3: outputstring=outputstring+" TREINTA"
	case 4: outputstring=outputstring+" CUARENTA"
	case 5: outputstring=outputstring+" CINCUENTA"
	case 6: outputstring=outputstring+" SESENTA"
	case 7: outputstring=outputstring+" SETENTA"
	case 8: outputstring=outputstring+" OCHENTA"
	case 9: outputstring=outputstring+" NOVENTA"
	end select
return
tensx:
	select case triplet2
	case 2: outputstring=outputstring+" VEINTI"
	case 3: outputstring=outputstring+" TREINTA Y "
	case 4: outputstring=outputstring+" CUARENTA Y "
	case 5: outputstring=outputstring+" CINCUENTA Y "
	case 6: outputstring=outputstring+" SESENTA Y "
	case 7: outputstring=outputstring+" SETENTA Y "
	case 8: outputstring=outputstring+" OCHENTA Y "
	case 9: outputstring=outputstring+" NOVENTA Y "
	end select
	select case triplet3
	case 1: outputstring=outputstring+"UN"
	case 2: outputstring=outputstring+"DOS"
	case 3: outputstring=outputstring+"TRES"
	case 4: outputstring=outputstring+"CUATRO"
	case 5: outputstring=outputstring+"CINCO"
	case 6: outputstring=outputstring+"SEIS"
	case 7: outputstring=outputstring+"SIETE"
	case 8: outputstring=outputstring+"OCHO"
	case 9: outputstring=outputstring+"NUEVE"
	end select
return
end function
