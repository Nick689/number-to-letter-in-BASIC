function spanish(inputvalue as double) as string
	dim inputstring as string
	dim outputstring as string
	dim negative as boolean
	dim leng as integer
	dim unit01 as variant
	dim unit11 as variant
	dim tens0 as variant
	dim tensx as variant
	dim cents as variant
	dim triplet as integer
	dim triplet1 as integer
	dim triplet2 as integer
	dim triplet3 as integer
	dim i as integer
	unit01=array("","UN","DOS","TRES","CUATRO","CINCO","SEIS","SIETE","OCHO","NUEVE")
	unit11=array(" DIEZ"," ONCE"," DOCE"," TRECE"," CATORCE"," QUINCE"," DIECISEIS"," DIECISIETE"," DIECIOCHO"," DIECINUEVE")
	tens0=array("",""," VEINTE"," TREINTA"," CUARENTA"," CINCUENTA"," SESENTA"," SETENTA"," OCHENTA"," NOVENTA")
	tensx=array("",""," VEINTI"," TREINTA Y "," CUARENTA Y "," CINCUENTA Y "," SESENTA Y "," SETENTA Y "," OCHENTA Y "," NOVENTA Y ")
	cents=array(""," CIENTO"," DOSCIENTOS"," TRESCIENTOS"," CUATROCIENTOS"," QUINIENTOS"," SEISCIENTOS"," SETECIENTOS"," OCHOCIENTOS"," NOVECIENTOS")
	if inputvalue>999999999999 then
		spanish=""
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
	for i=1 to 10 step 3
		triplet=cint(mid(inputstring,i,3))
		select case triplet
		case 0: if i=4 then if outputstring<>"" then outputstring=outputstring+" MILLONES"
		case 1
			select case i
			case 4: if outputstring="" then outputstring=" UN MILLON" else outputstring=outputstring+" UN MILLONES"
			case 1,7: outputstring=outputstring+" MIL"
			case 10: outputstring=outputstring+" UNO"
			end select
		case 100
			select case i
			case 1,7: outputstring=outputstring+" CIEN MIL"
			case 4:	outputstring=outputstring+" CIEN MILLONES"
			case 10: outputstring=outputstring+" CIEN"
			end select
		case else
			triplet1=cint(mid(inputstring,i,1))
			triplet2=cint(mid(inputstring,i+1,1))
			triplet3=cint(mid(inputstring,i+2,1))
			select case triplet1
			case 0:
			case else: outputstring=outputstring+cents(triplet1)
			end select
			select case triplet2
			case 0: outputstring=outputstring+" "+unit01(triplet3)
			case 1: outputstring=outputstring+unit11(triplet3)
			case else: if triplet3 then outputstring=outputstring+tensx(triplet2)+unit01(triplet3) else outputstring=outputstring+tens0(triplet2)
			end select
			select case i
			case 1,7: outputstring=outputstring+" MIL"
			case 4:	outputstring=outputstring+" MILLONES"
			case 10: if triplet3=1 then outputstring=outputstring+"O"
			end select
		end select
	next
	if negative then
		outputstring="MENOS"+outputstring
	else
		outputstring=mid(outputstring,2,)
	endif
	spanish=outputstring
end function
