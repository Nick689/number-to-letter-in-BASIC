Credit go to Fernand COSTA
 
 
  
# number-to-letter-in-ENGLISH
 
```
function numtoletter(vNombre as double) as string
	on error goto ErrorHandler
	Dim vOutput as string
	Dim vGroupe as integer
	Dim vMilliard as integer
	Dim vMillion as integer
	Dim vMille as integer
	Dim vUnite as integer
	Dim vNegatif as boolean
	Dim vString as string
	Dim vlength as integer
	Dim vCentaine as integer
	Dim vDizaine as integer
	Dim vUnites as integer
	Dim tVingt as Variant
	Dim tDix as Variant
	Dim tUnite as Variant
	if vNombre<0 then
		vNegatif=false
		vNombre=abs(vNombre)
	else
		vNegatif=true
	end if
	if vNombre>999999999999 then
		numtoletter=""
		exit function
	end if	
	vString=cstr(vNombre)
	vlength=len(vString)
	while vlength<12
		vString=" "+vString
		vlength=vlength+1
	wend
	vMilliard=cint(mid(vString,1,3))
	vMillion=cint(mid(vString,4,3))
	vMille=cint(mid(vString,7,3))
	vUnite=cint(mid(vString,10,3))
	if vOutput<>"" then vOutput=vOutput+" "
	if vMilliard then
		vGroupe=vMilliard
		gosub Groupe
		vOutput=vOutput+" BILLION"
	end if
	if vMillion then
		if vOutput<>"" then vOutput=vOutput+" "
		vGroupe=vMillion
		gosub Groupe
		vOutput=vOutput+" MILLION"
	end if
	if vMille then
		if vOutput<>"" then vOutput=vOutput+" "
		if vMille>=2 then
			vGroupe=vMille
			gosub Groupe
			vOutput=vOutput+" THOUSAND"
		else
			vOutput=vOutput+"THOUSAND"
		end if	
	end if
	if vUnite then
		if vOutput<>"" then vOutput=vOutput+" "
		vGroupe=vUnite
		gosub Groupe
	end if
	if vNombre<1 then vOutput="ZERO"
	if vNegatif=false then vOutput="- "+vOutput
	numtoletter=vOutput
	exit function
Groupe:
	tUnite=array("ONE","TWO","THREE","FOUR","FIVE","SIX","SEVEN","EIGHT","NINE","TEN","ELEVEN","TWELVE","THIRTEEN","FOURTEEN","FIFTEEN","SIXTEEN","SEVENTEEN","EIGHTEEN","NINETEEN")
	tDix=array("TWENTY","THIRTY","FORTY","FIFTY","SIXTY","SEVENTY","EIGHTY","NINETY")
	if vGroupe then
		vCentaine=int(vGroupe/100)
		vDizaine=int((vGroupe-(vCentaine*100))/10)
		vUnites=vGroupe-vCentaine*100-vDizaine*10
	endif
	if vCentaine then
		if vCentaine=1 then
			vOutput=vOutput+"ONE HUNDRED"
		else
			vOutput=vOutput+tUnite(vCentaine-1)+" HUNDRED"
		endif
	endif
	if ((vUnites or vDizaine) and vOutput<>"") then vOutput=vOutput+" "
	select case vDizaine
	case 0,1
		if vUnites or vDizaine then vOutput=vOutput+tUnite(vDizaine*10+vUnites-1)
	case else
		vOutput=vOutput+tDix(vDizaine-2)
		if vUnites>1 then vOutput=vOutput+" "+tUnite(vUnites-1)
	end select
return
ErrorHandler:
	numtoletter=""
end function
```
 
 
  
#number-to-letter-in-FRENCH
 
```
function numbertoletter(vNombre as double) as string
	on error goto ErrorHandler
	Dim vOutput as string
	Dim vGroupe as integer
	Dim vMilliard as integer
	Dim vMillion as integer
	Dim vMille as integer
	Dim vUnite as integer
	Dim vNegatif as boolean
	Dim vString as string
	Dim vlength as integer
	Dim vCentaine as integer
	Dim vDizaine as integer
	Dim vUnites as integer
	Dim tVingt as Variant
	Dim tDix as Variant
	Dim tUnite as Variant
	if vNombre<0 then
		vNegatif=false
		vNombre=abs(vNombre)
	else
		vNegatif=true
	endif
	if vNombre>999999999999 then
		numtoletter=""
		exit function
	end if
	vString=cstr(vNombre)
	vlength=len(vString)
	while vlength<12
		vString=" "+vString
		vlength=vlength+1
	wend
	vMilliard=cint(mid(vString,1,3))
	vMillion=cint(mid(vString,4,3))
	vMille=cint(mid(vString,7,3))
	vUnite=cint(mid(vString,10,3))
	if vOutput<>"" then vOutput=vOutput+" "
	if vMilliard then
		vGroupe=vMilliard
		gosub Groupe
		vOutput=vOutput+" MILLIARD"
		if vMilliard>1 then vOutput=vOutput+"S"
	end if
	if vMillion then
		if vOutput<>"" then vOutput=vOutput+" "
		vGroupe=vMillion
		gosub Groupe
		vOutput=vOutput+" MILLION"
		if vMillion>1 then vOutput=vOutput+"S"
	end if
	if vMille then
		if vOutput<>"" then vOutput=vOutput+" "
		if vMille>=2 then
			vGroupe=vMille
			gosub Groupe
			vOutput=vOutput+" MILLE"
		else
			vOutput=vOutput+"MILLE"
		end if	
	end if
	if vUnite then
		if vOutput<>"" then vOutput=vOutput+" "
		vGroupe=vUnite
		gosub Groupe
	end if
	if vNombre<1 then vOutput="ZÉRO"
	if vNegatif=false then vOutput="MOINS "+vOutput
	numbertoletter=vOutput
	exit function
Groupe:
	tUnite=array("UN","DEUX","TROIS","QUATRE","CINQ","SIX","SEPT","HUIT","NEUF","DIX","ONZE","DOUZE","TREIZE","QUATORZE","QUINZE","SEIZE","DIX-SEPT","DIX-HUIT","DIX-NEUF")
	tDix=array("VINGT","TRENTE","QUARANTE","CINQUANTE","SOIXANTE","SOIXANTE-DIX","QUATRE-VINGT","QUATRE-VINGT-DIX")
	if vGroupe then
		vCentaine=int(vGroupe/100)
		vDizaine=int((vGroupe-(vCentaine*100))/10)
		vUnites=vGroupe-vCentaine*100-vDizaine*10
	endif
	if vCentaine then
		if vCentaine=1 then
			vOutput=vOutput+"CENT"
		else
			vOutput=vOutput+tUnite(vCentaine-1)+" CENT"
		endif
	endif
	if ((vUnites or vDizaine) and vOutput<>"") then vOutput=vOutput+" "
	select case vDizaine
	case 0,1
		if vUnites or vDizaine then vOutput=vOutput+tUnite(vDizaine*10+vUnites-1)
	case 2,3,4,5,6,8
		vOutput=vOutput+tDix(vDizaine-2)
		if vDizaine=8 and vUnites=0 then vOutput=vOutput+"S"
		if vUnites=1 then
			if vDizaine=8 then vOutput=vOutput+"-UN" else vOutput=vOutput+" ET UN"
		else
			if vUnites>1 then vOutput=vOutput+"-"+tUnite(vUnites-1)
		end if
	case 7,9
		vOutput=vOutput+tdix(vDizaine-3)
		select case vUnites
		case 0
			vOutput=vOutput+"-DIX"
		case 1
			if vDizaine=7 then vOutput=vOutput+" ET ONZE" else vOutput=vOutput+"-ONZE"
		case else
			vOutput=vOutput+"-"+tunite(vunites+9)
		end select
	end select
return
ErrorHandler:
	numbertoletter=""
end function
```
