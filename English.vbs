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
	endif
	if vNombre>999999999999 then
		numtoletter=""
		exit function
	endif	
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
	endif
	if vMillion then
		if vOutput<>"" then vOutput=vOutput+" "
		vGroupe=vMillion
		gosub Groupe
		vOutput=vOutput+" MILLION"
	endif
	if vMille then
		if vOutput<>"" then vOutput=vOutput+" "
		if vMille>=2 then
			vGroupe=vMille
			gosub Groupe
			vOutput=vOutput+" THOUSAND"
		else
			vOutput=vOutput+"THOUSAND"
		endif	
	endif
	if vUnite then
		if vOutput<>"" then vOutput=vOutput+" "
		vGroupe=vUnite
		gosub Groupe
	endif
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
