function letter(vNombre as double) as string
	on error goto ErrorHandler
	Dim vOutput as string
	Dim vTriplette as integer
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
	Dim dixTable as Variant
	Dim unitTable as Variant
	if vNombre>999999999999 then
		letter=""
		exit function
	endif
	if vNombre<0 then
		vNegatif=false
		vNombre=abs(vNombre)
	else
		vNegatif=true
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
	if vMilliard then
		vTriplette=vMilliard
		gosub routine
		vOutput=vOutput+" BILLION "
	endif
	if vMillion then
		vTriplette=vMillion
		gosub routine
		vOutput=vOutput+" MILLION "
	endif
	if vMille then
		vTriplette=vMille
		gosub routine
		vOutput=vOutput+" THOUSAND "
	endif
	if vUnite then
		vTriplette=vUnite
		gosub routine
	endif
	if vNombre<1 then vOutput="ZERO"
	if vNegatif=false then vOutput="MINUS "+vOutput
	letter=vOutput
	exit function
routine:
	unitTable=array("","ONE","TWO","THREE","FOUR","FIVE","SIX","SEVEN","EIGHT","NINE","TEN","ELEVEN","TWELVE","THIRTEEN","FOURTEEN","FIFTEEN","SIXTEEN","SEVENTEEN","EIGHTEEN","NINETEEN")
	dixTable=array("","","TWENTY","THIRTY","FORTY","FIFTY","SIXTY","SEVENTY","EIGHTY","NINETY")
	if vTriplette then
		vCentaine=int(vTriplette/100)
		vDizaine=int((vTriplette-(vCentaine*100))/10)
		vUnites=vTriplette-vCentaine*100-vDizaine*10
	endif
	if vCentaine then vOutput=vOutput+unitTable(vCentaine)+" HUNDRED"
	if vTriplette-vCentaine*100<20 then
		vOutput=vOutput+" "+unitTable(vTriplette-vCentaine*100)
	else
		vOutput=vOutput+" "+dixTable(vDizaine)+" "+unitTable(vUnites)
	endif
return
ErrorHandler:
	letter=""
end function
