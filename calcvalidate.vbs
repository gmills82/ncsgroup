function isEnough(val_1,val_2)
  isEnough = True
  if (val_1.value <= 0) and (val_2.value <= 0) then
    msgbox "Please input a Contractor Rate or a Billing Rate."
    isEnough=False
  end if
end function

function isNumber(val)
  if len(val.value) > 0 then
	if isNumeric(val.value) then
		isNumber = true
    else
	  msgbox "This value (" & val.name & ") must be numeric."
	  val.select
	  isNumber = false
    end if
  else
    val.value = 0.00
    isNumber=True
  end if	
end function

function isLength(val)
  if len(val.value)>0 then
		isLength = true
  else
		msgbox "This value (" & val.name & ") can't be empty."
		val.select
		isLength = false
  end if
end function

function isProperdate(val)
  if isDate(trim(val.value)) then
		isProperdate = true
  else
		msgbox "Please enter a valid date. (" & val.name & ")"
		val.select
		isProperdate = false
  end if
end function

function isCCnumber(val)
  if isLength(val) then
	ccnumber = false
	val.value = trim(replace(val.value," ",""))
	if isNumeric(val.value) then
	  ccnum = val.value
	  for i = len(ccnum) to 2 step -2
		total = total + cint(mid(ccnum,i,1))
		tmp = cstr((mid(ccnum,i-1,1))*2)
		total = total + cint(left(tmp, 1))
		if len(tmp)>1 then
		  total = total + cint(right(tmp,1))
		end if
	  next
	  if len(ccnum) mod 2 = 1 then
		total = total + cint(left(ccnum,1))
	  end if
	  if total mod 10 = 0 then
		ccnumber = true
	  end if
	end if
	if ccnumber then
	  isCCnumber = true
	else
	  msgbox "Please use a valid credit card number."
	  val.select
	  isCCnumber = false
	end if
  end if
end function

function isCCdate(val)
	if isLength(val) then
		ccdate = false
		if len(val.value)>=3 and len(val.value)<=5 then
			if instr(val.value,"/")>0 then
			   tCC = split(val.value,"/")
			   if isNumeric(tCC(0)) and isNumeric(tCC(1)) then
			      mn = cint(tCC(0))
			      yr = cint(tCC(1))
				   currYear = cint(right(year(date()),2))
				   if mn>0 and mn<13 then
				      if yr = currYear then
				         if mn >= month(date()) then
				            ccdate = true
				         end if
				      else
				         if yr<currYear+4 then
				            ccdate = true
				         end if
				      end if
				   end if
				end if
			end if
		end if
		if ccdate then
			isCCdate = true
		else
			msgbox "Please enter a valid expiration date."
			val.select
			isCCdate = false
		end if
	end if
end function


function isPositive(val)
	if isNumber(val) then
		if val.value>=0 then
			isPositive = true
		else
			msgbox "Please enter a positive number."
			val.select
			isPositive = false
		end if
	end if
end function

function isNegative(val)
	if isNumber(val) then
		if val.value<0 then
			isNegative = true
		else
			msgbox "Please enter a negative number."
			val.select
			isNegative = false
		end if
	end if
end function

function isAlpha(val)
	if isLength(val) then
		isNot=" !@#$^*()_+=-'`~\|]}[{;:/?.,<>&%"
		invalid = false
		if instr(val.value,chr(34))>0 then
			invalid=true
		else
			for i = 1 to len(val.value)
				for x = 1 to len(isNotAlpha)
					if mid(val.value,i,1)=mid(isNot,x,1) then
						invalid = true
					end if
				next
			next
		end if
		if not invalid then
			isAlpha = true
		else
			msgbox "Please use alphanumeric characters."
			val.select
			isAlpha = false
		end if
	end if
end function

function isEmail(val)
	if isLength(val) then
		e = val.value
		if instr(e,"@")>0 and instr(e,".")>0 and len(e)>5 then
			isEmail = true
		else
			msgbox "Please enter a proper email address."
			val.select
			isEmail = false
		end if
	end if
end function

function isZip(val)
	isZip = false
	tVal = trim(val.value)
	if len(tVal)=5 then
		if isNumeric(tVal) then
			isZip = true
		end if
	elseif len(tVal)=10 then
		l5 = left(tVal,5)
		r4 = right(tVal,4)
		m1 = mid(tVal,6,1)
		if isNumeric(l5) and isNumeric(r4) and m1="-" then
			isZip = true
		end if
	end if
	if not isZip then
		msgbox "Please enter a valid Zip code."
		val.select
	end if		
end function

function isPath(val)
	if isLength(val) then
		path = false
		if instr(val.value,":\")>0 then
			isNot=" !@#$^*()'`~|]}[{;.>,<?%&+="
			path = true
			if instr(val.value,chr(34))>0 then
				path = false
			else
				for i = 1 to len(val.value)
					for x = 1 to len(isNot)
						if mid(val.value,i,1)=mid(isNot,x,1) then
							path = false
						end if
					next
				next
			end if
		end if
		if path then
			isPath = true
		else
			msgbox "Please enter a valid local path."
			val.select
			isPath = false
		end if
	end if
end function

function isURL(val)
	if isLength(val) then
		url = false
		if instr(val.value,"://")>0 then
			isNot=" !@$^*()'`~|]}[{;.>,<"
			url = true
			if instr(val.value,chr(34))>0 then
				url = false
			else
				for i = 1 to len(val.value)
					for x = 1 to len(isNot)
						if mid(val.value,i,1)=mid(isNot,x,1) then
							url = false
						end if
					next
				next
			end if
		end if
		if url then
			isURL = true
		else
			msgbox "Please enter a valid URL."
			val.select
			isURL = false
		end if
	end if
end function

