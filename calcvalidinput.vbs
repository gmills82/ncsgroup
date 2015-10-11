function isEnough(val1,val2,val3,val4)
  isEnough = True
  if (val1.value <= 0) and (val2.value <= 0) and (val3.value <=0) then
    msgbox "Please input a Contractor Rate and/or a Billing Rate."
    isEnough=False
  end if
  if (val4.value > 0) and (val1.value +val3.value)<= 0 then
    strMsg = "You have specified perdiem and a billing rate, "
	strMsg = strMsg & Chr(13) & "but did not specify a contractor rate."
	strMsg = strMsg & chr(13) & chr(13) & "If the contractor rate shown on the next screen,"
	strMsg = strMsg & chr(13) & "is less than $28/hr, it will be displayed in red."
	strMsg = strMsg & Chr(13) & chr(13) & "You will need to return to this screen and"
	strMsg = strMsg & Chr(13) & "increase the billing rate and re-calculate."
	msgbox strMsg
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

function isValidCTH(val1, val2)
  if val1.value > 1 and val2.value <= 0 then
    isValidCTH = false
	msgbox "You MUST specify a contractor rate when calculating Contract-to-Hire with no fee." 
  else
    isValidCTH = true
  end if
end function  	
   
function isMaxLength(val,iChars)
  val.value=Trim(val.value)
  
  ' If skills field add spaces after each comma
  if val.name="Skills" then
	'Remove any commas at the end
	Do While right(val.value,1)="," and len(val.value)>0
	 val.value=left(val.value,len(val.value)-1)
	Loop
	'Look for commas in string
	bolSpace = false
	strLastChar = Left(val.value,1)
	strValue = strLastChar
	for iLoop = 2 to len(val.value)
	  strChar = Mid(val.value,iLoop,1)
	  'Add space after comma if not already there
	  If strLastChar = "," and strChar <> " " then
	    strValue = strValue & " "
		bolSpace = True   
      end if
	  strValue = strValue + strChar
	  strLastChar = strChar
	next
	val.value = strValue
  end if
  ' end skills field manipulation
    
  if len(val.value) <= iChars then
    isMaxLength = false
  else
    msgbox "This field (" & val.name & ") exceeds maximum allowed characters. - " & len(val.value) & " out of " & iChars
    if bolSpace then msgbox "Spaces have been added to the Skills field for page formatting." 
	val.select
    isMaxLength = true
  end if
end function 

function isDaily(val)
  isDaily = true
  if val.value > 0 and val.value < 50 then
    resp = MsgBox("Is your Per Diem a DAILY rate?", vbYesNo + vbDefaultButton2 + vbQuestion, "PerDiem")
    If resp = vbNo Then isDaily = false
  end if		
end function		

function isAboveMin(val1, val2, val3)
 MINRATE=28.00
 isAboveMin = true
 if val1.value > 0 and val1.value < 28.00 then
   if val3.value > 0 then
	  NewPD = val1.value + val3.value/8.0 - MINRATE 
	  strMsg1 = "Minimum taxable rate when paying perdiem is $"
	  strMsg1 = strMsg1 + Trim(MINRATE) & "/hr"
	  If NewPD > 0 then
	   strMsg2 = Chr(13) & Chr(13)
	   strMsg2= StrMsg2 & "To keep the same contractor payout of "
       strMsg2= StrMsg2 & FormatCurrency(NewPD+MINRATE,2)
	   If val1.value > 0 and val2.value=0 then
	    strMsg2= StrMsg2 & Chr(13)
		strMsg2= strMsg2 & "Increase the contractor's base rate to "
	   	strMsg2= strMsg2 & FormatCurrency(MINRATE, 2)& " and"
	  end if	
		strMsg2= strMsg2 & Chr(13) &"Reduce the daily perdiem to "
		strMsg2= strMsg2 & FormatCurrency(NewPD*8.0,2)
	  end if
	  isAboveMin =false
      If val2.value > 0 then val1.value = 0 
	else
	  strMsg1 = ""
	  strMsg2 = "Sales manager approval is required,"
	  strMsg2 = strMsg2 & chr(13) & "if contractor will be working overtime."
	end if
    MsgBox strMsg1 & strMsg2
  end if	
end function		

function isPayOut(val1,val2,val3)
  isPayOut = false
  if val1.value > 0 and val2.value > 0 then
    strMsg = "Press the RESET button and "
    strMsg = strMsg & chr(13) & "Specify a base rate OR a total payout." 
    strMsg = strMsg & Chr(13) & Chr(13) & "Total Payout is the amount you want to"
	strMsg = strMsg & Chr(13) & "pay the contractor including per diem."
    Msgbox strMsg
	isPayOut = true
  else
	if val2.value > 0 then
	  val1.value = FormatNumber(val2.value - val3.value/8.0,2)
	end if
  end if		
end function		

function isValidStaffFee(val1, val2)
  if val1.value > 0 and val2.value > 0 then
    isValidStaffFee = false
	msgbox "A Staffing Fee can only be specified on Straight Contract position." 
  else
    isValidStaffFee = true
  end if
end function  	
