<%@ LANGUAGE=VBScript %>
<% Option Explicit %>
<SCRIPT LANGUAGE=VBScript>
<!--#include file="calcsecurity.asp"--> 
<!--#include file="calcvalidinput.vbs"-->
</SCRIPT>

<%
  'Define Constants
  Const ADMIN=0.03, TAXRATE=0.12, MINRATE=28, MINSPREAD=12
 
  'Declare form field variables
  Dim curContRate, curBillRate
  Dim curW2RefFee, curIncRefFee
  Dim curPerDiemDly, curPerDiemHrly
  Dim curOtherAmt, curStaffAmt 
  Dim strOtherDesc, strCalcType 
  Dim arrSpread(2,166),intCTHlength
 
    'Declare variables for calculation
  Dim bolSpreadLow, bolRateLow, bolStaffFee, bolCostBased, bolBillBased
  Dim curSubTot1, curSubTot2, curTotalCost
  Dim curAdjBilling, curContPayOut, curAdjTotalCost
  Dim curAdjSpread, curReqSpread, curCTHSpread 
  Dim curTaxAmt, curAdminAmt, bolNoFeeCTH
  Dim curActStaffFee, curReqStaffFee, pctIncrease
  Dim curTempTotalCost, curTempBillRate
  curCTHSpread = 0
  
   'Read in form field variables
  strCalcType=Request("CalcType")
  curContRate=Request("ContRate")*1.0
  curBillRate=Request("BillRate")*1.0
  curW2RefFee=Request("W2RefFee")*1.0
  curIncRefFee=Request("IncRefFee")*1.0
  curPerDiemDly=Request("PerDiem")*1.0
  curOtherAmt=Request("OtherAmt")*1.0
  curReqStaffFee=Request("ReqStaffFee")*1.0
  strOtherDesc=Request("OtherDesc")
  intCTHlength=Request("CTHtype")*1.0
  
  'Calculate hourly per diem
  curPerDiemHrly=curPerDiemDly/8.0

 'Determine whether this is a no fee calculation
 bolNoFeeCTH = false
 if intCTHLength > 1 then bolNoFeeCTH = true
 
 'Determine if this calculation has a staffing fee
 bolStaffFee = false
 if curReqStaffFee > 0 then bolStaffFee = true 
 
  '1. Billing rate and Contractor rate both input
  '   Move input billing rate to adjusted billing rate 
  curAdjBilling = 0
  bolCostBased = false
  If (curBillRate > 0) and (curContRate > 0 )then 
    curAdjBilling = curBillRate
	curBillRate= 0
	  'If staffing fee, bill rate, and contractor rate,
   ' then calculate staffing fee as percent of adjusted billing rate
   curStaffAmt = 0
   If curReqStaffFee > 0 then
     curActStaffFee = curReqStaffFee
	 curStaffAmt = curAdjBilling * (curReqStaffFee/100.0)
     bolCostBased = false
	 bolBillBased = true
   end if
 end if	 
  
  '2. No Contractor rate input. Must have billing rate (Client side validation)
  If (curContRate <= 0) then
    curAdjBilling = curBillRate
	
	'If staffing fee and bill rate given, calculate amount as percent of billing rate
	curStaffAmt = 0
	 If curReqStaffFee > 0 then
	    curActStaffFee = curReqStaffFee
	 	curStaffAmt = curBillRate * (curReqStaffFee/100.0)
	 end if
	
   'Calculate spread based on billing rate
     curReqSpread=Round((curBillRate * 0.25)-2.25 +0.005,2 )
  
   'If below minimum spread, the set to minimum spread
   if curReqSpread < MINSPREAD then curReqSpread = MINSPREAD
  
  'Calculate contractor rate from spread and billing rate for W-2 or INC/SUB
    if strCalcType = "W-2" then
      curContRate = Round(((curBillRate-curReqSpread-curStaffAmt)/ (1+ADMIN)-(1+TAXRATE)* curW2RefFee-curINCRefFee-curPerDiemHrly-curOtherAmt) / (1+TAXRATE), 2)
    else
      curContRate = Round((curBillRate-curReqSpread-curStaffAmt)/(1+ADMIN)-curW2RefFee-Round(TAXRATE*curW2RefFee)-curINCRefFee-curOtherAmt , 2)  
    end if
  'Check for contractor rate less than zero
	If curContRate < 0 then
	  response.write "Contractor rate cannot be less than zero."
	  response.end
	end if
  end if
  'Check for W-2 rate less than the minimum rate when paying per diem
   bolRateLow = false
   If (strCalcType = "W-2") and (curContRate < MINRATE) Then
     bolRateLow = true
	 If curPerDiemHrly > 0 Then
       response.write "***************************************************"
	   response.write "Minimum W2 taxable rate when paying per diem is $"
       response.write  MINRATE & "/hr"
	 Else
       response.write "*********************************************"
	   response.write "Sale's manager approval is required "
       response.write "if contractor will be working overtime."
	 End If
   End If 
   
   'Calculate the total cost
   curSubTot1= curContRate + curW2RefFee
   If strCalcType="W-2" then
     curTaxAmt = Round(curSubTot1* TAXRATE,2)
   else
    curTaxAmt = Round(curW2RefFee * TAXRATE,2)
   end if
   curSubTot1 = curSubTot1 + curIncRefFee
   If curSubTot1 > 0 Then
     curSubTot2 = curSubTot1 + curTaxAmt + curOtherAmt + curPerDiemHrly
     curAdminAmt = Round(curSubTot2 * ADMIN , 2)
     curTotalCost = curSubTot2 + curAdminAmt +curStaffAmt
   End If
  
  '3. No Billing rate input.  Must have contractor rate input (client side validation)
  If (curBillRate<=0) then
   bolCostBased=true
  ' Calculate the spread based on the total cost
   curReqSpread=Round(curTotalCost *0.33333 -3.0,2)
   if curReqSpread < MINSPREAD then curReqSpread = MINSPREAD
  
  ' If this as a no fee CTH, calculate the CTH spread
    if bolNoFeeCTH then
	 curCTHSpread = ((curTotalCost + curReqSpread)/intCTHlength + (6/intCTHLength)*curReqSpread)-curReqSpread
    end if
  ' Calculate the billing rate
    curBillRate = curTotalCost + curReqSpread + curCTHspread
  end if
  
  '4. Set up loop to gradually increase the Staffing Fee until it exceeds the required amount
  '   (Only when billing rate is not entered)  
  If bolStaffFee and bolCostBased and not bolBillBased then
    curActStaffFee = 0
	pctIncrease = 0.0
	Do While curActStaffFee < curReqStaffFee
	  curStaffAmt = curBillRate * (curReqStaffFee/100 + pctIncrease) 
	  curTempTotalCost = curTotalCost + curStaffAmt
	  curReqSpread=Round(curTempTotalCost*0.33333-3.0,2)
      if curReqSpread < MINSPREAD then curReqSpread = MINSPREAD
	  curTempBillRate = curTempTotalCost + curReqSpread
      ' Check percentage
	    If curStaffAmt/curTempBillRate * 100 < curReqStaffFee then
          pctIncrease = pctIncrease + 0.0005	  
	    else
		  curActStaffFee = Round(curStaffAmt,2)/Round(curTempBillRate,2) *100
		  curTotalCost = curTotalCost + curStaffAmt
		  curBillRate=curTotalCost+curReqSpread
		end if
	Loop
  end if

 'Calculate the Actual(Adjusted)Totals
  curContPayOut = curContRate + curPerDiemHrly
  curAdjTotalCost = curTotalCost
  If curAdjBilling=0 Then curAdjBilling = curBillRate
  curAdjSpread = curAdjBilling - curAdjTotalCost
 
  'If adjusted spread is less than required spread set variable
  bolSpreadLow = false
  If (curReqSpread + curCTHSpread - curAdjSpread) > 0.01 Then bolSpreadLow = true
 %> 

<html>
<head>
  <style>
  <!--
    TD{font: 9pt arial:}
  -->
  </style>
  <title>Cost Calculator for W-2 Contractor</title>
</head>
<body  background="images/background.jpg"  topmargin="0" leftmargin="0" style="font-family: arial; font-size: 8pt;">
<table border="1" align="center" id="calcMainTable">
<tr>
<td>
<table align="center"  cellpadding="5" cellspacing="0" bgcolor="#FFFFFF" width="770">
  <tr><!-- Start Row 1 --> 
    <td width="300" height="74" align="center" valign="top"><img border="0" src="images/logo.gif" width="206" height="76"></td>
    <td align="left" valign="top" rowspan="2" width="408"> <br><br>
        <table border="1" width="300" align="center" bgcolor="#FFFFFF">
          <tr> 
		    <%
			'Modify contractor rate label depending on calculation type
			If strCalcType = "W-2" then 
			  response.write " <td width=160>Contractor Rate (W2)</td>"
			elseif strCalcType = "INC" then
			  response.write " <td width=160>Contractor Rate (INC)</td>"
			elseif strCalcType = "SUB" then
			  response.write " <td width=160>Contractor Rate (SUB)</td>"
			else
			 response.write " <td width=160>Contractor Rate</td>"
			end if
			if bolRateLow then
			  response.write "<td bgcolor = #FF9999 width=88 align=right>"
			else
			  response.write "<td width=88 align=right>"
			end if
			response.write FormatNumber(curContRate, 2)
			response.write "</td>"
			%>
          
		  </tr>
          <tr> 
            <td width="160">W-2 Referral Fee</td>
            <td align="right" width="88"><%=FormatNumber(curW2RefFee,2)%></td>
          </tr>
          <tr> 
            <td width="160">INC Referral Fee</td>
            <td align="right" width="88"><%=FormatNumber(curIncRefFee,2)%></td>
          </tr>
          <tr> 
            <td align="right" width="160">Sub Total</td>
            <td align="right" width="88"><%=FormatNumber(curSubTot1,2)%></td>
          </tr>
          <tr> 
            <td nowrap width="160">W-2 Required Taxes</td>
            <td align="right" width="88"><%=FormatNumber(curTaxAmt,2)%></td>
          </tr>
          <tr> 
            <td width="160">Other:&nbsp; <u><%=strOtherDesc%></u></td>
            <td align="right" width="88"><%=FormatNumber(curOtherAmt,2)%></td>
          </tr>
          <%
		    if strCalcType="W-2" then
			  response.write "<tr>"
              response.write "<td width=160>Per Diem:  "
			  response.write FormatCurrency(curPerDiemDly,2)
			  response.write " / day</td>"
              response.write "<td align=right width=88>"
			  response.write (FormatNumber(curPerDiemHrly,2))
			  response.write "</td>"
			  response.write "</tr>"
			end if
		  %>	
          <tr> 
            <td align="right" width="160">Sub Total</td>
            <td align="right" width="88"><%=FormatNumber(curSubTot2,2)%></td>
          </tr>
          <tr> 
            <td width="160">Administrative Cost</td>
            <td align="right" width="88"><%=FormatNumber(curAdminAmt,2)%></td>
          </tr>
		  
          <% If bolStaffFee then
			response.write "<tr>"
			response.write "<td width=160>Staff Fee @ "
			response.write FormatNumber(curActStaffFee,2)
			response.write "% of Billing</td>"
			response.write "<td align=right width=88>"
			response.write FormatNumber(curStaffAmt,2)
			response.write "</td></tr>"
			end if
          %>
          <tr> 
            <td align="right" width="160">Total Cost</td>
            <td align="right" width="88"><%=FormatNumber(curTotalCost,2)%></td>
          </tr>
         
          <% If bolNoFeeCTH then 
			   response.write "<tr><td align=right width=160>"
			   response.write "Base Spread</td>"
			   response.write "<td align=right width=88>"
			   response.write FormatNumber(curReqSpread,2)
			   response.write "</td></tr>"
			   response.write "<tr><td align=right width=160>"
			   response.write "C-T-H Spread</td>"
			   response.write "<td align=right width=88>"
			   response.write FormatNumber(curCTHSpread,2)
			   response.write "</td></tr>"
			 else
			   response.write "<tr><td align=right width=160>"
			   response.write "Required Spread</td>"
			   response.write "<td align=right width=88>"
			   response.write FormatNumber(curReqSpread,2)
			   response.write "</td></tr>"
			 end if
		   %>  	 	 
			<tr> 
            <td align="right" width="160">Billing Rate</td>
            <td align="right" width="88"><%=FormatNumber(curBillRate,2)%></td>
          </tr>
		</table>
      </tr>
  <!-- End Row 1 --> 
  <tr><!-- Start Row 2 --> <!-- Row 2 Col 1 --> 
    <td width="300" align="center" valign="top" height="220"> <br><br>
      <div align="center"><b>New Calculation</b><br>
        <a href="calculator.asp?CalcType=W-2">W-2</a>&nbsp; &nbsp;&nbsp;&nbsp;
		<a href="calculator.asp?CalcType=INC">&nbsp;INC</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="calculator.asp?CalcType=SUB">&nbsp;SUB</a>
	 
        <br>
        <br>
        <font color="#000000">Adjusted Totals:</font><br>
        <table width="150" border="1">
          <tr bgcolor="#CCFFCC"> 
            <td>Payout</td>
            <td align="right"><%=FormatNumber(curContPayOut,2)%></td>
            </
          <tr> 
          <tr bgcolor="#CCFFCC"> 
            <td> Cost</td>
            <td align="right"><%=FormatNumber(curAdjTotalCost,2)%></td>
          </tr>
          <tr bgcolor="#CCFFCC"> 
		    <td>Spread</td>
            <%
			if bolSpreadLow then
			  response.write "<td bgcolor = #FF9999 align=right>"
			else 
			  response.write "<td align=right>"
			end if 
			response.write FormatNumber(curAdjSpread,2)
			response.write "</td>"
			%> </tr>
			<tr bgcolor="#CCFFCC"> 
		    <td>Billing</td>
            <%
			if bolSpreadLow then
			   response.write "<td bgcolor = #FF9999 align=right>"
			else 
			  response.write "<td align=right>"
			end if 
			response.write FormatNumber(curAdjBilling,2)
			response.write "</td>"
			%> </tr>
          
		</table>
		<%
		 If bolSpreadLow then
		  response.write "(Spread is " 
		  response.write FormatCurrency(curReqSpread+curCTHSpread-curAdjSpread,2)
		  response.write " below required)<br>" 
		 end if
		 if intCTHlength > 0 then
		   if intCTHlength = 1 then
		     response.write "<br>This is a CTH calculation for <br>6 months"
		   else
		     response.write "<br>This is a CTH calculation for <br>" & intCTHLength & " months"
		   end if
		   if bolNoFeeCTH	then
		    response.write " with no fee."
		   else
			response.write " with a 1 mo. billing fee." 
		   end if
		 else
		   response.write "<br>This is a straight contract calculation."  
		 end if
		%>
		 <br><br> <input type="button" onClick='CheckData()' value="Continue to Print">
       
	   </div>
      </td>
    <!-- End Row 2 Col 1 --> <!-- Row 2 Col 2 --> <!-- Row 2 Col 3 --> </tr>
  <!-- End Row 2 --> 
  
  <!-- Start Row 3, Col 1 in Table 1 -->
  <tr bgcolor="#FFFFFF"> 
    <td align="center" valign="top" colspan="2"> 
	  <hr>
      <b>Submittal Information</b><br>
	  <form name="f" method="post" action="calcprint.asp">
	  <!-- Set up table and form to input Submittal Information  -->
	  <table border="1"  bgcolor="FFFFFF" cellpadding="0" cellspacing="0">
	    <tr>
		    <td align="right" width="73">Contractor: </td>
		    <td height="20" width="209"> 
              <input tabindex="1" name="ContName" type="text" size="28">
            </td>
		    <td align="right" width="78">Company: </td>
		    <td width="220">
              <input tabindex="7" name="Company" type="text" size="28">
            </td>
		</tr>  
	  <tr>
		    <td align="right" width="73">City/State: </td>
		    <td width="209"> 
              <input tabindex="2" name="CityState" type="text" size="28">
            </td>
		    <td align="right" width="78"> 
              <div align="right">Job Title: </div>
            </td>
		    <td width="220">
              <input tabindex="8" name="JobTitle" type="text" size="28">
            </td>
		</tr>  
		<tr>
		    <td align="right" width="73">Hm Phone: </td>
		    <td width="209"> 
              <input tabindex="3" name="HmPhone" type="text" size="28">
            </td>
		    <td align="right" width="78"> 
              <div align="right">Job Site: </div>
            </td>
		    <td width="220">
              <input tabindex="9" name="JobSite" type="text" size="28">
            </td>
		</tr>  
		<tr>
		    <td align="right" width="73">Wk Phone: </td>
		    <td width="209"> 
              <input tabindex="4" name="WkPhone" type="text" size="28">
            </td>
		    <td rowspan="3" valign="top" align="right" width="78"> 
              <div align="right">Skills: </div>
            </td>
		    <td rowspan="3" valign="top" width="220"> 
              <textarea tabindex="10" name="Skills" rows="4" cols="21"></textarea>
            </td>
		</tr>  
		<tr>
		    <td align="right" width="73">Cont Rep: </td>
		    <td width="209"> 
              <input tabindex="5" name="ContRep" type="text" size="28">
            </td>
		</tr>  
		<tr>
		    <td align="right" width="73">Job Rep: </td>
		    <td width="209"> 
              <input tabindex="6" name="JobRep" type="text" size="28">
            </td>
		</tr>
		</table><br><b>
		Please indicate <u>Yes</u> on the items below that have been thoroughly discussed<br>with the candidate:
		</b><br><br>
		<table border="0" width="600" cellpadding="0" cellspacing="0"> 
		<tr>
	      <td width="80"><Select Name="Q02"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="220">Start date</td>
		  
	      <td width="80"><Select Name="Q01"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="220">How we pay</td>
	    </tr>
	  
	   <tr>
	      <td width="80"><Select Name="Q07"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="220">Rate</td>
	      
		  <td width="80"><Select Name="Q12"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="220">Notice</td>
	    </tr>
		
		<tr>
	      <td width="80"><Select Name="Q08"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="220">W2 or INC</td>
	      
		  <td width="80"><Select Name="Q03"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="220">In-face expenses policy</td>
	    </tr>
		
		<tr>
	      <td width="80"><Select Name="Q09"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="220">Are they committed?</td>
	      
		  <td width="80"><Select Name="Q13"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="220">Advances/advance repay</td>
	    </tr>
		<tr>
	      <td width="80"><Select Name="Q11"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="520" colspan="3">If working, talk with manager</td>
	    </tr>
		<tr>
	      <td width="80"><Select Name="Q10"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="520" colspan="3">Getting to the job / no relocation expenses</td>
	    </tr>
		<tr>
	      <td width="80"><Select Name="Q04"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="520" colspan="3">Our contract (termination clause; non-compete)</td>
	    </tr>
		<tr>
	      <td width="80"><Select Name="Q05"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="520" colspan="3">Background check / drug test (if applicable)</td>
	    </tr>
		<tr>
	      <td width="80"><Select Name="Q06"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="520" colspan="3">Have you checked in the database for a referral fee?</td>
	    </tr>
		<tr>
	      <td width="80"><Select Name="Q14"><Option value="">
                                            <Option value=Y>Yes
                                            <Option value=N>No
                  		</Select></td>
	      <td width="520" colspan="3">IS THERE AN ADDENDUM? (IF SO, PLEASE ATTACH)</td>
	    </tr>
	  </table>

	  <input type="button" onClick='CheckData()' value="Continue to Print">
	  
	  <!-- Pass all the data needed for the compensation information in hidden values--> 
	  <input type="hidden" value=<%=strCalcType%> name="CalcType">
	  <input type="hidden" value=<%=curContRate%> name="ContRate">
	  <input type="hidden" value=<%=curBillRate%> name="BillRate">
	  <input type="hidden" value=<%=curW2RefFee%> name="W2RefFee">
	  <input type="hidden" value=<%=curIncRefFee%> name="IncRefFee">
	  <input type="hidden" value=<%=curPerDiemDly%> name="PerDiemDly">
  	  <input type="hidden" value=<%=curPerDiemHrly%> name="PerDiemHrly">
	  <input type="hidden" value=<%=curOtherAmt%> name="OtherAmt">
	  <input type="hidden" value=<%=strOtherDesc%> name="OtherDesc">
	  <input type="hidden" value=<%=curSubTot1%> name="SubTot1">
	  <input type="hidden" value=<%=curSubTot2%> name="SubTot2">
	  <input type="hidden" value=<%=curTotalCost%> name="TotalCost">
	  <input type="hidden" value=<%=curAdjBilling%> name="AdjBilling">
	  <input type="hidden" value=<%=curContPayout%> name="ContPayout">
	  <input type="hidden" value=<%=curAdjTotalCost%> name="AdjTotalCost">
	  <input type="hidden" value=<%=curAdjSpread%> name="AdjSpread">
	  <input type="hidden" value=<%=curReqSpread%> name="ReqSpread">
  	  <input type="hidden" value=<%=curCTHSpread%> name="CTHSpread">
  	  <input type="hidden" value=<%=intCTHlength%> name="CTHlength">
   	  <input type="hidden" value=<%=bolNoFeeCTH%> name="NoFeeCTH">
   	  <input type="hidden" value=<%=bolStaffFee%> name="StaffFee">
	  <input type="hidden" value=<%=curTaxAmt%> name="TaxAmt">
	  <input type="hidden" value=<%=curAdminAmt%> name="AdminAmt">
	  <input type="hidden" value=<%=curStaffAmt%> name="StaffAmt">
	  <input type="hidden" value=<%=curActStaffFee%> name="ActStaffFee">
	  </form>
    </td>
	<!-- End Row 3, Col 3 in Table 1 -->
  </tr>
  
 </table>
 </td>
 </tr>
 </table>
</body>
</html>

<script language=vbscript>
function CheckData()
	if isMaxLength(document.f.ContName,30) then
		exit function
	elseif isMaxLength(document.f.CityState, 30) then
		exit function
	elseif isMaxLength(document.f.HmPhone, 30) then
		exit function
	elseif isMaxLength(document.f.WkPhone, 30) then
		exit function
	elseif isMaxLength(document.f.ContRep, 30) then
		exit function
	elseif isMaxLength(document.f.JobRep, 30) then
		exit function
	elseif isMaxLength(document.f.Company, 30) then
		exit function
    elseif isMaxLength(document.f.JobTitle, 30) then
		exit function
	elseif isMaxLength(document.f.JobSite, 30) then
		exit function
	elseif isMaxLength(document.f.Skills, 120) then
		exit function		
	end if
	document.f.submit
end function
</script>

