<% page = "calculator.asp" %> 
<% title = "Cost Calculator" %>
<!--#include file="calcsecurity.asp"--> 
<!--#include file="header.asp"-->

<SCRIPT LANGUAGE=VBScript>
<!--#include file="calcvalidinput.vbs"-->
<!--#include file="adovbs.inc"-->

</SCRIPT>
<%
' Read in type of calculation from querystring
Dim strCalcType
strCalcType = request("CalcType")
If len(trim(strCalcType))=0 then strCalcType = "W-2"
%>


<table border="0" cellpadding="0" cellspacing="0" id="mainTable">
  <tr valign="top">
    <td rowspan="2" id="leftMenuCell"><div id="photoBox"><img src="images/calculator.jpg" width="200" height="160" /></div>
      <ul>
        <li><a href="calculator.asp?CalcType=W-2"><span id="arrow">&raquo;</span> - W-2</a></li>
        <li><a href="calculator.asp?CalcType=INC"><span id="arrow">&raquo;</span> - INC</a></li>
        <li><a href="calculator.asp?CalcType=SUB"><span id="arrow">&raquo;</span> - SUB</a></li>
       </ul>
	   </td>
    <td id="contentCell">
     
      <p class="pageHeading">
   <%
	' Display heading based on the type of calculation
	If strCalcType = "SUB" then
	  response.write "&nbsp;(SUB) "
	elseif strCalcType = "INC" then
	  response.write "&nbsp;(INC) "
	elseif strCalcType = "W-2" then 
	  response.write "&nbsp;(W-2) "
    end if
	%>
	Cost Calculator
	   </p>
	  
	  
	  <form name="f" method=POST action="calcresults.asp">


       <table class="calcTable">
	    <tr>
		<td class="calcTable">&nbsp;Contract Type:&nbsp;</td>
		<td class="calcTable">
		<Select Name="cthtype"><Option value=0>Straight Contract Position
                           <Option value=1>CTH-6 months, + 1 month billing
		                   <Option value=3>CTH-3 months, adj spread/no fee
						   <Option value=6>CTH-6 months, adj spread/no fee
						   <Option value=9>CTH-9 months, adj spread/no fee
    	</Select>
		</td>
		</tr>						
		<tr>
		  <td class="calcTable">&nbsp;Contractor Rate:&nbsp;
		  </td>
		  <%
		  response.write "<td class='calcTable'>"
		  if strCalcType="W-2" then
     	    response.write "Base Rate:&nbsp;&nbsp;&nbsp;&nbsp;(or)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total PayOut:<br>"
		    response.write "<input name='ContRate' type='text'size='14'>" 
		    response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		    response.write "<input name='PayOut' type='text' size='14' default='0'>"
		  else
		    response.write "<input name='ContRate' type='text' size='21'>" 
            response.write "<input name='PayOut' type='hidden' value=0>"
		  end if
		  %>
		  </td> 
		</tr>
	    <tr>
		  <td class="calcTable">&nbsp;W2 Referral Fee:&nbsp;</td>
		  <td class="calcTable"><input name="W2RefFee" type="text" size="21"></td>
		</tr>
	    <tr>
		  <td class="calcTable">&nbsp;INC Referral Fee:&nbsp;</td>
		  <td class="calcTable"><input name="IncRefFee" type="text" size="21"></td>
		</tr>
		
		<%
		if strCalcType="W-2" then
		  response.write "<tr >"
		  response.write "<td class='calcTable'>&nbsp;Daily Per Diem:&nbsp;</td>"
		  response.write "<td class='calcTable'><input name='PerDiem' type='text' size='21'></td>"
		  response.write "</tr>"
		else
		  response.write "<input name='PerDiem' type='hidden' value=0>"
		end if
		%>
	    
		<tr>
		<td class="calcTable">&nbsp;Other &nbsp;<Select Name="OtherDesc">
		                        <Option value="">
								<Option value="Escrow">Escrow</Select></td>
		<td class="calcTable"><input name="OtherAmt" type="text" size="21"></td>
	    </tr>
	   
	    	    <tr>
		<td class="calcTable">&nbsp;Staffing Fee %:&nbsp;<br>&nbsp;(i.e. discounts)</td>
		<td class="calcTable">
		<Select Name="ReqStaffFee"><Option value=0>none
		                        <Option value=0.5>0.5% of bill rate
                        		<Option value=1.0>1.0% of bill rate
                                <Option value=1.5>1.5% of bill rate
								<Option value=3.0>3.0% of bill rate
								<Option value=5.0>5.0% of bill rate
						
		</Select>
		</td>
		</tr>						
	
	    <tr>
		  <td class="calcTable">&nbsp;Billing Rate:&nbsp;</td>
		  <td class="calcTable"><input name="BillRate" type="text" size="21"></td>
		</tr>
		<tr>
		<td colspan="2" align="center" class="calcTable"><br>
  <input type="hidden" value=<%=strCalcType%> name="CalcType">
      <input type="reset" value="Reset" name="Reset"> 
      <input type="button" onClick='CheckData()' value="Calculate">

		</td></tr>

 	 </table>
	

</form>
    
	   </td>
  </tr>
  <tr><td><!--#include file="footer.asp"--></td></tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>

<script language=vbscript>
function CheckData()
   	if not isNumber(document.f.ContRate) then
		exit function
  	elseif not isNumber(document.f.Payout) then
		exit function
	elseif not isNumber(document.f.W2RefFee) then
		exit function
	elseif not isNumber(document.f.IncRefFee) then
		exit function
	elseif not isNumber(document.f.PerDiem) then
		exit function
    elseif not isDaily(document.f.PerDiem) then
	  exit function
	elseif not isNumber(document.f.OtherAmt) then
		exit function
	elseif not isNumber(document.f.BillRate) then
		exit function
	elseif not isValidCTH(document.f.CTHtype, document.f.ContRate) then
		exit function
	elseif not isValidStaffFee(document.f.CTHtype, document.f.ReqStaffFee) then
		exit function
	elseif not isEnough(document.f.ContRate, document.f.BillRate, document.f.PayOut, document.f.PerDiem) then
	  exit function
 	elseif isPayOut(document.f.ContRate, document.f.PayOut, document.f.PerDiem) then
	  exit function
	elseif not isAboveMin(document.f.ContRate, document.f.PayOut, document.f.PerDiem) then
	  exit function
   
   	end if
	document.f.submit
end function
</script>

