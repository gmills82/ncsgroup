<%@ Language=VBSCript %>
<% Option Explicit %>
<SCRIPT LANGUAGE=VBScript>
<!--#include file="calcsecurity.asp"--> 
</SCRIPT>

<%
  'Declare form field variables
  Dim strContName, strCityState
  Dim strHmPhone, strWkPhone
  Dim strContRep, strJobRep
  Dim strCompany, strJobTitle 
  Dim strJobSite, strSkills
  Dim strCalcType, curContRate
  Dim curBillRate, curW2RefFee
  Dim curIncRefFee, curPerDiemDly, curPerDiemHrly
  Dim curOtherAmt, StrOtherDesc, curStaffAmt
  Dim curSubTot1, curSubTot2, curActStaffFee
  Dim curTotalCost, curAdjBilling, bolStaffFee
  Dim curContPayOut, curAdjTotalCost, bolNoFeeCTH
  Dim curAdjSpread, curReqSpread, curCTHSpread
  Dim curTaxAmt, curAdminAmt, intCTHlength
  Dim strQ01, strQ02, strQ03, strQ04, strQ05
  Dim strQ06, strQ07, strQ08, strQ09, strQ10
  Dim strQ11, strQ12, strQ13, strQ14 
 
 'Read in form field variables
  strCalcType=Request("CalcType")
  strContName=Request("ContName")
  strCityState=Request("CityState")
  strHmPhone=Request("HmPhone")
  strWkPhone=Request("WkPhone")
  strContRep=Request("ContRep")
  strJobRep=Request("JobRep")
  strCompany=Request("Company")
  strJobTitle=Request("JobTitle")
  strJobSite=Request("JobSite")
  strSkills=Request("Skills")
  curContRate=Request("ContRate")
  curBillRate=Request("BillRate")
  curW2RefFee=Request("W2RefFee")
  curIncRefFee=Request("IncRefFee")
  curPerDiemDly=Request("PerDiemDly")
  curPerDiemHrly=Request("PerDiemHrly")
  curOtherAmt=Request("OtherAmt")
  strOtherDesc=Request("OtherDesc")
  curSubTot1=Request("SubTot1")
  curSubTot2=Request("SubTot2")
  curTotalCost=Request("TotalCost")
  curAdjBilling=Request("AdjBilling")
  curContPayOut=Request("ContPayOut")
  curAdjTotalCost=Request("AdjTotalCost")
  curAdjSpread=Request("AdjSpread")*1.0
  curReqSpread=Request("ReqSpread")*1.0
  curCTHSpread=Request("CTHSpread")*1.0
  intCTHlength=Request("CTHlength")
  curTaxAmt=Request("TaxAmt")
  curAdminAmt=Request("AdminAmt")
  curStaffAmt=Request("StaffAmt")
  curActStaffFee=Request("ActStaffFee")*1.0
  bolNoFeeCTH=Request("NoFeeCTH")
  bolStaffFee=Request("StaffFee")
  strQ01=Request("Q01")
  strQ02=Request("Q02")
  strQ03=Request("Q03")
  strQ04=Request("Q04")
  strQ05=Request("Q05")
  strQ06=Request("Q06")
  strQ07=Request("Q07")
  strQ08=Request("Q08")
  strQ09=Request("Q09")
  strQ10=Request("Q10")		 
  strQ11=Request("Q11")
  strQ12=Request("Q12")
  strQ13=Request("Q13")
  strQ14=Request("Q14")
%>
  
<html>
<head>
<title>Cost Calculator</title>
<style type="text/css">
<!--
td {  font-family: Arial, Helvetica, Verdana, sans-serif; font-size: 10pt; line-height: 15pt}
-->
</style>
</head>
<body topmargin="0" leftmargin="0">

<H3 align="center"><font face="Arial, Helvetica, Verdana, sans-serif">
<a href="calculator.asp">(<%=strCalcType%>)</a> 
  CANDIDATE SUBMITTAL FORM - <%=Date%></font></H3>
<table width="740" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td align="left" width="100"><i>Contractor: </i></td>
    <td width="270"><%=strContName%></td>
    <td align="left" width="100"><i>Company: </i></td>
    <td width="270"><%=strCompany%></td>
  </tr>
  <tr> 
    <td align="left" width="100"><i>City/State: </i></td>
    <td width="270"><%=strCityState%></td>
    <td align="left" width="100"><i>Job Title: </i></td>
    <td width="270"><%=strJobTitle%></td>
  </tr>
  <tr> 
    <td align="left" width="100"><i>Hm Phone: </i></td>
    <td width="270"><%=strHmPhone%></td>
    <td align="left" width="100"><i>Job Site: </i></td>
    <td width="270"><%=strJobSite%></td>
  </tr>
  <tr> 
    <td align="left" width="100"><i>Wk Phone: </i></td>
    <td width="270"><%=strWkPhone%></td>
    <td align="left" width="100" valign="top" rowspan="3"><i>Skills:</i></td>
    <td width="270" valign="top" rowspan="3"><%=strSkills%></td>
  </tr>
  <tr> 
    <td align="left" width="100"><i>Cont Rep: </i></td>
    <td width="270"><%=strContRep%></td>
  </tr>
  <tr> 
    <td align="left" width="100"><i>Job Rep: </i></td>
    <td width="270"><%=strJobRep%></td>
  </tr>
</table>
<br>
<table border="1" cellpadding="0" cellspacing="0" align="center">
<tr>
<td>
<h4 align="center">COMPENSATION INFORMATION</h4>
      <table width="700" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr>
          <td width="200">Contractor Base Rate</td>
          <td width="70" align="right"><%=FormatNumber(curContRate,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250"><u>Referral Fee(s):</u></td>
          <td width="90">&nbsp;</td>
  </tr>
  <tr>
          <td width="200">Total Referral Fee (W2)</td>
          <td width="70" align="right"><%=FormatNumber(curW2RefFee,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250">&nbsp;</td>
          <td width="90">&nbsp;</td>
  </tr>
  <tr>
          <td width="200">Total Referral Fee (INC)</td>
          <td width="70" align="right"><%=FormatNumber(curIncRefFee,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250"><u>On the Contractor</u></td>
          <td width="90">&nbsp;</td>
  </tr>
  <tr>
          <td width="200" align="center">Sub Total</td>
          <td width="70" align="right"><%=FormatNumber(curSubTot1,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250">&nbsp;</td>
          <td width="90">&nbsp;</td>
  </tr>
  <tr>
          <td width="200">W-2 Required Taxes</td>
          <td width="70" align="right"><%=FormatNumber(curTaxAmt,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250">&nbsp;</td>
          <td width="90">&nbsp;</td>
  </tr>
  <tr>
          <td width="200">Other:</td>
          <td width="70" align="right"><%=FormatNumber(curOtherAmt,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250"><u>On the Job</u></td>
          <td width="90">&nbsp;</td>
  </tr>
  <%
  if strCalcType = "W-2" then
    response.write "<tr>"
    response.write "<td width=200>Per Diem: ("
	response.write FormatCurrency(curPerDiemDly,2)
	response.write " / day)</td>"
	response.write" <td width=70 align=right>" 
	response.write FormatNumber(curPerDiemHrly,2)
	response.write "</td><td width=90>&nbsp;</td>"
	response.write "<td width=250>&nbsp;</td><td width=90>&nbsp;</td></tr>"
  end if
  %>
  <tr>
          <td width="200" align="center">Sub Total</td>
          <td width="70" align="right"><%=FormatNumber(curSubTot2,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250">&nbsp;</td>
          <td width="90">&nbsp;</td>
  </tr>
  <tr>
          <td width="200">Administrative Cost</td>
          <td width="70" align="right"><%=FormatNumber(curAdminAmt,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250">&nbsp;</td>
          <td width="90">&nbsp;</td>
  </tr>
   <tr>
        <%
		 If bolStaffFee then
           response.write "<td width=200>Staffing Fee @ "
           response.write FormatNumber(curActStaffFee,2)
           response.write "%</td>"
           response.write "<td width=70 align=right>"
		   response.write FormatNumber(curStaffAmt,2)
		   response.write "</td>"
		   response.write "<td width=90>&nbsp;</td><td width=250>&nbsp;</td><td width=90>&nbsp;</td>"
         end if
		 %>
  </tr>
  <tr>
          <td width="200" align="center">Sub Total</td>
          <td width="70" align="right"><%=FormatNumber(curTotalCost,2)%></td>
	      <td width="90">&nbsp;</td>
          <td width="250">&nbsp;</td>
          <td width="90">&nbsp;</td>
  </tr>
</table>
<hr width="700">
<table width="740" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr>
          <td width="20">&nbsp;</td>
		  <td width="90">Billing Rate: </td>
          <td width="90"><%=FormatNumber(curBillRate,2)%></td>
          <td width="90">Total Cost: </td>
          <td width="90"><%=FormatNumber(curTotalCost,2)%></td>
          <td width="90">Required Spread: </td>
          <% 
		  if intCTHLength > 1 then
		    response.write "<td width=230>"
		    response.write FormatNumber(curReqSpread,2) & " + "
			response.write FormatNumber(curCTHSpread,2) & " = "
			response.write "<u>"
			response.write FormatNumber(curCTHSpread + curReqSpread,2)
			response.write "</u></td>"
		  else
 		    response.write "<td width=230><u>"
		    response.write FormatNumber(curReqSpread,2) & "</u></td>"	
		  end if
		  %>
  </tr>
</table>
</td>
</tr>
</table>
<br>
<table width="740" border="1" cellpadding="0" cellspacing="0" align="center">
  <tr>
  <td>
  <table width="726" border="0" cellpadding="5" cellspacing="0" align="center" valign="top">
    <tr>
          <td width="196" height="20" colspan="2" ><u>Adjusted Totals:</u></td>
          <td width="530" height="20" colspan="6" align="left" ><font size="1"><b>
		  <%
		   if intCTHlength > 0 then
		     if intCTHlength = 1 then
		       response.write "( This is a CTH calculation for 6 months"
		     else
		       response.write "( This is a CTH calculation for " & intCTHLength & " months"
		     end if
		     if bolNoFeeCTH	then
		       response.write " with no fee."
		     else
			   response.write " with a 1-month billing fee of "
			   response.write FormatCurrency(curAdjBilling * 173,2) 
		     end if
		   else
		     response.write "( This is a straight contract calculation."  
		   end if
		   response.write	" )</b></font></td>"
           %>
	</tr>
    <tr>
          <td width="96">Contractor Pay Out</td>
          <td width="100"><%=FormatNumber(curContPayOut,2)%></td>
          <td width="65">Adj.  Billing</td>
          <td width="106"><%=FormatNumber(curAdjBilling,2)%></td>
          <td width="56">Adj. Cost</td>
          <td width="109"><%=FormatNumber(curAdjTotalCost,2)%></td>
          <td width="77">Adj. Spread</td>
          <%
		  response.write "<td width=117>"
		  response.write FormatNumber(curAdjSpread,2)
		  If (curReqSpread + curCTHSpread - curAdjSpread) > 0.01 then response.write "  ****"
		  response.write "</td>"
	    %>
    </tr>
  </table>
  </td>
  </tr>
  </table>
<br>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="740">
  <tr> 
    <td width="300">&nbsp;</td>
    <td width="440" align="right">________________________________</td>
  </tr>
  <tr> 
    <td width="300"><u>Comments:</u></td>
    <td width="440" align="right">****<font face="Arial, Helvetica, Verdana, sans-serif" size="1">(Approval 
      for less than minimum required spread.)</font></td>
  </tr>
  <tr> 
    <td colspan="2"><font face="Arial, Helvetica, Verdana, sans-serif" size="1"> 
      <br><u>
      Please check below the topics which have been thoroughly discussed with 
      the candidate:</u></font></td>
  </tr>
</table>
<table border="0" width="740" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="241">(1) How we pay: <u><%=strQ01%></u></td>
    <td width="221">(7) Rate: <u><%=strQ07%></u></td>
    <td width="278">(10) Getting to the job/</td>
  </tr>
  <tr>
    <td width="241">(2) Start date: <u><%=strQ02%></u></td>
    <td width="221">(8) W-2 or INC: <u><%=strQ08%></u></td>
    <td width="278"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;no relocation 
      expenses: <u><%=strQ10%></u></td>
  </tr>
  <tr>
    <td height="26" width="243">(3) In-face expenses policy: <u><%=strQ03%></u></td>
    <td height="26" width="223">(9) Are they commited? <u><%=strQ09%></u></td>
    <td height="26" width="280">(11) If working, talk with manager: <u><%=strQ11%></u></td>
  </tr>
</table>
<table border="0" width="740" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="462">(4) Our contract (termination clause; non-compete): <u><%=strQ04%></u></td>
    <td width="278">(12) Notice: <u><%=strQ12%></u></td>
  </tr>
  <tr>
    <td width="462">(5) Background check / drug test (if applicable): <%=strQ05%></td>
    <td width="278">(13) Advances / advance repay: <u><%=strQ13%></u></td>
  </tr>
  <tr>
    <td width="462">(6) Have you checked in the database for a referral fee?: <u><%=strQ06%></u></td>
    <td width="278"></td>
  </tr>
  <tr>
    <td width="462">IS THERE AN ADDENDUM? (IF SO, PLEASE ATTACH) <u><%=strQ14%></u></td>
    <td width="278"></td>
  </tr>
</table>
<br>
</body>
</html>