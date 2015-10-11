
<FORM NAME="login" ACTION="loginconfirm.cfm" METHOD=POST onSubmit="return _CF_checklogin(this)" TARGET="_parent">
<div id="loginBox"><br>
   <table border="0" align="left" cellpadding="0" cellspacing="4">
	 <tr><td colspan="2" class="loginText"><%=loginHeading%><br></td></tr>
	 <tr>
	   <td width="76" align="right"><span class="loginText">&nbsp;User Name:</span></td>
	   <td width="88" ><INPUT TYPE="Text" NAME="txtLogin" SIZE="15" MAXLENGTH="20" VALUE=""></td>
	 </tr>
	 <tr>
	   <td width="76" align="right"><span class="loginText">&nbsp;Password:&nbsp;&nbsp; </span></td>
	   <td width="88"><INPUT TYPE="&nbsp;   Password " NAME="txtPassword" SIZE="15" MAXLENGTH="20"></td>
	 </tr>
	  <tr>
	  <td width="76">&nbsp;</td>
	  <td width="88" align="center"><span class="formButton">
	  <input type="image" name="Login" src="images/buttonLogin.jpg" value="Go" title="Login"></span>
	 
	  </td>
     </tr>
	 </table></div>
</form> 