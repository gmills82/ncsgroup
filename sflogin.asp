<%@ LANGUAGE=VBScript %>

<% page = "sfLogin.asp" %> 
<% title = "Log In" %> 
<% Option Explicit %>
<!--#include file="sfEmpSecurity.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>The NCS Group - <%=title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<script language="JavaScript1.2" src="scripts.js"></script>
<link href="webStyles.css" rel="stylesheet" type="text/css" />
 <script language="JavaScript">
    function setFocus()
    {
      frmLogin.name.focus();
    }
  </script>
</head>
<body onLoad="MM_preloadImages('images/buttonBackgroundFlip.jpg'); setFocus();">
<a name="top" id="top"></a><br>
<div id="logoBox"><img src="images/motto.gif" alt="1.800.854.0581 --- 1.800.218.1164 --- Utilizing Technology to Drive Business Results" name="mottoPhone" width="306" height="56" align="right" id="mottoPhone" /><img src="images/logo.gif" alt="The NCS Group" width="206" height="76" border="0" /></div>
<div id="topMenuBox">
  <div id="topLeftLinks"> <a href="contactUs.asp">Contact Us</a> <a href="consultants.asp">Career Center</a>


</div>

  <div id="borderedLinks"> <a href="home.asp">Home</a> <a href="company.asp">Company</a> <a href="clients.asp">Clients</a> <a href="consultants.asp">Consultants</a> <a href="services.asp">Services</a> </div>
</div>



<table border="0" cellpadding="0" cellspacing="0" id="mainTable">
  <tr valign="top">
    <td rowspan="2" id="leftMenuCell"><div id="photoBox"><img src="images/54.jpg" width="220" height="308" /></div>
   
    </td>
    <td id="contentCell" align="center">
	<p>&nbsp;</p> <p>&nbsp;</p>
  
 <form id="frmLogin" method="post" action="sfEmpReg.asp">
          Username:
          <input id="name" name="username" size=20 maxlength=20 value="">
          <br>Password:
          <input name="password" type="password" size=20 maxlength=20 value="">
		  <br><br><input type="submit" value="Login">
	  	  </form>
		
		<p>&nbsp;</p><p>&nbsp;</p>
		 </td>
  </tr>
  <tr><td><!--#include file="footer.asp"--></td></tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
