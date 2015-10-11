<% 
if session("refer0") = "" and session("refer1") = "" then
Session("refer0") = Request.ServerVariables("HTTP_REFERER")
Session("refer1") = Request("refid")
end if

if instr(request.querystring,"paybills") > 0 then 
Response.redirect("http://www.352media.com/paybills.asp")
end if

if instr(request.servervariables("SERVER_NAME"),"silverscape") > 0 then %>
<html>
<head>
<title>Welcome to 352 Media Group</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/352styles.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#006699" text="#FFFFFF" link="#99CC33" vlink="#99CC33" alink="#99CC33" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="100%" border="0" cellpadding="10" cellspacing="0">
  <tr valign="middle"> 
    <td width="50%" align="right">
<p class="header_sm"><font color="#FFFFFF">Silverscape Technologies is now</font></p>
      </td>
    <td width="50%" align="left"> <p><font color="#FFFFFF"><a href="http://www.352media.com"><img src="/images/logo291.gif" width="218" height="291" border="0"></a></font></p>
      <p><font color="#FFFFFF">Please visit our home page,<br>
        at <a href="http://www.352media.com">www.352media.com</a>.</font></p>
      </td>
  </tr>
</table>
</body>
</html>
<% else %>
<html>
<head>
<title>352 Media Group - Page not found</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/352styles.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#006699" text="#FFFFFF" link="#99CC33" vlink="#99CC33" alink="#99CC33" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="100%" border="0" cellpadding="10" cellspacing="0">
  <tr valign="middle"> 
    <td width="50%" align="right"><a href="http://www.352media.com"><img src="/images/logo291.gif" width="218" height="291" border="0"></a></td>
    <td width="50%" align="left"> 
      <p class="header_sm"><font color="#FFFFFF">Sorry, 
        the page or file <br>
	you are looking for<br>
        is under construction.</font></p>
      <p><font color="#FFFFFF">Please visit our home page,<br>
        at <a href="http://www.352media.com">www.352media.com</a>.</font></p></td>
  </tr>
</table>
</body>
</html>
<% end if%>