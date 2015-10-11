<%@ LANGUAGE=VBScript %>
<% Option Explicit %>
    <html>
     <head><title>Library Login</title></head>
     <body>
     <center>
	 <br><br>
	  <table width=300 cellpadding=10 cellspacing=0 bgcolor="#eeeeee" border=1>
       <tr>
	     <td align="center">
         <form method="post" action="libreg.asp">
          Username:
          <input name="username" size=20 maxlength=20 value="">
          <br>Password:
          <input name="password" type="password" size=20 maxlength=20 value="">
		  <br><br><input type="submit" value="Login">
	  	  </form>
          </td>
       </tr>
     </table>
	 </center>
     </body>
     </html>