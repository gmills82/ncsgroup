<%@ LANGUAGE=VBScript %>
<% Option Explicit %>
<%
  response.cookies("Username")=""
  response.cookies("Password")=""
  response.cookies("loggedIN")="0000"
  response.redirect "home.asp"
%>
