<%@ LANGUAGE=VBScript %>
<% Option Explicit %>
<%

Dim strUserName, strPassWord
strUserName = TRIM( Request( "username" ) )
strPassWord = TRIM( Request( "password" ) )

function ValidLogIn(UserName,PassWord)
'Read in username and password from text file
Dim strFileUserName, strFilePassWord, strSecurityLine
Dim objOpenFile, objFSO, strPath
Const ForReading=1
strPath=Request.QueryString("URL")
strPath=Server.MapPath("sfEmpLog.asp")
set objFSO=Server.CreateObject("Scripting.FileSystemObject")
set objOpenFile=objFSO.OpenTextFile(strPath,ForReading)
strSecurityLine=Trim(objOpenFile.ReadLine)
strFileUserName=Trim(objOpenFile.ReadLine)
strFilePassword=Trim(objOpenFile.ReadLine)
objOpenFile.close
Set objOpenFile=nothing
Set objFSO = nothing

' Check for valid username and password, then write cookies
ValidLogIn = True
if username <> "" and password <> "" then
  if username <> strFileUserName then
    ValidLogIn = false
    response.write "You did not enter a valid user name.<br>"
  end if
  if password <> strFilePassword then
    ValidLogIn = false
	response.write "You did not enter a valid password.<br>"
  end if 
else
  ValidLogIn = False
  response.write "You must log in to access this page.<br>"
end if
end function

if ValidLogIn(strUserName, strPassword) then
    response.cookies("Username")=strUserName
    response.cookies("Password")=strPassWord
	response.cookies("loggedIN")="2222"
	 response.cookies("Username").Expires=Date+1
    response.cookies("Password").Expires=Date+1
	response.cookies("loggedIN").Expires=Date+1
	response.redirect "sfEmployee.asp"
else
  response.write "<br>"
  response.write "Press the back button and enter the username and password."	
end if

%>
