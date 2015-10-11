<%
Dim Username, password
username = TRIM( Request( "username" ) )
password = TRIM( Request( "password" ) )

'Read in username and password from text file
Dim strFileUserName, strFilePassword, strSecurityLine
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
' Check for valid username and password, 
if username <> strFileUserName or password <> strFilePassword then response.redirect "sfEmpLogin.asp" 
%>