<%
Dim Username, password
username = TRIM( Request( "username" ) )
password = TRIM( Request( "password" ) )

'Read in username and password from text file
Dim strFileUserName, strFilePassword
Dim objOpenFile, objFSO, strPath
Const ForReading=1
strPath=Request.QueryString("URL")
strPath=Server.MapPath("liblist.txt")
set objFSO=Server.CreateObject("Scripting.FileSystemObject")
set objOpenFile=objFSO.OpenTextFile(strPath,ForReading)
strFileUserName=Trim(objOpenFile.ReadLine)
strFilePassword=Trim(objOpenFile.ReadLine)
objOpenFile.close
Set objOpenFile=nothing
Set objFSO = nothing
' Check for valid username and password, 
if username <> strFileUserName or password <> strFilePassword then response.redirect "liblogin.asp" 
%>