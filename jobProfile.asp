<% page = "jobProfile.asp" %> 
<% title = "Job Description" %> 
<% action = "LOGIN" %>
<% loginHeading = "  - Manage Your Profile" %>

<!--#include file="header.asp"-->
<table border="0" cellpadding="0" cellspacing="0" id="mainTable">
  <tr valign="top">
    <td rowspan="2" id="leftMenuCell">
	   <div id="photoBox"><img src="images/52.jpg" width="200" height="160" /></div>
       <ul>
	    <li><a href="consultants.asp"><span id="arrow">&raquo;</span> Consultants Overview</a></li>
        <li><a href="jobSearch.asp"><span id="arrow">&raquo;</span> - Job Search</a></li>
        <li><a href="registerApply.asp"><span id="arrow">&raquo;</span> - Register/Apply</a></li>
        <li><a href="careerResources.asp"><span id="arrow">&raquo;</span> - Career Resources</a></li>
       </ul>
		<!--#include file="login.asp"-->
	</td>	  
    
	<td id="contentCell">
	  <div id="breadCrumbs"><a href="home.asp"> Home&raquo;</a>
       <a href="consultants.asp"> Consultant Overview&raquo;</a>
	   <a href="jobSearch.asp"> Job Search&raquo;</a>
	    Job Description</div>
         
       <p class="pageHeading">Job Description</p>
		   <span class="categoryHeading">Title:</span>&nbsp;&nbsp;&nbsp;<span class="bold"> Sr. Applications Developer</span>
	   <br><span class="categoryHeading">Area:</span>&nbsp;&nbsp; Las Vegas, NV
	   <br><span class="categoryHeading">Date:</span>&nbsp;&nbsp; June 8, 2004
       <br><span class="categoryHeading">Job#:</span>&nbsp;&nbsp; 000000
	   <br><br><span class="categoryHeading">Requirements:</span><br><br>
Scientific Systems, a division in the Information Systems department at our client, is currently recruiting for a contractor
position.

<p>We are seeking a Sr. Applications Developer with the following skills:<br>
ColdFusion MX<br>
Flash MX<br>
Javascript<br>
XML<br>
Oracle<br>
Data Management experience<br>
Photoshop/graphical programs<br>
GIS 
</p>
...................................................................................................<br><br>
<span class="subHeading">Apply for this Position</span><br><br>
<a href="registerApply.asp">Click here</a> if you are applying for the first time.<br><br>
If you have previously SUBMITTED YOUR RESUME, input your user name and password and click login.
<% loginHeading = " - LOGIN TO APPLY NOW" %>
<!--#include file="login.asp"-->
  
</td>
  </tr>
<tr><td><!--#include file="footerURL.asp"--></td></tr>
  
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
