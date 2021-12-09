<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Agenda_Information.asp' is free software; 
' you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
' If the program (this page) is interactive, make it output a short notice 
' like this :
' 	" Copyright (C) 2001-2002  OverApps & contributors "
' at the bottom of the page with an active link from the name "OverApps" to 
' the Address http://www.overapps.com where the user (netsurfer) could find the 
' appropriate information about this license. 
'
'-----------------------------------------------------------------------------
%>

<% 	Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Doesn't Work with PWS ????
%>

<%
' ------------------------------------------------------------
' Name			: __Agenda_Information.asp
' Path   		 : /
' Version 		: 1.15.0
' description 	: Information about a Meeting
' by	 		: Pierre Rouarch, Dania Tcherkezoff
' Company 		: OverApps
' Date			: December 10, 2001
'
' 
' ------------------------------------------------------------

Dim myPage
myPage = "__Agenda_Information.asp"
Dim myPage_Application
myPage_Application="Agenda"


	
%>


<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' INCLUDES 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<!-- #include file="_INCLUDE/Global_Parameters.asp" -->


<!-- #include file="_INCLUDE/DB_Environment.asp" -->

<!-- #include file="_INCLUDE/Environment_Tools.asp" -->

<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CHECK IF THE USER CAN ENTER IN THIS APPLICATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

myApplication_Public_Type_ID = Get_Application_Public_Type_ID(myPage_Application)
if myApplication_Public_type_ID<myUser_type_ID then
	Response.redirect("__Quit.asp")
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET APPLICATION TITLE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myApplication_Title = Get_Application_Title(myPage_Application)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VARIABLES
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim mySQL_Select_tb_Meetings, mySet_tb_Meetings, mySQL_Select_tb_Meetings_Members, mySet_tb_Meetings_Members, mySQL_Select_tb_Phases, mySet_tb_Phases

Dim mySQL_Select_tb_Projects, mySet_tb_Projects

Dim  myMeeting_Site_ID, myMeeting_Member_ID, myMeeting_Project_ID, myMeeting_ID, myMeeting_Title,  myProject_Name,  myMeeting_Date_Beginning, myMeeting_Place, myMeeting_Hour, myMeeting_Minute, myMeeting_Length, myMeeting_Length_In_Minutes,  myMeeting_Agenda,  myMeeting_Comments, myMeeting_Author_Update,  myMeeting_Date_Update, myPhase_Name, myFlag


Dim myParticipants_Emails

Dim myProjects_Public_Type_ID
Dim myMembers_Public_Type_ID, myLocation


myProjects_Public_Type_ID = Get_Application_Public_Type_ID("Projects")
myMembers_Public_Type_ID = Get_Application_Public_Type_ID("Members")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Parameters
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Request Meeting ID
myMeeting_ID = Request.QueryString("Meeting_ID")

' if nothing go back 
if len(myMeeting_ID) = 0 then
	Response.Redirect("__Agenda_Day.asp")
end if

myLocation = request.querystring("location")

%>

<html>

<head>
<title><%=mySite_Name%> - Agenda Information</title>
</head>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>" marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">

<%
' TOP
%> 

<!-- #include file="_borders/Top.asp" --> 

<%
'CENTER
%> 

<TABLE WIDTH="<%=myGlobal_Width%>" BGColor=<%=myBorderColor%> BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> 

<%
' CENTER LEFT
%> 

<TD WIDTH="<%=myLeft_Width%>"> 
<!-- #include file="_borders/Left.asp" --> 
</TD>

<%
' CENTER APPLICATION
%> 


<%

' DB Connection	
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

' SELECT AND AVOID others Sites Members or others Members in the future multisite version

mySQL_Select_tb_Meetings = "SELECT * FROM tb_Meetings  WHERE Meeting_ID="&myMeeting_ID&" AND (Site_ID="&mySite_ID&" OR Member_ID="&myUser_ID&")"

set mySet_tb_Meetings = myConnection.Execute(mySQL_Select_tb_Meetings)

' Get Values
if not mySet_tb_Meetings.eof then

	myMeeting_Site_ID = mySet_tb_Meetings("Member_ID")
	myMeeting_Member_ID = mySet_tb_Meetings("Member_ID") ' creator
	myMeeting_Project_ID = mySet_tb_Meetings("Project_ID")
	myMeeting_Title	= mySet_tb_Meetings("Meeting_Title")
	myMeeting_Date_Beginning = mySet_tb_Meetings("Meeting_Date_Beginning")
	myMeeting_Hour = mySet_tb_Meetings("Meeting_Hour")
	myMeeting_Minute = mySet_tb_Meetings("Meeting_Minute")
	myMeeting_Length = mySet_tb_Meetings("Meeting_Length")
	myMeeting_Length_In_Minutes = mySet_tb_Meetings("Meeting_Length_In_Minutes")
	myMeeting_Place	= mySet_tb_Meetings("Meeting_Place")

	myMeeting_Agenda = mySet_tb_Meetings("Meeting_Agenda")
	if len(myMeeting_Agenda) > 0 then
		myMeeting_Agenda= Replace(myMeeting_Agenda,vbCrLf,"<br>")
	end if

	myMeeting_Comments = mySet_tb_Meetings("Meeting_Comments")
	if len(myMeeting_Comments) > 0 then
		myMeeting_Comments= Replace(myMeeting_Comments,vbCrLf,"<br>")
	end if

	myMeeting_Author_Update	= mySet_tb_Meetings("Meeting_Author_Update")
	myMeeting_Date_Update = mySet_tb_Meetings("Meeting_Date_Update")

	' Search Project Title
	mySQL_Select_tb_Projects = "SELECT Project_Name FROM tb_Projects WHERE Project_ID = "&myMeeting_Project_ID
	set mySet_tb_Projects  = myConnection.Execute(mySQL_Select_tb_Projects)
	if not mySet_tb_Projects.eof then
		myProject_Name = mySet_tb_Projects("Project_Name")
	end if

else
	' Error Meeting Go back to Agenda Day
	' Close Recordset
	mySet_tb_Meetings.close
	Set mySet_tb_Meetings=nothing
	' Close Connection 
	myConnection.Close
	set myConnection = Nothing
	Response.Redirect("__Agenda_Day.asp")
end if

' Close Recordset
mySet_tb_Meetings.close
Set mySet_tb_Meetings=nothing
' Close Connection 
myConnection.Close
set myConnection = Nothing




%> 
<td valign="top" align="left" bgcolor="<%=myBGColor%>" WIDTH="<%=myApplication_Width%>"> 

<table border="0" bgcolor="<%=myBGColor%>" cellpadding="5" cellspacing="1"> 

<TR> 
<TD colspan=2> 
<%
myBox_title=myApplication_Title
%>
<!-- #include file="__Agenda_Box.asp" --> 
</td>
</tr>


<%
' TITLE
%>


<tr> 
<td align="center" valign="top" colspan="2" bgcolor="<%=myApplicationColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR="<%=myApplicationTextColor%>"><B><%=myMessage_Information%></B></FONT>
</td>
</tr> 

<%
' Meeting Title
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Title%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><strong><%=myMeeting_Title%></strong></font>
</td>
</tr>

<%
' Date
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Date%></font></b>
</td>
<td valign="top" align="left" >
<font face="Arial, Helvetica, sans-serif" size="2"> 
<% If myDate_Format =1 Then%>
 &nbsp;<%=MonthName(Month(myMeeting_Date_Beginning))%>,
<%end if%>  

<%=Day(myMeeting_Date_Beginning)%>,

<% If myDate_Format <> 1 Then%>
&nbsp;<%=MonthName(Month(myMeeting_Date_Beginning))%>,
<%end if%>  

&nbsp;<%=Year(myMeeting_Date_Beginning)%></font>
</td>
</tr>

<%
'  Hour
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Hour%></font></b></td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2">
<%If myMeeting_Hour > 12 AND myDate_Format =1 Then%>
<%=myMeeting_Hour - 12 %>
<%else%>
<%=myMeeting_Hour %>
<%end if%>

&nbsp;h&nbsp;<%if myMeeting_Minute=0 then%>0<%end if%><%=myMeeting_Minute%>
<%If myMeeting_Hour > 11 AND myDate_Format =1 Then%>
P.M
<%end if%>
<%If myMeeting_Hour < 12 AND myDate_Format =1 Then%>
A.M
<%end if%>


</font>
</td>
</tr> 

<%
' Length
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Length%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myMeeting_Length%>&nbsp;h&nbsp;<%if myMeeting_Length_In_Minutes=0 then%>0<%end if%><%=myMeeting_Length_In_Minutes%></font>
</td>
</tr> 

<%
' Project
%>

<%

if myUser_type_ID<=myProjects_Public_type_ID then
%>
	<tr>
	<td align="right" valign="top"  bgcolor="<%=myBorderColor%>">
	<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Project%></font></b>
	</td>
	<td valign="top" align="left">
	<font face="Arial, Helvetica, sans-serif" size="2"><%=myProject_Name%></font>
	</td>
	</tr> 
<%end if%> 


<%
' Place
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Place%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myMeeting_Place%></font>
</td>
</tr>


<%
' Agenda
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Meeting_Agenda%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myMeeting_Agenda%> </font>
</td>
</tr> 


<%
' Participants
%>


<%

' DB Connection	
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

		
' Participants List
mySQL_Select_tb_Meetings_Members   = "SELECT *  FROM tb_Sites_Members INNER JOIN tb_Meetings_Members ON tb_Sites_Members.Member_ID = tb_Meetings_Members.Member_ID WHERE tb_Meetings_Members.Site_ID="&session("Site_ID")&" AND  tb_Sites_Members.Site_ID="&session("Site_ID")&" AND tb_Meetings_Members.Meeting_Role_ID=2 AND  tb_Meetings_Members.Meeting_ID = " & myMeeting_ID&" ORDER BY Member_Pseudo"

set mySet_tb_Meetings_Members = myConnection.Execute(mySQL_Select_tb_Meetings_Members)

if not mySet_tb_Meetings_Members.eof then 
%> 
	<tr>
	<td align="right" valign="top" bgcolor="<%=myBorderColor%>"> 
	<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Participants%></font></b><br> 
	<% 
	myParticipants_Emails=""
	do while not mySet_tb_Meetings_Members.eof
		if len(mySet_tb_Meetings_Members("Member_Email"))>0 then 
			myParticipants_Emails = myParticipants_Emails&mySet_tb_Meetings_Members("Member_Email")&";"
		end if
		mySet_tb_Meetings_Members.movenext
	loop
	%> 
	<a href="mailto:<%=myParticipants_Emails%>"><img valign="top" border="0"  alt="<%=myMessage_Participants%>" src="Images/OverApps-mail2.gif" WIDTH="32" HEIGHT="32"></a>
	</td>
	<td valign="top" align="left"> 
	<% 
	mySet_tb_Meetings_Members.movefirst  
	do while not mySet_tb_Meetings_Members.eof 
		if len(mySet_tb_Meetings_Members("Member_Email"))>0 then 
			%> 
			<a href="mailto:<%=mySet_tb_Meetings_Members("Member_Email")%>"><img src="Images/OverApps-mail1.gif" alt="<%=mySet_tb_Meetings_Members("Member_Pseudo")%>" align="absmiddle" border="0" WIDTH="22" HEIGHT="22"></a>
			&nbsp; 
			<% 
			if myUser_type_ID<=myMembers_Public_type_ID then
				%>
				<A HREF="__Site_Member_Information.asp?Member_ID=<%=mySet_tb_Meetings_Members("Member_ID")%>&Site_ID=<%=myMeeting_Site_ID%>"><Font face="Arial, Helvetica, sans-serif" size="2"><%=mySet_tb_Meetings_Members("Member_Pseudo")%></font></a>
				<%
			else
				%>
				<font face="Arial, Helvetica, sans-serif" size="2"><%=mySet_tb_Meetings_Members("Member_Pseudo")%></font>
				<%
			end if
			%>
			<br> 
			<% 
		end if
	mySet_tb_Meetings_Members.movenext 
	loop 
	%>
	</td>
	</tr>
<% 
end if 

' Close Recordset
mySet_tb_Meetings_Members.close
Set mySet_tb_Meetings_Members=Nothing
' Close Connection
myConnection.Close
set myConnection = Nothing
%> 

<%
' Comments
%>

<tr> 
<td align="right" valign="top"  bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Comments%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myMeeting_Comments%></font>
</td>
</tr>


<%
' Date Author
%>

<tr> 
<td align="center"  colspan="2" bgcolor="<%=myApplicationColor%>">&nbsp; 
<b><font face="Arial, Helvetica, sans-serif" size="1" color="<%=myApplicationTextColor%>"> 
<% = myDate_Display(myMeeting_Date_Update,2) %>&nbsp;--&nbsp;<% = myMeeting_Author_Update %></font></b> 
</td>
</tr>

</table>

<%
' NAVIGATION-ADMINISTRATION
%>

<% if myMeeting_Member_ID=MyUser_ID or myUser_Type_ID=1 then%>
	<table border="0" width="90%" cellpadding="3" cellspacing="0"> 
	<tr>
	<td width="1%">&nbsp;
	</td>
	<td>
	<a href="__Agenda_Modification.asp?Meeting_ID=<%=myMeeting_ID%>&location=<%= myLocation %>"><font face="Arial, Helvetica, sans-serif" size="2"><%=myMessage_Modify%></font></a> , <a href="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Agenda_Modification.asp?Meeting_ID=<%=myMeeting_ID%>&amp;Date_Agenda=<%=myMeeting_Date_Beginning%>&amp;action=Delete&location=<%= myLocation %>';"><font face="Arial, Helvetica, sans-serif" size="2"><%=myMessage_Delete%></font></a>
	</td>
	</tr>
	</table>
<%end if%>

<%
' /NAVIGATION-ADMINISTRATION
%>


</td>
</TR>
</TABLE>

<%
' /CENTER 
%> 

<%
' DOWN
%> 


<!-- #include file="_borders/Down.asp" --> 


<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.							'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright												'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 

</body>
</html>

<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>