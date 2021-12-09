<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program is free software; you can redistribute it and/or modify
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
'-----------------------------------------------------------------------------
%>

<% 	Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Doesn''t work with PWS ???
%>



<%
' ------------------------------------------------------------
' Name 		: __Agenda_Week.asp
' Path    : /
' Description 	: Agenda by Week
' by 		: Pierre Rouarch	
'contributor  :Dania Tcherkezoff
' Company 	: OverApps
' Date	: December,10, 2001 
'Versions 1.15.0
' ------------------------------------------------------------

Dim myPage
myPage = "__Agenda_Week.asp"

	
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
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET APPLICATION TITLE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myApplication_Title = Get_Application_Title(myPage_Application)


Dim  myMeeting_Member_ID, myMeeting_Site_ID, myMeeting_Title, myMeeting_Date_Beginning, myMeeting_ID, myMeeting_Hour, myMeeting_Minute, myMeeting_Length, myMeeting_Length_In_Minutes, myPrive, myLM, myLH, indice

Dim mySQL, mySQL_Select_tb_Meetings, mySet_tb_Meetings, mySQL_Select_tb_Meetings_Members, mySet_tb_Meetings_Members

Dim myMember_ID,  myMember_Pseudo, myMember_Participant_ID, myMember_Participant_Pseudo, myMeeting_Site_Name




Dim myStrDate_In_Week
Dim myIntDate_In_Week
Dim  myDay_In_Week



Dim myProjects_Public_Type_ID
Dim myMembers_Public_Type_ID

myProjects_Public_Type_ID = Get_Application_Public_Type_ID("Projects")
myMembers_Public_Type_ID = Get_Application_Public_Type_ID("Members")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET Parameters
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


mytype=1 ' WEEK Presentation


' Get Date Agenda
myStrDate_Agenda=request("Date_Agenda")
if Not isDate(myStrDate_Agenda) then 
		myStrDate_Agenda = FormatDateTime(now(),2)
end if
	
' get Month
myMonth_Agenda= Month(myStrDate_Agenda)
if Len(myMonth_Agenda)=0 then
	myMonth_Agenda=Month(now())
end if

' Get Day
myDay_Agenda= Day(myStrDate_Agenda)
if Len(myDay_Agenda)=0 then
	myDay_Agenda=Day(now())
end if

' Get Year
myYear_Agenda= Year(myStrDate_Agenda)
if Len(myYear_Agenda)=0 then
	myYear_Agenda=Year(now())	
end if

' In Integer Format -  The output is the short Country Date, But when you put this Int Value in the DB there is a confusion risk
myIntDate_Agenda=DateSerial(Year(myStrDate_Agenda),Month(myStrDate_Agenda),Day(myStrDate_Agenda))

' In General Format YYYY/MM/DD
myStrDate_Agenda=Year(myStrDate_Agenda)&"/"&Month(myStrDate_Agenda)&"/"&Day(myStrDate_Agenda)



myIntFirst_Day_Of_Week = DatePart("w",myStrDate_Agenda) - 1
	

' Maybe Not Used
	Function ConvertDate(myDateToConvert)
		ConvertDate=Month(myDateToConvert)& "/" &Day(myDateToConvert) & "/" &Year(myDateToConvert)
		ConvertDate=CDate(ConvertDate)	
	End Function
	
' Force Member to User
myMember_ID=myUser_ID
myMember_Pseudo=myUser_Pseudo

%>

<HTML>
<HEAD>
<TITLE><%=mySite_Name%> Agenda <%=myMessage_Week%></TITLE>
</HEAD>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>" marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%>

<!-- #include file="_borders/Top.asp" -->

<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR=<%=myBorderColor%> BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
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

<td WIDTH="<%=myApplication_Width%>" BGColor="<%=myBGColor%>">

<table BGColor=<%=myBGColor%> border=0 cellpadding=3 cellspacing=1>
<TR> 
<TD> 

<%
myBox_title=myApplication_Title
%>

<!-- #include file="__Agenda_Box.asp" --> 

</td>
</tr>

<%
' WEEK Presentation
%>

<%
' WEEK Navigation
%>

<tr>
<td bgcolor="<%=myApplicationColor%>">
<table border=0 cellpadding="5" cellspacing=0 > 
<tr bgcolor="<%=myApplicationColor%>" >
<td>
<a href="__Agenda_Week.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda,myDay_Agenda-7)%>"><img border=0 height=11 src="Images/OverApps-left_small.gif" width=11></a><font face=Arial,Helvetica size=+1 Color="<%=myApplicationTextColor%>"><b>
<%
If myDate_Format = 1 Then 
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 1 Then response.write   myMessage_January
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 2 Then response.write   myMessage_February
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 3 Then response.write   myMessage_March
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 4 Then response.write   myMessage_April
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 5 Then response.write   myMessage_May
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 6 Then response.write   myMessage_June
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 7 Then response.write   myMessage_July
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 8 Then response.write   myMessage_August
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 9 Then response.write   myMessage_September
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 10 Then response.write   myMessage_October
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 11 Then response.write   myMessage_November
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 12 Then response.write   myMessage_December
 response.write ", " & Day(myIntDate_Agenda-myIntFirst_Day_Of_Week)
 
 else 
	
 response.write Day(myIntDate_Agenda-myIntFirst_Day_Of_Week)& " "
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 1 Then response.write   myMessage_January
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 2 Then response.write   myMessage_February
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 3 Then response.write   myMessage_March
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 4 Then response.write   myMessage_April
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 5 Then response.write   myMessage_May
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 6 Then response.write   myMessage_June
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 7 Then response.write   myMessage_July
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 8 Then response.write   myMessage_August
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 9 Then response.write   myMessage_September
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 10 Then response.write   myMessage_October
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 11 Then response.write   myMessage_November
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week) = 12 Then response.write   myMessage_December
end if		
%>&nbsp;<%=Year(myIntDate_Agenda-myIntFirst_Day_Of_Week)%>
<font size="+1" face="Courier New, Courier, mono">--></font>
<%
If myDate_Format = 1 Then 
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 1 Then response.write   myMessage_January
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 2 Then response.write   myMessage_February
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 3 Then response.write   myMessage_March
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 4 Then response.write   myMessage_April
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 5 Then response.write   myMessage_May
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 6 Then response.write   myMessage_June
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 7 Then response.write   myMessage_July
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 8 Then response.write   myMessage_August
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 9 Then response.write   myMessage_September
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 10 Then response.write   myMessage_October
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 11 Then response.write   myMessage_November
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 12 Then response.write   myMessage_December
 response.write ", " & Day(myIntDate_Agenda-myIntFirst_Day_Of_Week+6)
 
 else 
	
 response.write Day(myIntDate_Agenda-myIntFirst_Day_Of_Week+6)& ", "
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 1 Then response.write   myMessage_January
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 2 Then response.write   myMessage_February
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 3 Then response.write   myMessage_March
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 4 Then response.write   myMessage_April
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 5 Then response.write   myMessage_May
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 6 Then response.write   myMessage_June
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 7 Then response.write   myMessage_July
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 8 Then response.write   myMessage_August
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 9 Then response.write   myMessage_September
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 10 Then response.write   myMessage_October
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 11 Then response.write   myMessage_November
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+6) = 12 Then response.write   myMessage_December
end if		
%>&nbsp;<%=Year(myIntDate_Agenda-myIntFirst_Day_Of_Week+6)%>
&nbsp;(<%=MyMessage_Week & "&nbsp;" &  DatePart("ww",myIntDate_Agenda-myIntFirst_Day_Of_Week)%>)



<a href="__Agenda_Week.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda,myDay_Agenda+7)%>"><img border=0 height=11 src="Images/OverApps-right_small.gif" width=11></a>
</td>
<td align=right>&nbsp;
 
</td>
</tr>
</table>
</td>
</tr>


<%
' OutPut Pseudo and Add Link
%> 

<TR>
<TD BGCOLOR=<%=myBGColor%> VALIGN=top COLSPAN="2">
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="<%=myApplication_Width%>"> 
<TR>
<TD ALIGN=left WIDTH="67%">
<FONT FACE=Arial,Helvetica SIZE=-1 color="<%=myBGTextColor%>"><B><%=myUser_Pseudo%></B></FONT>
</TD>
<TD WIDTH="33%" ALIGN="RIGHT">
<A HREF="__Agenda_Modification.asp?Date_Agenda=<%=myStrDate_Agenda%>"><FONT FACE=Arial,Helvetica SIZE=+1 Color="<%=myBGTextColor%>"><B><%=myMessage_Add%></B></FONT></A> 
</TD>
</TR>
</TABLE>
</TD>
</TR>


<%
' DAYS
%> 

<%
 i=0 

 Do while i < (7 )  
	myintDate_In_Week=myIntDate_Agenda-myIntFirst_Day_Of_Week+i
	myStrDate_In_Week= myDate_Construct(Year(myIntDate_In_Week),Month(myIntDate_In_Week),Day(myIntDate_In_Week),0,0,0)
	myStrDate_In_Week = left(myStrDate_In_Week,10)


	if myStrDate_In_Week = myStrDate_Agenda then
		%> 
		<tr bgcolor=MistyRose>
		<%
	else
		%>
		<tr bgcolor="#DCDCDC">
		<%
	end if
	%>
	<td>
	<Table WIDTH="100%">
	<Tr>
	<td> 
	<font face=Arial,Helvetica size=-1  Color=black><b>
	<%

If DatePart("w",myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 1 Then response.write myMessage_Sunday
If DatePart("w",myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 2 Then response.write myMessage_Monday
If DatePart("w",myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 3 Then response.write myMessage_Tuesday
If DatePart("w",myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 4 Then response.write myMessage_Wednesday
If DatePart("w",myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 5 Then response.write myMessage_Thursday
If DatePart("w",myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 6 Then response.write myMessage_Friday
If DatePart("w",myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 7 Then response.write myMessage_Saturday


	
	%>
	
	, <a href="__Agenda_Day.asp?Date_Agenda=<%=FormatDateTime(myIntDate_Agenda-myIntFirst_Day_Of_Week+i, 2)%>"><font color="#0000ff">
	
&nbsp;<%
If myDate_Format = 1 Then 
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 1 Then response.write   myMessage_January
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 2 Then response.write   myMessage_February
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 3 Then response.write   myMessage_March
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 4 Then response.write   myMessage_April
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 5 Then response.write   myMessage_May
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 6 Then response.write   myMessage_June
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 7 Then response.write   myMessage_July
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 8 Then response.write   myMessage_August
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 9 Then response.write   myMessage_September
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 10 Then response.write   myMessage_October
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 11 Then response.write   myMessage_November
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 12 Then response.write   myMessage_December
 response.write ", " & Day(myIntDate_Agenda-myIntFirst_Day_Of_Week+i)
 
 else 
	
 response.write Day(myIntDate_Agenda-myIntFirst_Day_Of_Week+i)& " "
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 1 Then response.write   myMessage_January
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 2 Then response.write   myMessage_February
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 3 Then response.write   myMessage_March
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 4 Then response.write   myMessage_April
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 5 Then response.write   myMessage_May
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 6 Then response.write   myMessage_June
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 7 Then response.write   myMessage_July
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 8 Then response.write   myMessage_August
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 9 Then response.write   myMessage_September
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 10 Then response.write   myMessage_October
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 11 Then response.write   myMessage_November
 If Month(myIntDate_Agenda-myIntFirst_Day_Of_Week+i) = 12 Then response.write   myMessage_December
end if		
%>&nbsp;<%=Year(myIntDate_Agenda-myIntFirst_Day_Of_Week+i)%>
	
	
	
	</font></a></b></font>
	</td>
	<td ALIGN="RIGHT">
	<font face=Arial,Helvetica size=-2><a href="__Agenda_Modification.asp?Date_Agenda=<%= Year(FormatDateTime(myIntDate_Agenda-myIntFirst_Day_Of_Week+i, 2))%>/<%=Month(FormatDateTime(myIntDate_Agenda-myIntFirst_Day_Of_Week+i, 2))%>/<%=Day(FormatDateTime(myIntDate_Agenda-myIntFirst_Day_Of_Week+i, 2))%>" align="right"><font color="#0000ff"><%=myMessage_Add%></font></a></font>
	</td>
	</tr>
	</table>
	</td>
	</tr>
	<%
	' MEETINGS PLANNING
	%>
	<%
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String
	    
	' SELECT only where i'm the author or a participant
	mySQL_Select_tb_Meetings = "SELECT DISTINCT tb_Meetings.Site_ID,  tb_Meetings.Member_ID,  tb_Meetings.Project_ID,  tb_Meetings.Member_ID, tb_Meetings.Phase_ID,  tb_Meetings.Meeting_ID, tb_Meetings.Meeting_Title, tb_Meetings.Meeting_Date_Beginning,  tb_Meetings.Meeting_Hour, tb_Meetings.Meeting_Minute,  tb_Meetings.Meeting_Length, tb_Meetings.Meeting_Length_In_Minutes  FROM tb_Meetings INNER JOIN tb_Meetings_Members ON tb_Meetings_Members.Meeting_ID=tb_Meetings.Meeting_ID WHERE ( Meeting_Date_Beginning='"&myStrDate_In_Week&"' AND (tb_Meetings.Member_ID="&myMember_ID&" OR tb_Meetings_Members.Member_ID="&myMember_ID&")) ORDER BY Meeting_Hour asc, Meeting_Length desc"

	set mySet_tb_Meetings = CreateObject("ADODB.Recordset") 
	mySet_tb_Meetings.open mySQL_Select_tb_Meetings,myConnection
        
	do while not mySet_tb_Meetings.eof
		myMeeting_Site_ID = mySet_tb_Meetings("Site_ID")
		myMeeting_Member_ID = mySet_tb_Meetings("Member_ID")
		myMeeting_ID = mySet_tb_Meetings("Meeting_ID")
		myMeeting_Title = mySet_tb_Meetings("Meeting_title")
		myMeeting_Date_Beginning = mySet_tb_Meetings("Meeting_Date_Beginning")
		myMeeting_Hour = mySet_tb_Meetings("Meeting_Hour")
		myMeeting_Minute = mySet_tb_Meetings("Meeting_Minute")
		myMeeting_Length = mySet_tb_Meetings("Meeting_Length")
		myMeeting_Length_In_Minutes = mySet_tb_Meetings("Meeting_Length_In_Minutes")
		myLM=myMeeting_Minute+myMeeting_Length_In_Minutes
		myLH=myMeeting_Hour+myMeeting_Length
		if myLM>=60 then
			myLH=myLH+1
			myLM=myLM-60
		end if
		if myLH>24 then 
			myLH=myLH-24
 		end if
	
		' Get Participants
		mySQL_Select_tb_Meetings_Members = "SELECT DISTINCT tb_Meetings_Members.Member_ID FROM tb_Meetings_Members INNER JOIN tb_Sites_Members ON tb_Meetings_Members.Member_ID= tb_Sites_Members.Member_ID WHERE (tb_Meetings_Members.Meeting_ID="&myMeeting_ID&" AND tb_Meetings_Members.Meeting_Role_ID=2) ORDER BY tb_Meetings_Members.Member_ID"
				
		set mySet_tb_Meetings_Members = CreateObject("ADODB.Recordset") 
		mySet_tb_Meetings_Members.open mySQL_Select_tb_Meetings_Members,myConnection
		%>
		<tr>
		<td bgcolor="<%= myBGColor %>" >
		<Font size='2' face="Arial,Elvetic" Color="<%= myBGTextColor %>">&nbsp;&nbsp;
	<%if myHour_Format <> 1 Then%>
<%If myMeeting_Hour < 10 Then response.write "0" & myMeeting_Hour & ":" %>
<%If myMeeting_Hour > 9 Then response.write  myMeeting_Hour & ":"%>
<%If myMeeting_Minute < 10 Then response.write "0" & myMeeting_Minute%>
<%If myMeeting_Minute > 9   Then response.write myMeeting_Minute%>
<%end if

	if myHour_Format = 1 Then
	   	If myMeeting_Hour>11 Then
		 If myMeeting_Hour <> 12 AND  myMeeting_Hour < 22 Then response.Write "0" & myMeeting_Hour-12 & ":" 
		 If myMeeting_Hour <> 12 AND myMeeting_Hour  > 21 Then response.Write myMeeting_Hour-12 & ":"
		 If myMeeting_Hour = 12 Then response.Write myMeeting_Hour & ":"
		 If myMeeting_Minute < 10 Then response.write "0" & myMeeting_Minute
         If myMeeting_Minute > 9   Then response.write myMeeting_Minute
		 response.write " P.M"
        else 
		 
		 If myMeeting_Hour < 10 Then response.write "0" & myMeeting_Hour  & ":"
		 If myMeeting_Hour > 9 Then response.write  myMeeting_Hour  & ":" 
		 If myMeeting_Minute < 10 Then response.write "0" & myMeeting_Minute
         If myMeeting_Minute > 9   Then response.write myMeeting_Minute
		 Response.write " A.M"
       end if

	   
	 end if  
	 %>
		
		- 	
<%if myHour_Format <> 1 Then%>
<%If myLH < 10 Then response.write "0" & myLH & ":"%>
<%If myLH  > 9 Then response.write  myLH & ":" %>
<%If myLM < 10 Then response.write "0" & myLM%>
<%If myLM  > 9 Then response.write  myLM %>


	
	<%end if
	if myHour_Format = 1 Then
	    If myLH >11 Then
		 If myLH <> 12 AND myLH < 22 Then response.Write "0" & myLH-12 & ":"
 		 If myLH <> 12 AND myLH > 21 Then response.Write myLH-12 & ":" 
		 If myLH = 12 Then response.Write myLH & ":"
		 If myLM < 10 Then response.write "0" & myLM
         If myLM > 9   Then response.write myLM
	
		 response.write " P.M"
        else 
		 If myLH < 10 Then response.write "0" & myLH  & ":"
		 If myLH > 9 Then response.write  myLH  & ":" 
		 If myLM < 10 Then response.write "0" & myLM
         If myLM > 9   Then response.write myLM
		 Response.write " A.M"
       end if
	 end if  
	 %>
</font><% if (myUser_ID=myMember_ID) Or (myMeeting_Site_ID=mySite_ID) then%><font  Color="<%= myBGTextColor %>"><%=myMeeting_Site_Name%> : </font><BR> <font color=blue face="Arial"><A HREF="__Agenda_Information.asp?Meeting_ID=<%=myMeeting_ID%>&Date_Agenda=<%=FormatDateTime(myIntDate_Agenda-myIntFirst_Day_Of_Week+i,2)%>"><font color="<%= myBGTextColor %>"><%=myMeeting_Title%></font></A>

		<%
		' Participants
		%>

		&nbsp;&nbsp;<Font size="1" face="Arial,Helvetica"  Color="<%= myBGTextColor %>"><%=myMessage_Participants%>&nbsp;: 
		<%
		do while not mySet_tb_Meetings_Members.EOF 
			if indice=1 then %> ; <%end if%>
			 <%
			'Get Participant Pseudo 
			myMember_Participant_ID=mySet_tb_Meetings_Members("Member_ID")
			mySQL_Select_tb_Sites_Members = "SELECT  Member_Pseudo FROM tb_Sites_Members  WHERE Site_ID="&myMeeting_Site_ID&" and Member_ID="&myMember_Participant_ID 
			set mySet_tb_Sites_Members = CreateObject("ADODB.Recordset") 
mySet_tb_Sites_Members.open mySQL_Select_tb_Sites_Members,myConnection
			myMember_Participant_Pseudo=mySet_tb_Sites_Members("Member_Pseudo")
		
			if myUser_type_ID<=myMembers_Public_type_ID then
				%>
				<A HREF="__Site_Member_Information.asp?Member_ID=<%=mySet_tb_Meetings_Members("Member_ID")%>&Site_ID=<%=myMeeting_Site_ID%>"><Font face="Arial, Helvetica, sans-serif" size="1" color="<%= myBGTextColor %>"><%=myMember_Participant_Pseudo%></font></a>
				<%
			else
				%>
				<font face="Arial, Helvetica, sans-serif" size="1" color="<%= myBGTextColor %>"><%=myMember_Participant_Pseudo%></font>
				<%
			end if
			%>
			<%	
			if indice=0 then indice=1	
			mySet_tb_Meetings_Members.MoveNext
		loop
		indice=0
		mySet_tb_Meetings_Members.close			
	end if	
	%> 
	</font>
	</td>
	</tr>
	<%							
	mySet_tb_Meetings.MoveNext
	loop
	%>
<%

i = i + 1



loop
%>
</table>
</td>
</TR>
</TABLE>

<%
' DOWN
%> 

<!-- #include file="_borders/Down.asp" --> 

<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.							'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0">
<TR ALIGN="RIGHT">
<TD>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</font>

</TD>
</TR>
</TABLE>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</body>
</html>
<% 
myConnection.Close
set myConnection = Nothing
%>
<html><script language="JavaScript"></script></html>