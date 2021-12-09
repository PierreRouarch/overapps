<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Agenda_Month.asp" is free software; you can redistribute it and/or modify
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
' Doesn't Work with PWS
%>

<%
' ------------------------------------------------------------
' Name			: __Agenda_Month.asp
' Path    		: /
' Version 		: 1.15.0
' Description 	:  Agenda Month Presentation
' By 			: Pierre Rouarch
' Company		: Overapps
' Date			: December,10 , 2001
'Contributor : Dania Tcherkezoff
'
' Modify by	:
' Company	:
' Date		:
' ------------------------------------------------------------

Dim myPage
myPage = "__Agenda_Month.asp"
	
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Local Variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim  mySQL_Select_tb_Meetings, mySet_tb_Meetings, mySQL_Select_tb_Meetings_Members, mySet_tb_Meetings_Members

Dim h, myNumDate_Agenda, V, myfcolors, myNumDayMonth

Dim   myMeeting_Title, myMeeting_Date_Beginning, myMeeting_ID, myMeeting_Hour, myMeeting_Minute, myMeeting_Length, myMeeting_Length_In_Minutes, mySup, myLM, myLH, myPrive, indice

Dim myMember_ID, myMember_Pseudo, myMeeting_Site_ID, myMeeting_Site_Name

Dim myMember_Participant_ID, myMember_Participant_Pseudo


Dim myStrDate_In_Month
Dim myIntDate_In_Month
Dim myDay_In_Month

Dim myProjects_Public_Type_ID
Dim myMembers_Public_Type_ID

myProjects_Public_Type_ID = Get_Application_Public_Type_ID("Projects")
myMembers_Public_Type_ID = Get_Application_Public_Type_ID("Members")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Parameters
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

mytype=2 ' Month Presentation

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


'  First day of the month
myIntFirst_Day_Of_Month = DateSerial(Year(myStrDate_Agenda),Month(myStrDate_Agenda),1)
myNumDay=DatePart("w",myIntFirst_Day_Of_Month)-1
myNumDayMonth = DatePart("m",myStrDate_Agenda) - 1


' Display Month and year

	Function myDisplayMonthAndYear (myString)
		myString=Mid(myString,InStr(2,myString," "))
		myString=Mid(myString,InStr(2,myString," "))
		myDisplayMonthAndYear=myString
	End Function
	
' Maybe Not Be used
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
<TITLE><%=mySite_Name%> - Agenda <%=myMessage_Month%></TITLE>
</HEAD>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>" marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
'TOP
%> 

<!-- #include file="_borders/Top.asp" -->

<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BGColor="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
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

<td BGCOLOR="<%=myBGColor%>" width="<%=myApplication_Width%>">

<table border=0 bgcolor="<%=myBGColor%>" cellpadding="5" width="<%=myApplication_Width%>" cellspacing="0" > 

<%
' CENTER AGENDA NAVIGATION
%> 

<TR> 

<TD colspan=8> 

<%
myBox_title=myApplication_Title
%>

<!-- #include file="__Agenda_Box.asp" --> 

</td>
</tr> 


<%
' CENTER MONTH PRESENTATION
%> 

<tr>
<td bgcolor="<%=myApplicationColor%>" colspan=8>
<table border=0 cellpadding=3 cellspacing=0 width="<%=myApplication_Width%>"> 
<tr bgcolor="<%=myApplicationColor%>"> 
<td>
<a href="__Agenda_Month.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda-1,1)%>"><img border=0 height=11 src="Images/OverApps-left_small.gif" width=11></a>
<font face=Arial,Helvetica size=+1 color="<%=myApplicationtextColor%>">
<b>
<%
If myDate_Format = 1 Then 
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 1 Then response.write   myMessage_January
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 2 Then response.write   myMessage_February
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 3 Then response.write   myMessage_March
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 4 Then response.write   myMessage_April
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 5 Then response.write   myMessage_May
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 6 Then response.write   myMessage_June
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 7 Then response.write   myMessage_July
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 8 Then response.write   myMessage_August
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 9 Then response.write   myMessage_September
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 10 Then response.write   myMessage_October
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 11 Then response.write   myMessage_November
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 12 Then response.write   myMessage_December
 response.write ", " & Day(myIntFirst_Day_Of_Month-myNumDay)
 
 else 
	
 response.write Day(myIntFirst_Day_Of_Month-myNumDay)& " "
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 1 Then response.write   myMessage_January
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 2 Then response.write   myMessage_February
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 3 Then response.write   myMessage_March
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 4 Then response.write   myMessage_April
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 5 Then response.write   myMessage_May
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 6 Then response.write   myMessage_June
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 7 Then response.write   myMessage_July
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 8 Then response.write   myMessage_August
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 9 Then response.write   myMessage_September
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 10 Then response.write   myMessage_October
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 11 Then response.write   myMessage_November
 If Month(myIntFirst_Day_Of_Month-myNumDay) = 12 Then response.write   myMessage_December
end if		
%>&nbsp;<%=Year(myIntFirst_Day_Of_Month-myNumDay)%>

<font size="+1" face="Courier New, Courier, mono"><b>--> </b></font>
<%
If myDate_Format = 1 Then 
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 1 Then response.write   myMessage_January
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 2 Then response.write   myMessage_February
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 3 Then response.write   myMessage_March
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 4 Then response.write   myMessage_April
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 5 Then response.write   myMessage_May
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 6 Then response.write   myMessage_June
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 7 Then response.write   myMessage_July
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 8 Then response.write   myMessage_August
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 9 Then response.write   myMessage_September
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 10 Then response.write   myMessage_October
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 11 Then response.write   myMessage_November
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 12 Then response.write   myMessage_December
 response.write ", " & Day(myIntFirst_Day_Of_Month-myNumDay+41)
 
 else 
	
 response.write Day(myIntFirst_Day_Of_Month-myNumDay+41)& " "
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 1 Then response.write   myMessage_January
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 2 Then response.write   myMessage_February
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 3 Then response.write   myMessage_March
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 4 Then response.write   myMessage_April
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 5 Then response.write   myMessage_May
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 6 Then response.write   myMessage_June
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 7 Then response.write   myMessage_July
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 8 Then response.write   myMessage_August
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 9 Then response.write   myMessage_September
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 10 Then response.write   myMessage_October
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 11 Then response.write   myMessage_November
 If Month(myIntFirst_Day_Of_Month-myNumDay+41) = 12 Then response.write   myMessage_December
end if		
%>&nbsp;<%=Year(myIntFirst_Day_Of_Month-myNumDay+41)%>








 
</b></font>
<a href="__Agenda_Month.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda+1,1)%>"><img border=0 height=11 src="Images/OverApps-right_small.gif" width=11></a> 
</td>
<td align=right>&nbsp;

</td>
</tr> 
</table>
</td>
</tr>


<%
' OutPut Pseudo and Default Day Add Meeting Link
%> 

<TR>
<TD BGCOLOR=<%=myBGColor%> VALIGN=top COLSPAN="8">
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="<%=myApplication_Width%>"> 
<TR>
<TD ALIGN=left WIDTH="67%">
<FONT FACE=Arial,Helvetica SIZE=-1 color="<%=myBGTextColor%>"> 
<B><%=myUser_Pseudo%></B></FONT>
</TD>
<TD WIDTH="33%" ALIGN="RIGHT">
<A HREF="__Agenda_Modification.asp?Date_Agenda=<%=myStrDate_Agenda%>"> 
<FONT FACE=Arial,Helvetica SIZE=+1 Color="<%=myBGTextColor%>"><B><%=myMessage_Add%></B></FONT></A> 
</TD>
</TR>
</TABLE>
</TD>
</TR>


<% 
' DAY TITLES
%> 

<tr bgcolor="<%=myApplicationColor%>">

<td width="16%" bgcolor="#cccccc">
<br> 
</td>

<%If myAgenda_Week_Start = 0 Then%>
<td width="12%" bgcolor="#cccccc">
<b><font face=Arial,Helvetica size=-1 color=Black><%=myMessage_Sunday%></font></b>
</td>
<%end if%>

<td width="12%" bgcolor="#cccccc">
<b><font face=Arial,Helvetica size=-1  color=Black><%=myMessage_Monday%></font></b>
</td>
<td width="12%" bgcolor="#cccccc">
<b><font face=Arial,Helvetica size=-1  color=Black><%=myMessage_Tuesday%></font></b>
</td>
<td width="12%" bgcolor="#cccccc" >
<b><font face=Arial,Helvetica size=-1 color=Black><%=myMessage_Wednesday%></font></b>
</td><td width="12%" bgcolor="#cccccc">
<b><font face=Arial,Helvetica size=-1 color=Black><%=myMessage_Thursday%></font></b>
</td>
<td width="12%" bgcolor="#cccccc">
<b><font face=Arial,Helvetica size=-1 color=Black><%=myMessage_Friday%></font></b>
</td>

<td width="12%" bgcolor="#cccccc" >
<b><font face=Arial,Helvetica size=-1 color=Black><%=myMessage_Saturday%></font></b>
</td>


<%If myAgenda_Week_Start = 1 Then%>
<td width="12%" bgcolor="#cccccc">
<b><font face=Arial,Helvetica size=-1 color=Black><%=myMessage_Sunday%></font></b>
</td>
<%end if%>

</tr>
 
<% 
	For h=0 to 35 Step 7 
%> 
	<tr> 
	<% ' Week Row %> 
	<% myNumDate_Agenda=myIntFirst_Day_Of_Month-myNumDay+h	  %> 
	<td bgcolor=#dcdcdc rowspan=2>
	<% ' Day %>
	<small><a href="__Agenda_Week.asp?Date_Agenda=<%=DateSerial(Year(myNumDate_Agenda), Month(myNumDate_Agenda), Day(myNumDate_Agenda))%>"><font color="#0000ff" face="Arial,Helvetica" size="2"><%=myMessage_Week%>&nbsp;<%=DatePart("ww",myNumDate_Agenda)%></font></a></small>
	</td>
	<% 
	For v=myAgenda_Week_Start to (6+myAgenda_Week_Start) 
		myIntDate_In_Month=myIntFirst_Day_Of_Month-myNumDay+h+v
        if Month(myIntDate_In_Month)<>Month(myIntDate_Agenda) then
			mycolors="#dcdcdc"
			myfcolors="#000000"
		elseIf myIntDate_In_Month=myIntDate_Agenda then
			mycolors="MistyRose"
			myfcolors="#0000ff"
		else
			mycolors=myBorderColor
			myfcolors=myBorderTextColor
		end if
        %> 
		<td bgcolor="<%=mycolors%>">
		<a href="__Agenda_Day.asp?Date_Agenda=<%=DateSerial(Year(myIntDate_In_Month), Month(myIntDate_In_Month), Day(myIntDate_In_Month))%>"><font face="Arial,Helvetica" color="<%=myfcolors%>"><b><%=Day(myIntDate_In_Month)%></b></font></a>
		<a href="__Agenda_Modification.asp?Date_Agenda=<%=Year(myIntDate_In_Month)&"/"& Month(myIntDate_In_Month)&"/"& Day(myIntDate_In_Month)%>"><font face="Arial,Helvetica" color="<%=myfcolors%>"><b><%=myMessage_Add%></b></font></a> 
		</td>
	<%
	Next
	%> 
	</tr> 
	<tr bgcolor=#ffffff >
	<% '  Meetings Row %> 
	<% 
	For i=myAgenda_Week_Start to (6+myAgenda_Week_Start) 
              
	myIntDate_In_Month=myIntFirst_Day_Of_Month-myNumDay+h+i
	myDay_In_Month=myIntFirst_Day_Of_Month-myNumDay+h+i

	myStrDate_In_Month= myDate_Construct(Year(myIntDate_In_Month),Month(myIntDate_In_Month),Day(myIntDate_In_Month),0,0,0)

   myStrDate_In_Month = left(myStrDate_In_Month,10)
	

  
	' Search Meetings
    
set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String
		

' SELECT only where i'm the author or a participant
		
mySQL_Select_tb_Meetings = "SELECT DISTINCT tb_Meetings.Site_ID,  tb_Meetings.Member_ID,  tb_Meetings.Project_ID,  tb_Meetings.Member_ID, tb_Meetings.Phase_ID,  tb_Meetings.Meeting_ID, tb_Meetings.Meeting_Title, tb_Meetings.Meeting_Date_Beginning,  tb_Meetings.Meeting_Hour, tb_Meetings.Meeting_Minute,  tb_Meetings.Meeting_Length, tb_Meetings.Meeting_Length_In_Minutes  FROM tb_Meetings INNER JOIN tb_Meetings_Members ON tb_Meetings_Members.Meeting_ID=tb_Meetings.Meeting_ID WHERE (Meeting_Date_Beginning='"&myStrDate_In_Month&"'  AND (tb_Meetings.Member_ID="&myMember_ID&" OR tb_Meetings_Members.Member_ID="&myMember_ID&"))  ORDER BY Meeting_Hour asc, Meeting_Length desc"

				
set mySet_tb_Meetings = CreateObject("ADODB.Recordset") 
mySet_tb_Meetings.open mySQL_Select_tb_Meetings,myConnection
%> <td bgcolor="<%= myBGCOLOR %>" VALIGN="TOP"> <% do while not mySet_tb_Meetings.eof 
		 

		myMeeting_Title = mySet_tb_Meetings("Meeting_title")
		myMeeting_Date_Beginning = mySet_tb_Meetings("Meeting_Date_Beginning")
		myMeeting_Site_ID = mySet_tb_Meetings("Site_ID")
		myMeeting_ID = mySet_tb_Meetings("Meeting_ID")
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

				
' Get Participants ID	
		mySQL_Select_tb_Meetings_Members = "SELECT DISTINCT tb_Meetings_Members.Member_ID FROM tb_Meetings_Members INNER JOIN tb_Sites_Members ON tb_Meetings_Members.Member_ID= tb_Sites_Members.Member_ID WHERE (tb_Meetings_Members.Meeting_ID="&myMeeting_ID&" AND tb_Meetings_Members.Meeting_Role_ID=2) ORDER BY tb_Meetings_Members.Member_ID"


			
		set mySet_tb_Meetings_Members = CreateObject("ADODB.Recordset") 
		mySet_tb_Meetings_Members.open mySQL_Select_tb_Meetings_Members,myConnection



		%> <Font size="1" face="Arial" Color="<%= myBGTextColor %>">
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
	&nbsp;	<%if myHour_Format <> 1 Then%>
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

</font> <br><% if (myUser_ID=myMember_ID) Or (myMeeting_Site_ID=mySite_ID)   then%> 
<A HREF="__Agenda_Information.asp?Meeting_ID=<%=myMeeting_ID%>&Date_Agenda=<%=DateSerial(Year(myIntDate_In_Month), Month(myIntDate_In_Month), Day(myIntDate_In_Month))%>"><font color="<%= myBGTextColor %>" face="Arial,Helvetica"><%=myMeeting_Title%></font></A> 
&nbsp; <Font size="1" face="Arial,Helvetica" color="<%= myBGTextColor %>"><%=myMessage_Participants%> : <%do while not mySet_tb_Meetings_Members.EOF 
				if indice=1 then%>; <%end if%> <%



'GEt Participant Pseudo 
myMember_Participant_ID=mySet_tb_Meetings_Members("Member_ID")
mySQL_Select_tb_Sites_Members = "SELECT  Member_Pseudo FROM tb_Sites_Members  WHERE Site_ID="&myMeeting_Site_ID&" and Member_ID="&myMember_Participant_ID 

		set mySet_tb_Sites_Members = CreateObject("ADODB.Recordset") 
		mySet_tb_Sites_Members.open mySQL_Select_tb_Sites_Members,myConnection
		myMember_Participant_Pseudo=mySet_tb_Sites_Members("Member_Pseudo")


		
%> 

			<%
			if myUser_type_ID<=myMembers_Public_type_ID then
				%>
				<A HREF="__Site_Member_Information.asp?Member_ID=<%=mySet_tb_Meetings_Members("Member_ID")%>&Site_ID=<%=myMeeting_Site_ID%>"><Font face="Arial, Helvetica, sans-serif" size="1" ><font color="<%= myBGTextColor %>"><%=myMember_Participant_Pseudo%></font></font></a>
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
				%> </font> <BR><br> <%							
mySet_tb_Meetings.MoveNext
	
	loop%> </td><% 
myConnection.Close
set myConnection = Nothing
%> <%Next%> </tr> <%Next%> </table></td></TR> </TABLE>


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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> 
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</BODY></HTML>

<html><script language="JavaScript"></script></html>