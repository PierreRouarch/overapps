<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Agenda_Day.asp" is free software; you can redistribute it and/or modify
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
' Does n't Work with PWS ????
%>
<%
' ----------------------------------------------------------------------------
' Name		: __Agenda_Day.asp
' Path   : /
' Description 	: Agenda by Day
' by 		: Pierre Rouarch, 		
' Company 	: OverApps
' Date		: February, 15,  2001
' Version : 1.15.0
' Contributions  : Jean-Luc Lesueur (Abawé), Christophe Humbert (Pharmagest), Dania Tcherkezof (Overapps)
'
' Modify by :
' Company
' Date :
'-----------------------------------------------------------------------------

Dim myPage
myPage = "__Agenda_Day.asp"
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
' Variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim myMember_ID, myMember_Name, myMember_Pseudo, myMember_Participant_Pseudo, myMember_Participant_ID 

Dim  myMeeting_Site_Name

DIM mySQL_Select_tb_Meetings, mySet_tb_Meetings, mySet_tb_Meetings_Members,     mySQL_Select_tb_Meetings_Members
	
Dim  test 


Dim myMeeting_Site_ID, myMeeting_Member_ID, myMeeting_Project_ID, myMeeting_Phase_ID, myMeeting_ID, myMeeting_Title, myMeeting_Date_Beginning,  myMeeting_Hour, myMeeting_Minute, myMeeting_Length, myMeeting_Length_in_Minutes, myMeeting_Place 

Dim mySup, myLM, myLH,  indice


Dim myProjects_Public_Type_ID
Dim myMembers_Public_Type_ID

myProjects_Public_Type_ID = Get_Application_Public_Type_ID("Projects")
myMembers_Public_Type_ID = Get_Application_Public_Type_ID("Members")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET Parameters
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

mytype=0 ' Day type Agenda
	
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
myStrDate_Agenda=myDate_Construct(Year(myStrDate_Agenda),Month(myStrDate_Agenda),Day(myStrDate_Agenda),0,0,0)
myStrDate_Agenda = left(myStrDate_Agenda,10)



	
' Force Member to User
myMember_ID=myUser_ID
myMember_Pseudo=myUser_Pseudo


%>

<HTML>
<HEAD>
<TITLE><%=mySite_Name%> -  <%=myMessage_Agenda%> <%=myMessage_Day%></TITLE>
</HEAD>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<!-- #include file="_borders/Top.asp" --> <TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> <%
' LEFT
%> <TD Width="<%=myLeft_Width%>"> <!-- #include file="_borders/Left.asp" --> </TD><%
' APPLICATION
%> <%
' DB Connection

set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

' SELECT only where i'm the author or a participant
		
mySQL_Select_tb_Meetings = "SELECT DISTINCT tb_Meetings.Site_ID,  tb_Meetings.Member_ID,  tb_Meetings.Project_ID,  tb_Meetings.Member_ID, tb_Meetings.Phase_ID,  tb_Meetings.Meeting_ID, tb_Meetings.Meeting_Title, tb_Meetings.Meeting_Date_Beginning,  tb_Meetings.Meeting_Hour, tb_Meetings.Meeting_Minute,  tb_Meetings.Meeting_Length, tb_Meetings.Meeting_Length_In_Minutes, tb_Meetings.Meeting_Place  FROM tb_Meetings INNER JOIN tb_Meetings_Members ON tb_Meetings_Members.Meeting_ID=tb_Meetings.Meeting_ID WHERE (Meeting_Date_Beginning='"&myStrDate_Agenda&"'  AND Meeting_Hour >= " & myAgenda_Start &"     AND (tb_Meetings.Member_ID="&myMember_ID&" OR tb_Meetings_Members.Member_ID="&myMember_ID&"))  ORDER BY Meeting_Hour asc, Meeting_Length desc"


' Set Recordset			
set mySet_tb_Meetings = CreateObject("ADODB.Recordset") 
mySet_tb_Meetings.open mySQL_Select_tb_Meetings,myConnection



 
%> 
<TD WIDTH="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>"> <TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" BGCOLOR="<%=myBGColor%>" HEIGHT="100%" CELLPADDING="0" CELLSPACING="5">
 
<TR>
<TD> 
<%
myBox_title=myApplication_Title
%>

<!-- #include file="__Agenda_Box.asp" --> 

</td>
</tr>

<%
' OutPut Current Date with navigation's arrows
%>
<TR>
<TD BGCOLOR=<%=myApplicationColor%> COLSPAN="2">
<TABLE BORDER=0 CELLPADDING=5 CELLSPACING=0 WIDTH="<%=myApplication_Width%>" > 
<TR BGCOLOR=<%=myApplicationColor%>>
<TD>
<A HREF="__Agenda_Day.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda,myDay_Agenda-1)%>"> 
<IMG BORDER=0 HEIGHT=11 SRC="Images/OverApps-left_small.gif" WIDTH=11></A> <FONT FACE=Arial,Helvetica SIZE=+1 Color="<%=myApplicationTextColor%>"><B>
<%
If datepart("w",myStrDate_Agenda) = 1 Then response.write myMessage_Sunday
If datepart("w",myStrDate_Agenda) = 2 Then response.write myMessage_Monday
If datepart("w",myStrDate_Agenda) = 3 Then response.write myMessage_Tuesday
If datepart("w",myStrDate_Agenda) = 4 Then response.write myMessage_Wednesday
If datepart("w",myStrDate_Agenda) = 5 Then response.write myMessage_Thursday
If datepart("w",myStrDate_Agenda) = 6 Then response.write myMessage_Friday
If datepart("w",myStrDate_Agenda) = 7 Then response.write myMessage_Saturday
%>&nbsp;<%
If myDate_Format = 1 Then 
 If Month(myStrDate_Agenda) = 1 Then response.write   myMessage_January
 If Month(myStrDate_Agenda) = 2 Then response.write   myMessage_February
 If Month(myStrDate_Agenda) = 3 Then response.write   myMessage_March
 If Month(myStrDate_Agenda) = 4 Then response.write   myMessage_April
 If Month(myStrDate_Agenda) = 5 Then response.write   myMessage_May
 If Month(myStrDate_Agenda) = 6 Then response.write   myMessage_June
 If Month(myStrDate_Agenda) = 7 Then response.write   myMessage_July
 If Month(myStrDate_Agenda) = 8 Then response.write   myMessage_August
 If Month(myStrDate_Agenda) = 9 Then response.write   myMessage_September
 If Month(myStrDate_Agenda) = 10 Then response.write   myMessage_October
 If Month(myStrDate_Agenda) = 11 Then response.write   myMessage_November
 If Month(myStrDate_Agenda) = 12 Then response.write   myMessage_December
 response.write ", " & Day(myStrDate_Agenda)
 
 else 
	
 response.write Day(myStrDate_Agenda)& " "
 If Month(myStrDate_Agenda) = 1 Then response.write   myMessage_January
 If Month(myStrDate_Agenda) = 2 Then response.write   myMessage_February
 If Month(myStrDate_Agenda) = 3 Then response.write   myMessage_March
 If Month(myStrDate_Agenda) = 4 Then response.write   myMessage_April
 If Month(myStrDate_Agenda) = 5 Then response.write   myMessage_May
 If Month(myStrDate_Agenda) = 6 Then response.write   myMessage_June
 If Month(myStrDate_Agenda) = 7 Then response.write   myMessage_July
 If Month(myStrDate_Agenda) = 8 Then response.write   myMessage_August
 If Month(myStrDate_Agenda) = 9 Then response.write   myMessage_September
 If Month(myStrDate_Agenda) = 10 Then response.write   myMessage_October
 If Month(myStrDate_Agenda) = 11 Then response.write   myMessage_November
 If Month(myStrDate_Agenda) = 12 Then response.write   myMessage_December
end if		
%>&nbsp;<%=Year(myStrDate_Agenda)%>

</B></FONT><A HREF="__Agenda_Day.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda,myDay_Agenda+1)%>"><IMG BORDER=0 HEIGHT=11 SRC="Images/OverApps-right_small.gif" WIDTH=11></A> </TD><TD ALIGN=right>&nbsp; 
</TD>
</TR>
</TABLE>
</TD>
</TR>


<%
' OutPut Pseudo , global vision link and Add Link
%>
<TR>
<TD BGCOLOR=<%=myBGColor%> VALIGN=top COLSPAN="2">
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="<%=myApplication_Width%>"> 
<TR>
<TD ALIGN=left WIDTH="67%">
<FONT FACE=Arial,Helvetica SIZE=-1 color="<%=myBGTextColor%>"><B><%=myUser_Pseudo%></B></FONT>
<br>
<font face="Arial" size=1 color="<%= myBGTextColor %>"><A href="__Agenda_Global_Day.asp?Date_Agenda=<%= myStrDate_Agenda %>"><b><%=myMessage_Global_Agenda%></b></a>
</TD>
<TD WIDTH="33%" ALIGN="RIGHT">
<A HREF="__Agenda_Modification.asp?Date_Agenda=<%=myStrDate_Agenda%>"> 
<FONT FACE=Arial,Helvetica SIZE=+1 Color="<%=myBGTextColor%>"><B> <%=myMessage_Add%></B></FONT></A> 
</TD>
</TR>
</TABLE>

</TD>
</TR>

<%
' Output Meetings from 7h00 to 23h00
%>

<TR VALIGN=top BGCOLOR="<%=myBGColor%>">
<TD COLSPAN="2">
<TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR BGCOLOR=#cfcfcf>
<TD>
<TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" CELLPADDING="5" CELLSPACING="1"> 

<%
if mySet_tb_Meetings.BOF then
	test=1
else
	test=0    
end if           
%>

<%
For i=myAgenda_Start to myAgenda_End 
	%>
	<TR>
	<TD BGCOLOR="#EFEFEF" WIDTH="61" ALIGN="right">
	<A HREF="__Agenda_Modification.asp?Date_Agenda=<%=myStrDate_Agenda%>&HourAgenda=<%=i%>"><FONT COLOR="#0000ff" face="Arial"><B>
<%if myHour_Format <> 1 Then
	if i < 10 Then 
	  response.write "0" & i
	else 
	  response.write i
	end if  
	%>:00
<%end if
	
	if myHour_Format = 1 Then
	     If i >11 Then
		  If i <> 12 and i < 22 Then response.Write "0" & i-12 & ":00 P.M"
		  If i <> 12 and i > 21 Then response.Write i-12 & ":00 P.M"
		  If i = 12 Then response.Write i & ":00 P.M"
        else
			if i < 10 Then 
	  response.write "0" & i
	else 
	  response.write i
	end if   
		response.write  ":00 A.M"
       end if
	 end if  
	 %>

	
	</B></FONT></A> 
	</TD>
	<%        		
	if not mySet_tb_Meetings.BOF then
		do while not mySet_tb_Meetings.eof
			 if mySet_tb_Meetings("Meeting_Hour")<>i then exit do
				myMeeting_Site_ID = mySet_tb_Meetings("Site_ID")
				myMeeting_Member_ID = mySet_tb_Meetings("Member_ID")	
				myMeeting_Project_ID = mySet_tb_Meetings("Project_ID")
				myMeeting_Phase_ID = mySet_tb_Meetings("Phase_ID")
				myMeeting_ID = mySet_tb_Meetings("Meeting_ID")
				myMeeting_Title = mySet_tb_Meetings("Meeting_title")
				myMeeting_Date_Beginning = mySet_tb_Meetings("Meeting_Date_Beginning")
				myMeeting_Hour = mySet_tb_Meetings("Meeting_Hour")
				myMeeting_Minute = mySet_tb_Meetings("Meeting_Minute")
				myMeeting_Length = mySet_tb_Meetings("Meeting_Length")
				myMeeting_Length_in_Minutes = mySet_tb_Meetings("Meeting_Length_in_Minutes")
				myMeeting_Place = mySet_tb_Meetings("Meeting_Place")
				if CInt(myMeeting_Length_in_Minutes)>0 then
					mySup=1
				else
					mySup=0
				end if
				myLM=myMeeting_Minute+myMeeting_Length_in_Minutes
				myLH=myMeeting_Hour+myMeeting_Length
				if myLM>=60 then
					myLH=myLH+1
					myLM=myLM-60
				end if
				if myLH>24 then 
					myLH=myLH-24
 				end if
	
				' Get other Participants ID
				mySQL_Select_tb_Meetings_Members = "SELECT DISTINCT tb_Meetings_Members.Member_ID FROM tb_Meetings_Members INNER JOIN tb_Sites_Members ON tb_Meetings_Members.Member_ID= tb_Sites_Members.Member_ID WHERE (tb_Meetings_Members.Meeting_ID="&myMeeting_ID&" AND tb_Meetings_Members.Meeting_Role_ID=2) ORDER BY tb_Meetings_Members.Member_ID"
				
				set mySet_tb_Meetings_Members = CreateObject("ADODB.Recordset") 
				mySet_tb_Meetings_Members.open mySQL_Select_tb_Meetings_Members,myConnection

				mySQL_Select_tb_Sites = "SELECT Site_Name FROM tb_Sites WHERE Site_ID="&myMeeting_Site_ID
				
				set mySet_tb_sites = CreateObject("ADODB.Recordset") 
				mySet_tb_Sites.open mySQL_Select_tb_Sites,myConnection
				myMeeting_Site_Name =  mySet_tb_Sites("Site_Name")

				indice=0
				%> 
				<TD VALIGN="top" ALIGN="Left" BGCOLOR=<%= myBGColor %> ROWSPAN=<%=myMeeting_Length+mySup%>> 
				<B><FONT SIZE="2" FACE="Arial,Elvetica" Color="<%= myBGTextColor %>">
			
	<%if myHour_Format <> 1 Then
	  if myMeeting_Hour < 10 Then response.write "0" & myMeeting_Hour
	  if myMeeting_Hour >= 10 Then response.write myMeeting_Hour
	 %>:<%
	 if myMeeting_Minute > 9 Then response.Write myMeeting_Minute
	 if myMeeting_Minute < 10 Then response.Write "0" & myMeeting_Minute	 
     end if
	if myHour_Format = 1 Then
	   	If myMeeting_Hour > 11 Then
		 If myMeeting_Hour <> 12 AND myMeeting_Hour < 22 Then response.Write "0" & myMeeting_Hour-12 & ":" 
		 If myMeeting_Hour <> 12 AND myMeeting_Hour > 21 Then response.Write myMeeting_Hour-12 & ":" 
		 If myMeeting_Hour = 12 Then response.Write myMeeting_Hour & ":"
		 if myMeeting_Minute=0 then response.write 0
		 Response.Write myMeeting_Minute
		 response.write " P.M"
        else 
		 If myMeeting_Hour < 10 Then
		  response.write "0" & myMeeting_Hour  & ":" 
		 else 
		  response.write  myMeeting_Hour  & ":"
		 end if
		 if myMeeting_Minute=0 then response.write 0
		 Response.Write myMeeting_Minute
		 Response.write " A.M"
       end if
	 end if  
	 %>
	-
	<%if myHour_Format <> 1 Then
	If myLH < 10 then response.write "0" & myLH & ":"
	If myLH > 9  then response.write  myLH & ":"	
    If myLM < 10 then response.write "0" & myLm 
	If myLM > 9  then response.write  myLm
	
end if
	if myHour_Format = 1 Then
	    If myLH >11 Then
		 If myLH <> 12 AND myLH < 22 Then response.Write "0" & myLH-12 & ":"
		 If myLH <> 12 AND myLH > 21 Then response.Write myLH-12 & ":" 
		 If myLH = 12 Then response.Write myLH & ":"
		 if myLM=0 then response.write "0"
		 Response.Write myLM
		 response.write " P.M"
        else 
		 If myLH < 10 Then
		  response.write "0" & myLH  & ":"
		 else
		   response.write myLH & ":"
		 end if 
		 if myLM=0 then response.write 0
		 Response.Write myLM
		 Response.write " A.M"
       end if
	 end if  
	 %>
	 : </FONT></B> <% if (myUser_ID=myMember_ID) Or (myMeeting_Site_ID=mySite_ID)   then%><font color="<%= myBGTextColor %>" face=Arial> <%=myMeeting_Place%></font><br><A HREF="__Agenda_Information.asp?Meeting_ID=<%=myMeeting_ID%>&Date_Agenda=<%=myStrDate_Agenda%>"><font color="<%= myBGTextColor %>" Face="Arial"><%=myMeeting_Title%></A><BR><FONT SIZE="1" FACE="Arial,Helvetica" Color="<%= myBGTextColor %>"><%if not mySet_tb_Meetings_Members.EOF Then response.write myMessage_Participants& "&nbsp;:"%> 
				<%
				 do while not mySet_tb_Meetings_Members.EOF 
						if indice=1 then%>; <%end if
					
					'GEt Participant Pseudo 
					myMember_Participant_ID=mySet_tb_Meetings_Members("Member_ID")
					mySQL_Select_tb_Sites_Members = "SELECT  Member_Pseudo FROM tb_Sites_Members  WHERE	 Site_ID="&myMeeting_Site_ID&" and Member_ID="&myMember_Participant_ID 

					set mySet_tb_Sites_Members = CreateObject("ADODB.Recordset") 
					mySet_tb_Sites_Members.open mySQL_Select_tb_Sites_Members,myConnection
					myMember_Participant_Pseudo=mySet_tb_Sites_Members("Member_Pseudo")
					%> 
					<% 
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
				mySet_tb_Meetings_Members.close
		end if				
		%> 
		</FONT>
		</TD>
		<%					
		mySet_tb_Meetings.MoveNext
		loop
	else 
		%> 
		<TD BGCOLOR="efefef">
		</TD>
		<% 
	end if
	%> 
	</TR> 
	<%
next
%> 
</TABLE>
</TD>
</TR> 

</TABLE>
</TD>
</TR>
</TABLE>

<%
' / CENTER APPLICATION
%> 

</TD>
</TR>
</TABLE>

<%
' / CENTER
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
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</BODY>
</HTML>
<% 
	myConnection.Close
	set myConnection = Nothing
%>
<html><script language="JavaScript"></script></html>