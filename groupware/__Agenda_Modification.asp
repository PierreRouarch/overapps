<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   = OverApps = http://www.overapps.com
'
' This program "__Agenda_Modification.asp" is free software; you can 
' redistribute it and/or modify
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
<% 	
Option Explicit 
Response.Buffer = true	
Response.ExpiresAbsolute = Now () - 1
Response.Expires = 0
' Response.CacheControl = "no-cache"  ' Doesn't work with  PWS ?????
%>


<%
' ------------------------------------------------------------
' Name 			: __Agenda_Modification.asp
' Path  	 	: /
' Version 		: 1.15.0
' description 	: Meeting insertion or modification
' by 			: Pierre Rouarch
' Company		: OverApps
' Date			: November, 21, 2001
'
' Contributions  : Jean-Luc Lesueur, Christophe Humbert, Dania Tcherkezoff 
'
' Modify by 	:
' Company 	:
' Date		:
' ------------------------------------------------------------
Dim myPage
myPage = "__Agenda_Modification.asp"

Dim myPage_Application
myPage_Application="Agenda"

%>


<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' INCLUDES 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<!-- #include file="_INCLUDE/Global_Parameters.asp" -->

<!-- #include file="_INCLUDE/Form_validation.asp" -->

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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Variables Definitions													'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim  myHourAgenda, myHour_Indicator
Dim  my_Day, my_Month, my_Year, mylocation

Dim myMeeting_Site_ID, myMeeting_Member_ID, myMeeting_Project_ID, myMeeting_ID, myMeeting_Title, myMeeting_Date_Beginning,  myMeeting_Hour, myMeeting_Minute, myMeeting_Length, myMeeting_Length_In_Minutes, myMeeting_Place,  myMeeting_Agenda, myMeeting_Comments, myMeeting_Author_Update, myMeeting_Date_Update


Dim  myAction, myTitle

Dim mySQL_Select_tb_Meetings,  mySQL_Delete_tb_Meetings, mySQL_Insert_tb_Meetings, mySQL_Update_tb_Meetings,  mySet_tb_Meetings

Dim mySQL_Select_tb_Projects, mySet_tb_Projects 
Dim myProject_ID, myProject_Name, myMeeting_Public


Dim  mySQL_Select_tb_Meetings_Members, mySet_tb_Meetings_Members, mySQL_Insert_tb_Meetings_members,  mySQL_Delete_tb_Meetings_Members 

Dim URL, myCounter, myParticipant_ID, myMember_Name, myParticipant_Pseudo, myParticipant_Pseudo_Value

Dim myProjects_Public_Type_ID
Dim myMembers_Public_Type_ID

myProjects_Public_Type_ID = Get_Application_Public_Type_ID("Projects")
myMembers_Public_Type_ID = Get_Application_Public_Type_ID("Members")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Parameters			  												'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

myAction = Request("Action")
myMeeting_ID = Request("Meeting_ID")
mylocation = Request.QueryString("location")
if len(myLocation) = 0 Then myLocation = Request.Form("location")

' Get Date Agenda
myStrDate_Agenda=request("Date_Agenda")

if Not isDate(myStrDate_Agenda) then 
	myStrDate_Agenda = Date
end if

' get Month
myMonth_Agenda= Month(myStrDate_Agenda)
if Len(myMonth_Agenda)=0 then
	myMonth_Agenda=Month(now())
end if

if Len(myMonth_Agenda)=1 then
	myMonth_Agenda= "0" & 	myMonth_Agenda
end if


my_Month=myMonth_Agenda

' Get Day
myDay_Agenda= Day(myStrDate_Agenda)
if Len(myDay_Agenda)=0 then
	myDay_Agenda=Day(now())
end if

if Len(myDay_Agenda)=1 then
	myDay_Agenda= "0" & myDay_Agenda
end if


my_Day=myDay_Agenda

' Get Year
myYear_Agenda= Year(myStrDate_Agenda)
if Len(myYear_Agenda)=0 then
	myYear_Agenda=Year(now())	
end if
my_Year=myYear_Agenda
' In System Format
myStrDate_Agenda=DateSerial(Year(myStrDate_Agenda),Month(myStrDate_Agenda),Day(myStrDate_Agenda))


myMeeting_Date_Beginning=myStrDate_Agenda

myHourAgenda=Request("HourAgenda")
if len(myHourAgenda)=0 or myHourAgenda<7 or myHourAgenda>23 then
	myHourAgenda = 7
end if
myMeeting_Hour=myHourAgenda




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Delete Meeting		 												  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myAction="Delete" then
		
	' DB Connection
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

	' Delete in tb_Meetings and in tb_Meetings_Members
 	mySQL_Delete_tb_Meetings_Members = "DELETE  FROM tb_Meetings_Members WHERE Meeting_ID =" & myMeeting_ID
 	myConnection.Execute(mySQL_Delete_tb_Meetings_Members)
 	mySQL_Delete_tb_Meetings = "DELETE  FROM tb_Meetings WHERE Meeting_ID =" & myMeeting_ID 
 	myConnection.Execute(mySQL_Delete_tb_Meetings)

	' Close Connection
	myConnection.close
	set myConnection = nothing
	' and go back
	If myLocation= "global" Then
		URL="__Agenda_Global_Day.asp?Date_Agenda="&myStrDate_Agenda
	else
			URL="__Agenda_Day.asp?Date_Agenda="&myStrDate_Agenda
	end if		
	Response.Redirect(URL)

End if



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation														'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if Request.form("Validation")=myMessage_Go then

	' Get inputs :

	myMeeting_Title =Replace(Request.Form("Meeting_Title"),"'"," ")
	my_Day	= Request.Form("Day")
	my_Month	= Request.Form("Month")
	my_Year	= Request.Form("Year")
	if (my_Day<> "" and my_Month <> "" and my_Year <> "") then
		myMeeting_Date_Beginning = left(myDate_Construct(my_Year,my_Month,my_Day,0,0,0),10)
	end if
	myMeeting_Project_ID = Request.Form("Project_ID") 

	if len(myMeeting_Project_ID)=0 then 
		myMeeting_Project_ID=0
	end if

	myMeeting_Hour = Request.Form("Meeting_Hour")
	
	myMeeting_Public = Request.Form("Meeting_Public")
	
	'TEST HOUR FOR DIFFERENT FORMAT
	if myHour_Format = 1 Then
    myHour_Indicator = Request.Form("Hour_Indicator")
    if myHour_Indicator = 2 AND myMeeting_Hour <> 12 Then myMeeting_Hour = myMeeting_Hour + 12
	if myMeeting_Hour < 7   Then myMeeting_Hour = 7
	if myMeeting_Hour > 23 Then myMeeting_Hour = 23
	
	end if
	myMeeting_Length = Request.Form("Meeting_Length")
	myMeeting_Minute = Request.Form("myMeeting_Minute")
	myMeeting_Length_In_Minutes= Request.Form("Meeting_Length_In_Minutes")

	myMeeting_Place = Replace(Request.Form("Meeting_Place"),"'"," ")
	myMeeting_Agenda = Replace(Request.Form("Meeting_Agenda"),"'"," ")
	myMeeting_Comments = Replace(Request.Form("Meeting_Comments"),"'"," ")

	' Test inputs :
	Call myFormSetEntriesInString

	myFormCheckEntry null, "Meeting_Title",true,null,null,0,100
	myFormCheckEntry myErr_Numerical, "Day",true,null,null,0,2
	myFormCheckEntry null, "Month",true,null,null,0,2
	myFormCheckEntry myErr_Numerical, "Year",true,1990,2050,4,4


	if not myform_entry_error and isDate(myMeeting_Date_Beginning) then 

		myStrDate_Agenda=myMeeting_Date_Beginning  ' to go back at the new date
		myMeeting_Author_Update	= myUser_Pseudo
		myMeeting_Date_Update	= myDate_Now()

		' DB Connection
		set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String


		if myAction="Update" then

			' Delete all participants they will be re-insert after
			mySQL_Delete_tb_Meetings_Members = "DELETE  FROM tb_Meetings_Members WHERE Meeting_ID ="&myMeeting_ID
			 myConnection.Execute(mySQL_Delete_tb_Meetings_Members)

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' UPDATE MEETING in tb_Meetings
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			mySQL_Select_tb_Meetings = "SELECT * FROM tb_Meetings WHERE Meeting_ID =" & myMeeting_ID
			Set mySet_tb_Meetings= Server.CreateObject("ADODB.Recordset")
	    	mySet_tb_Meetings.open mySQL_Select_tb_Meetings, myConnection, 3, 3


			mySet_tb_Meetings.fields("Project_ID")=myMeeting_Project_ID
			mySet_tb_Meetings.fields("Meeting_Title")=myMeeting_Title
			mySet_tb_Meetings.fields("Meeting_Date_Beginning")=myMeeting_Date_Beginning
			mySet_tb_Meetings.fields("Meeting_Hour")=myMeeting_Hour
			mySet_tb_Meetings.fields("Meeting_Minute")=myMeeting_Minute
			mySet_tb_Meetings.fields("Meeting_Length")=myMeeting_Length
			mySet_tb_Meetings.fields("Meeting_Length_In_Minutes")=myMeeting_Length_In_Minutes
			mySet_tb_Meetings.fields("Meeting_Place")=myMeeting_Place
			mySet_tb_Meetings.fields("Meeting_Agenda")=myMeeting_Agenda
			mySet_tb_Meetings.fields("Meeting_Comments")=myMeeting_Comments
			mySet_tb_Meetings.fields("Meeting_Date_Update")=myMeeting_Date_Update
			mySet_tb_Meetings.fields("Meeting_Author_Update")=myMeeting_Author_Update
			mySet_tb_Meetings.fields("Meeting_Public") = myMeeting_Public
			
			mySet_tb_Meetings.Update
			'Close Recordset
			mySet_tb_Meetings.close
			Set mySet_tb_Meetings = Nothing

		else 

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' INSERT Meeting IN tb_Meetings
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
			mySQL_Select_tb_Meetings = "SELECT * FROM tb_Meetings"
			Set mySet_tb_Meetings = server.createobject("adodb.recordset")
			mySet_tb_Meetings.open mySQL_Select_tb_Meetings, myConnection, 3, 3
			mySet_tb_Meetings.AddNew
			mySet_tb_Meetings.fields("Site_ID")=mySite_ID
			mySet_tb_Meetings.fields("Member_ID")=myUser_ID
			mySet_tb_Meetings.fields("Project_ID")=myMeeting_Project_ID
			mySet_tb_Meetings.fields("Meeting_Title")=myMeeting_Title
			mySet_tb_Meetings.fields("Meeting_Date_Beginning")=myMeeting_Date_Beginning
			mySet_tb_Meetings.fields("Meeting_Hour")=myMeeting_Hour
			mySet_tb_Meetings.fields("Meeting_Minute")=myMeeting_Minute
			mySet_tb_Meetings.fields("Meeting_Length")=myMeeting_Length
			mySet_tb_Meetings.fields("Meeting_Length_In_Minutes")=myMeeting_Length_In_Minutes
			mySet_tb_Meetings.fields("Meeting_Place")=myMeeting_Place
			mySet_tb_Meetings.fields("Meeting_Agenda")=myMeeting_Agenda
			mySet_tb_Meetings.fields("Meeting_Comments")=myMeeting_Comments
			mySet_tb_Meetings.fields("Meeting_Date_Update")=myMeeting_Date_Update
			mySet_tb_Meetings.fields("Meeting_Author_Update")=myMeeting_Author_Update
			mySet_tb_Meetings.fields("Meeting_Public") = myMeeting_Public

			mySet_tb_Meetings.Update

			' Close Recordset 
	  		mySet_tb_Meetings.close
	  		Set mySet_tb_Meetings = Nothing



	
			' Get the Meeting ID  rem : "SELECT @@IDENTITY AS myMeeting_ID" doesn't Work with ACCESS
	
	     	mySQL_Select_tb_Meetings="SELECT Meeting_ID FROM tb_Meetings where Site_ID="&mySite_ID&" AND  Member_ID="&myUser_ID&" AND  Meeting_title='"&myMeeting_title &"' AND Meeting_Date_Beginning='"&myMeeting_Date_Beginning&"' AND Meeting_Author_Update='"&myMeeting_Author_Update&"'" 
            response.write mySQL_Select_tb_Meetings
			Set mySet_tb_Meetings = server.createobject("adodb.recordset")
			mySet_tb_Meetings.open mySQL_Select_tb_Meetings, myConnection, 3, 3

			mySet_tb_Meetings.MoveLast
			myMeeting_ID=mySet_tb_Meetings("Meeting_ID")


			' Close Recordset 
		  	mySet_tb_Meetings.close
		  	Set mySet_tb_Meetings = Nothing

		end if 'UPDATE OR NEW

		''''''''''''''''''''''''''''''''''
		' INSERT PARTICIPANTS AND AUTHOR
		''''''''''''''''''''''''''''''''''

		'Get and put members in tb_meeting_members

		mySQL_Select_tb_Sites_Members = "Select * from tb_Sites_Members WHERE  Site_ID="&mySite_ID 
		set mySet_tb_Sites_Members = myConnection.Execute(mySQL_Select_tb_Sites_Members)

		do while not mySet_tb_Sites_Members.eof

			myParticipant_ID=mySet_tb_Sites_Members("Member_ID")
			myParticipant_Pseudo=mySet_tb_Sites_Members("Member_Pseudo")
			myParticipant_Pseudo_Value=request.form(myParticipant_Pseudo)

			if len(myParticipant_Pseudo_Value)>0 then 
		
				mySQL_Select_tb_Meetings_Members = "SELECT * FROM tb_Meetings_Members"
				Set mySet_tb_Meetings_Members = server.createobject("adodb.recordset")
				mySet_tb_Meetings_Members.open mySQL_Select_tb_Meetings_Members, myConnection, 3, 3
				mySet_tb_Meetings_Members.AddNew

				mySet_tb_Meetings_Members.fields("Site_ID")=mySite_ID
				mySet_tb_Meetings_Members.fields("Member_ID")=myParticipant_ID
				mySet_tb_Meetings_Members.fields("Meeting_ID")=myMeeting_ID
				mySet_tb_Meetings_Members.fields("Meeting_Role_ID")=2

				mySet_tb_Meetings_Members.Update
				' Close Recordset 
		  		mySet_tb_Meetings_Members.close
		  		Set mySet_tb_Meetings_Members = Nothing

			end if 'Participant Checked


			' Update Author
			if myParticipant_ID=myUser_ID then 


				mySQL_Select_tb_Meetings_Members = "SELECT * FROM tb_Meetings_Members"
				Set mySet_tb_Meetings_Members = server.createobject("adodb.recordset")
				mySet_tb_Meetings_Members.open mySQL_Select_tb_Meetings_Members, myConnection, 3, 3
				mySet_tb_Meetings_Members.AddNew

				mySet_tb_Meetings_Members.fields("Site_ID")=mySite_ID
				mySet_tb_Meetings_Members.fields("Member_ID")=myUser_ID
				mySet_tb_Meetings_Members.fields("Meeting_ID")=myMeeting_ID
				mySet_tb_Meetings_Members.fields("Meeting_Role_ID")=1

				mySet_tb_Meetings_Members.Update
				' Close Recordset 
		  		mySet_tb_Meetings_Members.close
		  		Set mySet_tb_Meetings_Members = Nothing


			end if ' Participant=User


			mySet_tb_Sites_Members.movenext

		loop

		' Close Connection
		myConnection.close
		set myConnection = nothing
		if mylocation = "global" Then
			URL="__Agenda_Global_Day.asp?Date_Agenda="&myStrDate_Agenda
		else			
			URL="__Agenda_Day.asp?Date_Agenda="&myStrDate_Agenda
		end if	
		Response.Redirect(URL)

	end if ' NO ERROR VALIDATION

end if ' VALIDATION

%>

<HTML><HEAD></HEAD><TITLE><%=mySite_Name%> - Add/Modification Agenda</TITLE>
<BODY  bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>" marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' BODY TOP
%> 

<!-- #include file="_borders/Top.asp" --> 

<%
' BODY CENTER
%> 

<TABLE WIDTH="<%=myGlobal_Width%>" bgcolor=<%=mybordercolor%> BORDER="0" CELLPADDING="0" CELLSPACING="0"> 

<TR VALIGN="TOP"> 


<%
' BODY CENTER LEFT
%> 
<TD WIDTH="<%=myLeft_Width%>"> 
<!-- #include file="_borders/Left.asp" -->
 </TD>

<%
' BODY CENTER APPLICATION
%> 

<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form															'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if len(myMeeting_ID)=0 or myMeeting_ID=0 then

	' NEW MEETING
	myMeeting_ID=0
	myAction = "New"
	myTitle  = myMessage_Add
	if Len(myMeeting_Date_Beginning)=0 then myMeeting_Date_Beginning = Date()
	if Len(myMeeting_Hour)=0 then myMeeting_Hour=7

else
	' MEETING MODIFICATION
	myAction = "Update"
	myTitle  = myMessage_Modification
	myMeeting_ID=CInt(myMeeting_ID)


	' Get Fields Values in the db if it's the first time 
	if Request.form("Validation")<>myMessage_Go then 

		' Read in the DB	
		set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String

		' SELECT - VERIFICATION IF IT'S The good Site (for multi-site Extension)

mySQL_Select_tb_Meetings = "SELECT *,Meeting_Author_Update FROM tb_Meetings  WHERE Meeting_ID="&myMeeting_ID&" AND (Site_ID="&mySite_ID&" OR Member_ID="&myUser_ID&")"

		set mySet_tb_meetings = myConnection.Execute(mySQL_Select_tb_Meetings)

		' Get Fields
		if not mySet_tb_meetings.eof then
			myMeeting_Site_ID = mySet_tb_meetings("Site_ID")
			myMeeting_Member_ID	= mySet_tb_meetings("Member_ID")
			myMeeting_Project_ID	= mySet_tb_meetings("Project_ID")
			myMeeting_ID = mySet_tb_meetings("Meeting_ID")
			myMeeting_Title	= mySet_tb_meetings("Meeting_Title")
			myMeeting_Date_Beginning = mySet_tb_meetings("Meeting_Date_Beginning")
			myStrDate_Agenda=myMeeting_Date_Beginning
			' get Month
			myMonth_Agenda= Month(myStrDate_Agenda)
			my_Month=myMonth_Agenda
			' Get Day
			myDay_Agenda = Day(myStrDate_Agenda)
			my_Day=myDay_Agenda
			' Get Year
			myYear_Agenda= Year(myStrDate_Agenda)
			my_Year=myYear_Agenda
			myMeeting_Hour = mySet_tb_meetings("Meeting_Hour")
			myMeeting_Minute = mySet_tb_meetings("Meeting_Minute")
			myMeeting_Length = mySet_tb_meetings("Meeting_Length")
			myMeeting_Length_In_Minutes	= mySet_tb_meetings("Meeting_Length_In_Minutes")
			myMeeting_Place	= mySet_tb_meetings("Meeting_Place")
			myMeeting_Agenda = mySet_tb_meetings("Meeting_Agenda")
			myMeeting_Comments = mySet_tb_meetings("Meeting_Comments")
			myMeeting_Date_Update = mySet_tb_meetings("Meeting_Date_Update")
			myMeeting_Author_Update	= mySet_tb_meetings("Meeting_Author_Update")

			myMeeting_Public = mySet_tb_Meetings("Meeting_Public")

		else
	
			' Close Recordset
			mySet_tb_meetings.close
			Set mySet_tb_meetings=nothing
			'Close Connection
			myConnection.close
			set myConnection = nothing
			URL="__Agenda_Day.asp?Date_Agenda="&myStrDate_Agenda
			Response.Redirect(URL)
		end if


	' Close Recordset
	mySet_tb_meetings.close
	Set mySet_tb_meetings=nothing
	'Close Connection
	myConnection.close
	set myConnection = nothing

	end if

end if

	
%> 
<TD BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left" WIDTH="<%=myApplication_Width%>"> 
<TABLE CELLSPACING="1" BGCOLOR="<%=myBGColor%>" WIDTH="<%=myApplication_Width%>">
<TR> 
<TD colspan=2> 
<%
myBox_title=myApplication_Title
%>

<!-- #include file="__Agenda_box.asp" -->
</td>
</tr>


 <FORM METHOD="POST" NAME="myform" ACTION="<%=myPage%>">

<%
' APPLICATION TITLE AND HIDDEN FIELDS
%>
 
<TR ALIGN="center">
<TD COLSPAN="2" BGCOLOR="<%=myApplicationColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR="<%=myApplicationTextColor%>"><B><%=myTitle%></B></FONT>

<INPUT TYPE="hidden" NAME="Meeting_ID" VALUE="<%=myMeeting_ID%>">
<INPUT TYPE="hidden" NAME="Action" VALUE="<%=myAction%>"> 
<INPUT TYPE="hidden" NAME="Date_Agenda" VALUE="<%=myStrDate_Agenda%>">
</TD>
</TR> 


<%
' Meeting Title
%>

<TR>
<TD ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><B><%=myMessage_Title%> 
*<BR><%=myFormGetErrMsg("Meeting_Title")%></B></FONT>
</TD>
<TD ALIGN="left"> 
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<input type=hidden name=location value="<%=myLocation%>">
<INPUT TYPE="text" SIZE="60" NAME="Meeting_Title" VALUE="<%=myMeeting_Title%>"> 
</FONT> 
</TD>
</TR>

<%
' Date
%>

<TR> 
<TD ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" Size="2" Color="<%=myBorderTextColor%>"><B><%=myMessage_Date%> 
*</B></FONT>
</TD>
<TD ALIGN="left" valign="middle">
<P>
<%
If myDate_Format = 1 Then
%>
<SELECT NAME="Month"> 
<OPTION VALUE="0" <%if my_Month=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if my_Month=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if my_Month=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if my_Month=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if my_Month=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if my_Month=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if my_Month=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if my_Month=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if my_Month=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if my_Month=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if my_Month=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if my_Month=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if my_Month=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>,
<%end if%>
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBGTextColor%>"> 
 <%=myFormGetErrMsg("Month")%>&nbsp; 
<INPUT TYPE="text" SIZE="2" MAXLENGTH="2" NAME="Day" VALUE="<%=my_Day%>">&nbsp;<%=myFormGetErrMsg("Day")%>, 
<%
If myDate_Format <>1 Then
%>
&nbsp;<SELECT NAME="Month"> 
<OPTION VALUE="0" <%if my_Month=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if my_Month=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if my_Month=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if my_Month=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if my_Month=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if my_Month=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if my_Month=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if my_Month=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if my_Month=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if my_Month=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if my_Month=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if my_Month=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if my_Month=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>,
<%end if%>

&nbsp;<INPUT TYPE="text" SIZE="4" MAXLENGTH="4" NAME="Year" VALUE="<%=my_Year%>">&nbsp;<%=myFormGetErrMsg("Year")%>
<%if not isDate(myMeeting_Date_Beginning) then %>
	&nbsp;<%=myError_Message_Invalid_Date%>
<%end if %>
&nbsp;
</FONT>
</B>
</P>

<P><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> <%=myMessage_At%>: 

<SELECT NAME="Meeting_Hour" >
<%

if myHour_Format <>1 Then
i=myAgenda_Start
Do While i <= myAgenda_End
	%>
	<OPTION <%if i=Cint(myMeeting_Hour) then%>selected<%end if%> VALUE="<%=i%>">
	<%
	If i< 10 Then 
		response.write "0" & i
	else 
		response.write i
	end if	
	%>h</OPTION> 
	<%
	i=i+1
loop
%> 
<%else
i=1
if myMeeting_Hour > 11 Then 
 If myMeeting_Hour >12 Then myMeeting_Hour = myMeeting_Hour - 12
 myHour_Indicator = 2
else 
 myHour_indicator = 1 
end if  

Do While i<13
	%>
	<OPTION <%if i=Cint(myMeeting_Hour) then%>selected<%end if%> VALUE="<%=i%>">
	<%=i%>
	</OPTION> 
	<%
	i=i+1
loop
%> 
<%end if%>
</select>
&nbsp;<b></b>&nbsp;

<SELECT NAME="myMeeting_Minute">
<%
i=0
Do While i<60
	%> 
	<OPTION <%if i=Cint(myMeeting_Minute) then%>selected<%end if%> VALUE="<%=i%>">:<%if i=0 then%>0<%end if%><%=i%></OPTION> 
	<%
	i=i+15
loop
%>
</SELECT>
<%if myHour_Format =1 Then%>
<select name="Hour_Indicator">
  <option <% if myHour_indicator = 1 then%> selected<%end if%> value="1">A.M</option>
  <option <% if myHour_indicator = 2 then%> selected<%end if%> value="2">P.M</option>
<%end if%>
</select>

 <%=myMessage_Min%>
 </FONT>
<FONT FACE="Arial,Helvetica" SIZE="-1"><%=myMessage_During%>: 

<SELECT NAME="Meeting_Length">
<%
i=0
Do While i<12
	%> 
	<OPTION <%if i=Cint(myMeeting_Length) then%>selected<%end if%> VALUE="<%=i%>"><%=i%> 
h</OPTION>
	<%
	i=i+1
loop
%>
</SELECT>

<SELECT NAME="Meeting_Length_In_Minutes">
<%
i=0
	Do While i<60
	%> 
	<OPTION <%if i=Cint(myMeeting_Length_In_Minutes) then%>selected<%end if%> VALUE="<%=i%>">:<%if i=0 then%>0<%end if%><%=i%></OPTION> 
	<%
	i=i+15
loop
%>
</SELECT>
 <%=myMessage_Min%>
</FONT></P>
</TD>
</TR>


<%
' Projects
%>

<%
if myUser_type_ID<=myProjects_Public_Type_ID then
%> 
	<TR>
	<TD ALIGN="right" BGCOLOR="<%=myBorderColor%>">
	 <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><B><%=myMessage_Project%></B></FONT>
	</TD>
	<TD ALIGN="left"> 
	<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
	<SELECT NAME="Project_ID" SIZE="1" TABINDEX="1">
	<OPTION VALUE="0" <%if myMeeting_Project_ID=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
	<%
	'Projects List
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String
	' Read tb_projects
	mySQL_Select_tb_Projects = "SELECT tb_projects.*,  tb_sites_Members.Member_Pseudo as Project_Leader_Pseudo FROM tb_Projects INNER JOIN tb_Sites_Members on tb_Projects.Project_leader_ID=tb_Sites_Members.Member_ID WHERE tb_Projects.Site_ID ="& mySite_ID

	set mySet_tb_Projects = myConnection.Execute(mySQL_Select_tb_Projects)
	do while not mySet_tb_Projects.eof
		myProject_ID=mySet_tb_Projects("Project_ID")
		myProject_Name=mySet_tb_Projects("Project_Name") 
		%> 
		<OPTION VALUE="<%=myProject_ID%>" <%if 	myProject_ID=Cint(myMeeting_Project_ID) then%>Selected<%end if%>><%=myProject_Name%></OPTION> 
		<%
		mySet_tb_Projects.movenext
	loop
	mySet_tb_Projects.close
	Set mySet_tb_Projects=nothing
	myConnection.close
	set myConnection = nothing
	%> 
	</SELECT>
	</FONT>
	</TD>
	</TR>
<%
end If
%>


<%
' Place
%>

<TR>
<TD ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" Size="2" Color="<%=myBorderTextColor%>"><B><%=myMessage_Place%> 
</B></FONT>
</TD>
<TD ALIGN="left">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> 
<INPUT TYPE="text" SIZE="60" NAME="Meeting_Place" VALUE="<%=myMeeting_Place%>"> 
</FONT>
</TD>
</TR>


<%
' Agenda
%>

<TR>
<TD VALIGN="top" ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" Size="2" Color="<%=myBorderTextColor%>"><B><%=myMessage_Meeting_Agenda%></B></FONT>
</TD>
<TD ALIGN="left"> 
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<TEXTAREA ROWS="6" COLS="60" NAME="Meeting_Agenda"><%=myMeeting_Agenda%></TEXTAREA> 
</FONT>
</TD>
</TR>



<%
' PARTICIPANTS
%> 

<TR>
<TD ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><B><%=myMessage_Participants%></B></FONT>
</TD>
<TD ALIGN="left" bgcolor="<%=myBGColor%>">
<table border="0" cellpadding="5" cellspacing="0"> 
<%
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Sites_Members ="  Select * from tb_Sites_Members  WHERE  Site_ID="&mySite_ID &" ORDER BY Member_Pseudo"
set mySet_tb_Sites_Members = myConnection.Execute(mySQL_Select_tb_Sites_Members)

' each 3 members CRLF
myCounter = 1
do while not mySet_tb_Sites_Members.eof
			
	myParticipant_ID     = mySet_tb_Sites_Members("Member_ID")
	myParticipant_Pseudo = mySet_tb_Sites_Members("Member_Pseudo")
		
	if myCounter = 1 then
			Response.Write "<tr bgcolor=#ffffff>"
	end if
			
	mySQL_Select_tb_Meetings_Members = "Select * from tb_Meetings_Members WHERE  Site_ID="&mySite_ID&" AND Member_ID="&myParticipant_ID&" AND Meeting_Role_ID=2 AND Meeting_ID="&myMeeting_ID
	set mySet_tb_Meetings_Members = myConnection.Execute(mySQL_Select_tb_Meetings_Members)
%> 

	<td valign="top"  bgcolor="<%=myBGColor%>">
	<small>
	<input type="checkbox" name="<%=myParticipant_Pseudo%>"
	<%If myAction = "New" and myParticipant_Pseudo = myUser_Login then response.Write " checked " %>

	<%if not mySet_tb_Meetings_Members.eof then%> value="on" checked<%end if%>> 
	<% 
	if myUser_type_ID <= myMembers_Public_type_ID then
		%>
		<A HREF="__Site_Member_Information.asp?Member_ID=<%=mySet_tb_Sites_Members("Member_ID")%>&Site_ID=<%=myMeeting_Site_ID%>"><Font face="Arial, Helvetica, sans-serif" size="2"><%=myParticipant_Pseudo%></font></a>
		<%
	else
		%>
		<font face="Arial, Helvetica, sans-serif" size="2"><%=myParticipant_Pseudo%></font>
		<%
	end if
	%>
	</small>
	</td>
	<%
	if myCounter = 3 then
		Response.Write "</tr>"
		myCounter = 1
	else
		myCounter= myCounter + 1
	end if
			
	mySet_tb_Sites_Members.movenext
loop

if myCounter = 3 then
		Response.Write "<td bgcolor="& myBGColor &">&nbsp; </td><td  bgcolor="& myBGColor &">&nbsp; </td></tr>"
elseif myCounter = 4 then
		Response.Write "<td  bgcolor="& myBGColor &">&nbsp; </td></tr>"
end if
	
%> 
</table>

<%
mySet_tb_Sites_Members.close
Set mySet_tb_Sites_Members=Nothing
mySet_tb_Meetings_Members.close
Set mySet_tb_Meetings_Members=Nothing
myConnection.close
set myConnection = nothing
%> 
     

<%
' Comments
%>

<TR>
<TD VALIGN="top" ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><b><%=myMessage_Comments%></B></FONT>
</TD>
<TD ALIGN="left"> 
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<TEXTAREA ROWS="6" COLS="60" NAME="Meeting_Comments"><%=myMeeting_Comments%></TEXTAREA> 
</FONT>
</TD>
</TR>

<%
' Confidentiality
%>

<TR>
<TD VALIGN="top" ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><b><%= myMessage_Meeting_Confidentiality %></B></FONT>
</TD>
<TD ALIGN="left"> 
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<input type="radio" name="Meeting_public" value="0" <%If myMeeting_Public=0 or len(myMeeting_Public) = 0 Then response.write "checked" %> > <%= my_File_Message_Pulic_list %><br>
<input type="radio" name="Meeting_public" value="1" <%If myMeeting_Public = 1 Then response.write "checked" %>> <%= my_File_Message_Private %>

</FONT>
</TD>
</TR>



<%
'Validation
%>

<TR> <TD VALIGN="top" ALIGN="right" BGCOLOR="<%=myBorderColor%>" WIDTH="92"><FONT FACE="Arial, Helvetica, sans-serif" Size="1" Color="<%=myBorderTextColor%>">* = <%=myMessage_Required%></FONT>
</TD>
<TD ALIGN="left"> 
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<INPUT TYPE="submit" VALUE="<%=myMessage_Go%>" NAME="Validation"> 
</FONT>
</TD>
</TR>

<%
' Date and Author
%>


<TR ALIGN="CENTER">
<TD VALIGN="top" BGCOLOR="<%=myApplicationColor%>" COLSPAN="2"> 
<%'if myAction="Update" then%> 
	<b><font face="Arial, Helvetica, sans-serif" size="1" color="<%=myApplicationTextColor%>"><%=myDate_Display(myMeeting_Date_Update,2) %> -- <%=myMeeting_Author_Update%></font></b> 
<%'end if%>
</TD>
</TR>

</FORM>

</TABLE>

<%if(myMeeting_Member_ID=MyUser_ID or myUser_Type_ID=1) and myAction="Update" then%>
	<table border="0" width="468" cellpadding="3" cellspacing="0"> 
	<tr>
	<td>
	<a href="Javascript:if(confirm('<%=myMessage_Delete%>  ?'))document.location='__Agenda_Modification.asp?Meeting_ID=<%=myMeeting_ID%>&amp;Date_Agenda=<%=myStrDate_Agenda%>&amp;action=Delete';"><FONT FACE="Arial, Helvetica, sans-serif" Size="2"><%=myMessage_Delete%></FONT></a>
	</td>
	</tr>
	</table>
<%end if%>

</TD>
</TR>
</TABLE>

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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</FONT>
</TD>
</TR>
</TABLE>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright												'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</body>

</html>


<html><script language="JavaScript"></script></html>