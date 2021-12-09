<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   z OverApps z http://www.overapps.com
'
' This program "__Event_modification.asp" is free software; 
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
' Does n't Work With PWS
%>

<%
' ------------------------------------------------------------
'
' Name 			: __Event_modification.asp
' PAth    		: /
' Version 		: 1.15.0
' Description 	: Add, Modify, Delete Events
' by 			: Pierre Rouarch
' Company		: OverApps
' Date 			: December, 10, 2001
' Contributions : Jean-Luc Lesueur , Christophe Humbert, Dania Tcherkezoff 
' Modify by 	:
' Company		:
' Date			:
' ------------------------------------------------------------

Dim myPage
myPage = "__Event_modification.asp"

Dim myPage_Application
myPage_Application="Events"
	
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

Dim myEvent_Site_ID, myEvent_Member_ID, myEvent_Calendar_ID, myEvent_Calendar_Name, myEvent_ID, myEvent_Name, myEvent_Presentation,  myEvent_Date_Beginning, myEvent_Date_End,   myEvent_Author_Update, myEvent_Date_Update
Dim myStrEvent_Date_Beginning, myStrEvent_Date_End
Dim myDay_Beginning, myMonth_Beginning, myYear_Beginning, myDay_End, myMonth_End, myYear_End


Dim myAction, myList, myURL

Dim mySQL_Select_tb_Events, mySQL_Delete_tb_Events, mySQL_Insert_tb_Events, mySQL_Update_tb_Events, mySet_tb_Events

Dim mySQL_Select_tb_Calendars,  mySet_tb_Calendars

' Get Parameters

myAction = Request("Action")
myList = Request("List")
myEvent_ID  = Request("Event_ID")
if len(myEvent_ID)=0 then
	myEvent_ID=0
	myAction="New"
end if
myEvent_ID=Cint(myEvent_ID)
if myAction="New" then
	myEvent_Date_Beginning=Now()
	myDay_Beginning	 = Day(myEvent_Date_Beginning)
	myMonth_Beginning	 = Month(myEvent_Date_Beginning)
	myYear_Beginning	 = Year(myEvent_Date_Beginning)
	myStrEvent_Date_Beginning= myDate_Construct(myYear_Beginning,myMonth_Beginning,myDay_Beginning,23,59,59)
	myEvent_Date_End    = myDate_Now()
	myDay_End	 = Day(myEvent_Date_End)
	myMonth_End	 = Month(myEvent_Date_End)
	myYear_End	 = Year(myEvent_Date_End)
	myStrEvent_Date_End= myDate_Construct(myYear_End,myMonth_End,myDay_End,23,59,59)
end if



' Not Used Force to 1
'myEvent_Calendar_ID=Request.Form("Calendar_ID")
myEvent_Calendar_ID=1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Delete
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction = "Delete" then
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String
	mySQL_Delete_tb_Events = "DELETE FROM tb_Events WHERE Site_ID = "&mySite_ID&" AND Event_ID = "&myEvent_ID
	myConnection.Execute(mySQL_Delete_tb_Events)
	myConnection.Close
	set myConnection = Nothing
	' and go back
	Response.Redirect("__Events_List.asp")
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.form("Validation")=myMessage_Go then

	' Get Entries
	myEvent_Name    = Replace(Request.Form("Event_Name"),"'"," ")
	myEvent_Presentation   = Replace(Request.Form("Event_Presentation"),"'"," ")

	myDay_Beginning=Request.Form("Day_Beginning")
	myMonth_Beginning=Request.Form("Month_Beginning")
	myYear_Beginning=Request.Form("Year_Beginning")
	if (myDay_Beginning<> "" and myMonth_Beginning <> "" and myYear_Beginning <> "") then
		myEvent_Date_Beginning = myDate_Construct(myYear_Beginning,myMonth_Beginning,myDay_Beginning,23,59,59)
	end if

	myDay_End=Request.Form("Day_End")
	myMonth_End=Request.Form("Month_End")
	myYear_End=Request.Form("Year_End")
	if (myDay_End<> "" and myMonth_End <> "" and myYear_End <> "") then
		myEvent_Date_End = myDAte_Construct(myYear_End,myMonth_End,myDay_End,23,59,59)
	end if

	' Test Entries
	Call myFormSetEntriesInString
	myFormCheckEntry null, "Event_Name",true,null,null,0,255
	myFormCheckEntry null, "Event_Presentation",false,null,null,0,255
	myFormCheckEntry myErr_numerical, "Day_Beginning",true,1,31,0,2
	myFormCheckEntry myErr_numerical, "Month_Beginning",true,1,12,0,2
	myFormCheckEntry myErr_numerical, "Year_Beginning",true,1990,2050,0,4
	myFormCheckEntry myErr_numerical, "Day_End",true,1,31,0,2
	myFormCheckEntry myErr_numerical, "Month_End",true,1,12,0,2
	myFormCheckEntry myErr_numerical, "Year_End",true,1990,2050,0,4

	if not myform_entry_error  then

		myStrEvent_Date_Beginning=myEvent_Date_Beginning
	    myStrEvent_Date_End=myEvent_Date_End

		' DataBase Connection
		set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String

		myEvent_Author_Update   = myUser_Pseudo
		myEvent_Date_Update     = myDate_Now()

		if myAction = "New" then

			' Insert
			mySQL_Select_tb_Events = "SELECT * FROM tb_Events"
			Set mySet_tb_Events = server.createobject("adodb.recordset")
			mySet_tb_Events.open mySQL_Select_tb_Events, myConnection, 3, 3
			mySet_tb_Events.AddNew

			mySet_tb_Events.fields("Site_ID")=mySite_ID
			mySet_tb_Events.fields("Member_ID")=myUser_ID
			mySet_tb_Events.fields("Calendar_ID")=myEvent_Calendar_ID
			mySet_tb_Events.fields("Event_Name")=myEvent_Name
			mySet_tb_Events.fields("Event_Presentation")=myEvent_Presentation
			mySet_tb_Events.fields("Event_Date_Beginning")=myStrEvent_Date_Beginning
			mySet_tb_Events.fields("Event_Date_End")=myStrEvent_Date_End
			mySet_tb_Events.fields("Event_Date_Update")=myEvent_Date_Update
			mySet_tb_Events.fields("Event_Author_Update")=myEvent_Author_Update
	
			mySet_tb_Events.Update
			' Close Recordset 
			mySet_tb_Events.close
			Set mySet_tb_Events = Nothing

		elseif myAction = "Update" then

			mySQL_Select_tb_Events = "SELECT * FROM tb_Events WHERE Event_ID =" & myEvent_ID
			Set mySet_tb_Events = server.createobject("adodb.recordset")
			mySet_tb_Events.open mySQL_Select_tb_Events, myConnection, 3, 3
	
			mySet_tb_Events.fields("Event_Name")=myEvent_Name
			mySet_tb_Events.fields("Event_Presentation")=myEvent_Presentation
			mySet_tb_Events.fields("Event_Date_Beginning")=myStrEvent_Date_Beginning
			mySet_tb_Events.fields("Event_Date_End")=myStrEvent_Date_End
			mySet_tb_Events.fields("Event_Date_Update")=myEvent_Date_Update
			mySet_tb_Events.fields("Event_Author_Update")=myEvent_Author_Update
	
			mySet_tb_Events.Update
			' Close Recordset 
			mySet_tb_Events.close
			Set mySet_tb_Events = Nothing

		end if ' Insert or Update


		' Close Connection
		myConnection.close
		set myConnection = nothing
		' and go back

		myURL="__Events_List.asp?List="&myList
		Response.Redirect(myURL)



	end if ' no error

end if ' Validation


%>
<html>

<head>
<title><%=mySite_Name%> Event Add/Modify/Delete</title>

</head>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%> <!-- #include file="_borders/Top.asp" --> <%
' CENTER
%> <TABLE WIDTH="<%=myGlobal_Width%>" bgColor="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> <%
' CENTER LEFT
%> <TD WIDTH="<%=myLeft_Width%>"> <!-- #include file="_borders/Left.asp" --> </td><%
' CENTER APPLICATION
%> <%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form					                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myAction="Update" And Request.form("Validation")<>myMessage_Go then


' db Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String



	' in Update Way

	if myAction = "Update" then


		' Get  Event's informations
		mySQL_Select_tb_Events = "SELECT * FROM tb_Events WHERE Site_ID = "&mySite_ID&" AND Event_ID = " & myEvent_ID
		set mySet_tb_Events = myConnection.Execute(mySQL_Select_tb_Events)


		if not mySet_tb_Events.eof then


		myEvent_Site_ID = mySet_tb_Events("Site_ID")
		myEvent_Member_ID = mySet_tb_Events("Member_ID")
 		myEvent_Calendar_ID=mySet_tb_Events("Calendar_ID")
		myEvent_ID = mySet_tb_Events("Event_ID")
		myEvent_Name  = mySet_tb_Events("Event_Name")
		myEvent_Presentation  = mySet_tb_Events("Event_Presentation")
		myEvent_Date_Beginning  = mySet_tb_Events("Event_Date_Beginning")
		myDay_Beginning	 = myDay(myEvent_Date_Beginning)
		myMonth_Beginning	 = myMonth(myEvent_Date_Beginning)
		myYear_Beginning	 = myYear(myEvent_Date_Beginning)
		myStrEvent_Date_Beginning=myYear(myEvent_Date_Beginning)&"/"&myMonth(myEvent_Date_Beginning)&"/"&myDay(myEvent_Date_Beginning)

		myEvent_Date_End    = mySet_tb_Events("Event_Date_End")
		myDay_End	 = myDay(myEvent_Date_End)
		myMonth_End	 = myMonth(myEvent_Date_End)
		myYear_End	 = myYear(myEvent_Date_End)
		myStrEvent_Date_End=myYear(myEvent_Date_End)&"/"&myMonth(myEvent_Date_End)&"/"&myDay(myEvent_Date_End)

		myEvent_Author_Update   = mySet_tb_Events("Event_Author_Update")
		myEvent_Date_Update     = mySet_tb_Events("Event_Date_Update")

		else
			'Close Connection
			myConnection.close
			set myConnection = nothing
			Response.Redirect("__Events_List.asp")
		end if


	end if


	' Close Connection
	myConnection.close
	set myConnection = nothing

end if

%> 

<td valign="top" Width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" > 
<form method="POST" action="<%=myPage%>" name="myForm">


<table WIDTH="<%=myApplication_Width%>" border="0" cellpadding="5" cellspacing="1" BGCOLOR="<%=myBGColor%>">

<% 
' Application Title and hidden fields
%>


<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">

<font color="<%=myApplicationTextColor%>" face="Arial, Helvetica, sans-serif" size="4"><b><%=myApplication_Title%></b></font>
<INPUT TYPE="hidden" NAME="Calendar_ID" VALUE="<%=myEvent_Calendar_ID%>">
<INPUT TYPE="hidden" NAME="Event_ID" VALUE="<%=myEvent_ID%>"> 
<INPUT TYPE="hidden" NAME="Action" VALUE="<%=myAction%>">
<INPUT TYPE="hidden" NAME="List" VALUE="<%=myList%>">
</td>
</tr>

<%
' Event 's Name
%>
 
<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Name%>*<BR> 
<%=myFormGetErrMsg("Event_Name")%></FONT></B>
</td>
<td align="left"  valign="top"> 
<input type="text" size="60" name="Event_Name" value="<%=myEvent_Name%>" MAXLENGTH="100"> 
</td>
</tr>

<%
' Event's Presenatation
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><font size="2" face="Arial, Helvetica, sans-serif" COLOR="<%=myBorderTextColor%>"><%=myMessage_Presentation%> 
<BR><%=myFormGetErrMsg("Event_Presentation")%></FONT></B>
</td>
<td align="left"  valign="top">
<TEXTAREA COLS="55" NAME="Event_Presentation" ROWS="5"><%=myEvent_Presentation%></TEXTAREA> 
</td>
</tr>

<%
' Date Beginnning
%>

<tr>
<td align="right"  bgcolor="<%=myBorderColor%>">
<B><font size="2" face="Arial, Helvetica, sans-serif" COLOR="<%=myBorderTextColor%>"><%=myMessage_Beginning%>*
<br><%=myFormGetErrMsg("Day_Beginning")%>
<br><%=myFormGetErrMsg("Month_Beginning")%>
<br><%=myFormGetErrMsg("Year_Beginning")%></FONT></B>
</td>
<td>
<P>
<% If myDate_Format = 1 Then %>
<SELECT NAME="Month_Beginning">
<OPTION VALUE="0" <%if myMonth_Beginning=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMonth_Beginning=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if myMonth_Beginning=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if myMonth_Beginning=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if myMonth_Beginning=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if myMonth_Beginning=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if myMonth_Beginning=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if myMonth_Beginning=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if myMonth_Beginning=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if myMonth_Beginning=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if myMonth_Beginning=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if myMonth_Beginning=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if myMonth_Beginning=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>,
&nbsp;
<%end if%>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<INPUT TYPE="text" SIZE="2" MAXLENGTH="2" NAME="Day_Beginning" VALUE="<%=myDay_Beginning%>"> , 
 
<% If myDate_Format <> 1 Then %>
<SELECT NAME="Month_Beginning">
<OPTION VALUE="0" <%if myMonth_Beginning=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMonth_Beginning=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if myMonth_Beginning=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if myMonth_Beginning=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if myMonth_Beginning=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if myMonth_Beginning=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if myMonth_Beginning=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if myMonth_Beginning=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if myMonth_Beginning=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if myMonth_Beginning=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if myMonth_Beginning=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if myMonth_Beginning=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if myMonth_Beginning=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>,
&nbsp;
<%end if%>
<INPUT TYPE="text" SIZE="4" MAXLENGTH="4" NAME="Year_Beginning" VALUE="<%=myYear_Beginning%>">
</FONT>
</P>
&nbsp;
&nbsp;

</td>
</tr>


<%
' Date End
%>

<tr>
<td align="right"  bgcolor="<%=myBorderColor%>">
<B><font size="2" face="Arial, Helvetica, sans-serif" COLOR="<%=myBorderTextColor%>"><%=myMessage_End%>*<br> 
<%=myFormGetErrMsg("Day_End")%><br>
<%=myFormGetErrMsg("Month_End")%><br>
<%=myFormGetErrMsg("Year_End")%></FONT></B>
</td>
<td> 
<P>
<% If myDate_Format = 1 Then %>
<SELECT NAME="Month_End">
<OPTION VALUE="0" <%if myMonth_End=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMonth_End=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if myMonth_End=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if myMonth_End=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if myMonth_End=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if myMonth_End=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if myMonth_End=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if myMonth_End=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if myMonth_End=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if myMonth_End=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if myMonth_End=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if myMonth_End=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if myMonth_End=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>,
&nbsp;
<%end if%>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<INPUT TYPE="text" SIZE="2" MAXLENGTH="2" NAME="Day_End" VALUE="<%=myDay_End%>"> , 
<% If myDate_Format <> 1 Then %>
<SELECT NAME="Month_End">
<OPTION VALUE="0" <%if myMonth_End=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMonth_End=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if myMonth_End=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if myMonth_End=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if myMonth_End=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if myMonth_End=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if myMonth_End=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if myMonth_End=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if myMonth_End=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if myMonth_End=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if myMonth_End=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if myMonth_End=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if myMonth_End=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>,
&nbsp;
<%end if%>
<INPUT TYPE="text" SIZE="4" MAXLENGTH="4" NAME="Year_End" VALUE="<%=myYear_End%>">
</FONT></P> &nbsp;&nbsp; 

</td>
</tr>

<%
' Validation
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" Color="<%=myBorderTextColor%>">* = <%=myMessage_Required%></FONT></B>
</td>
<td valign="top" align="left"> 
<input type="submit" value="<%=myMessage_Go%>" name="Validation">
</td>
</tr>

<% 
' Date - Author
%>
<TR>
<TD VALIGN="top" ALIGN="right" COLSPAN="2" BGCOLOR="<%=myApplicationColor%>"> 
<P ALIGN="center"><B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" Color="<%=myApplicationTextColor%>"><%=myDAte_Display(myEvent_Date_Update,2)%> -- <%=myEvent_Author_Update%></FONT></B></P>
</TD>
</TR> 
</table>
</form>

<%
' ADMINISTRATION - EveryBody Can Add, Delete or Modify in this Version
%> 

<% 
If myAction="Update" then
 %> 
	<TABLE BORDER="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="3" CELLSPACING="0"> 
	<TR>
	<TD>
	&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><A HREF="__Event_Modification.asp?Calendar_ID=<%=myEvent_Calendar_ID%>"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Add%>&nbsp;<%=myMessage_Event%></font></A></FONT>
	</TD>
	</TR>
	<TR>
	<TD>
	&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><A HREF="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Event_modification.asp?Action=Delete&amp;Event_ID=<%=myEvent_ID%>';"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Delete%></font></A></FONT>
	</TD>
	</TR>
	</TABLE>
<% 
End If 
%>

</td>
</TR>
</TABLE>

<!-- #include file="_borders/Down.asp" --> 

<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.														'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 

<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0">
<TR ALIGN="RIGHT">
<TD>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> 
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</FONT>
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
<html></html>