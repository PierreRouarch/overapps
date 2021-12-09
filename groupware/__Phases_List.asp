<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - O.v.e.r.A.p.p.s. - http://www.overapps.com
'
' This program "__Phases_List.asp" is free software; 
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
' Does n't Work with PWS ????
%>

<%
' ------------------------------------------------------------
' Name	 		: __Phases_List.asp
' Path			: /
' Version 		: 1.15.0
' Description 	: Graphical Presentation of a project
'
' By			: Pierre Rouarch
' Company		: OverApps
' Date			: March, 27, 2001
' 
' Modify by		:	
' Modifications :
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Phases_List.asp"

Dim myPage_Application
myPage_Application="Projects"
	
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

' Local Variables

Dim myDay_Width, myToday,  myIntMonth, myMonthName, myNumber_of_Months, myDate, myIntYear, myIntYear1, myNumber_of_Days, myStrYear

Dim myProject_Site_ID, myProject_Member_ID, myProject_ID,  myProject_Name, myProject_Date_Beginning, myDate_Beginning, myProject_Date_End, myDate_End, myProject_Presentation, myProject_Leader_ID, myProject_Author_Update, myProject_Date_Update

Dim myDate_Beginning_NumDay_in_Month
Dim myDate_End_Next_Month
Dim myDate_Beginning_First_Day

Dim myPhase_Site_ID, myPhase_Member_ID, myPhase_ID, myPhase_Name, myPhase_Date_Beginning, myPhase_Date_End, myPhase_Period_Beginning,  myPhase_Date_Beginning2, myPhase_Date_End2, myPhase_Period_Beginning2, myPhase_Leader_ID

Dim myPhase_Length, myPhase_Length2, myProject_Day_Beginning

Dim myPhase_Period_Beginning_in_Months, myPhase_Length_in_Months, myPhase_Period_Beginning2_in_Months, myPhase_Length2_in_Months

Dim mySQL_Select_tb_Projects, mySet_tb_Projects, mySQL_Select_tb_Phases, mySet_tb_Phases, mySQL_Select_tb_Phases_SubPhases, mySet_tb_Phases_SubPhases

' FOR PROJECTS 'NAVIGATION
Dim myList, myNumPage, mySearch

' Function GetLastDay

Dim intMonthNum, intYearNum, datTestDay
function GetLastDay(datTheDate)
	intMonthNum = Month(datTheDate)
	intYearNum = Year(datTheDate)
	datTestDay = DateSerial(intYearNum, intMonthNum + 1,1) - 1
	GetLastDay = Day(datTestDay)
end function

' Get Parameters
myProject_ID=request("Project_ID")
if  Len(myProject_ID)=0 then
	Response.Redirect("__Projects_List.asp")
end if
	

%>

<html>

<head>

<title><%=mySite_Name%> - Phases List - Planning </title>
</head>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' Top
%> 

<!-- #include file="_borders/Top.asp" -->

<%
' Center
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR>
<TD VALIGN="TOP">

<%
' CENTER LEFT
%> 

<TD WIDTH="<%=myLeft_Width%>" bgcolor="<%=myBorderColor%>">

<!-- #include file="_borders/Left.asp" -->

</td>

<%
' CENTER Application
%>

<%

' Db Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


' Day Pixel's length
myDay_Width=1
' Today
myToday = Date()

' Project' Information
mySQL_Select_tb_Projects = "SELECT * FROM tb_Projects WHERE Project_ID="&myProject_ID
set mySet_tb_Projects = myConnection.execute(mySQL_Select_tb_Projects)

if not mySet_tb_Projects.eof then
	myProject_Site_ID=mySet_tb_Projects("Site_ID")
	myProject_Member_ID=mySet_tb_Projects("Member_ID")
	myProject_ID     = mySet_tb_Projects("Project_ID")
	myProject_Member_ID=mySet_tb_Projects("Member_ID")
	myProject_Name   = mySet_tb_Projects("Project_Name")
	myProject_Leader_ID= mySet_tb_Projects("Project_Leader_ID")
	myProject_Date_Beginning = CDate(mySet_tb_Projects("Project_Date_Beginning"))
	myProject_Day_Beginning = Day(myProject_Date_Beginning)
	
	myDate_Beginning                    = DateSerial(myYear(mySet_tb_Projects("Project_Date_Beginning")),myMonth(mySet_tb_Projects("Project_Date_Beginning")),myDay(mySet_tb_Projects("Project_Date_Beginning")))
	myDate_Beginning_First_Day = DateSerial(myYear(mySet_tb_Projects("Project_Date_Beginning")),myMonth(mySet_tb_Projects("Project_Date_Beginning")),1)
	
	myProject_Date_End   = CDate(mySet_tb_Projects("Project_Date_End"))
	myDate_End = mySet_tb_Projects("Project_Date_End")
	
	myProject_Presentation	  = mySet_tb_Projects("Project_Presentation")
	if len(myProject_Presentation) > 0 then
			myProject_Presentation = Replace(myProject_Presentation,vbCrLf,"<br>")
	end if
	myProject_Author_Update	  = mySet_tb_Projects("Project_Author_Update")
	myProject_Date_Update	  = mySet_tb_Projects("Project_Date_Update")

else
	' Close Recordset
	mySet_tb_Projects.close
	Set mySet_tb_Projects=Nothing
	' Close Connection 
	myConnection.close
	set myConnection = nothing
	' And Go Back
	Response.Redirect ("__Projects_List.asp")
end if

' Read Phases
mySQL_Select_tb_Phases = "SELECT * FROM tb_phases WHERE (Project_ID = " & myProject_ID & " AND Phase_Parent_ID = 0) ORDER BY Phase_Date_Beginning "
set mySet_tb_Phases = myConnection.execute(mySQL_Select_tb_Phases)


%> 



<td align="left" bgcolor="<%=myBGColor%>" valign="top" WIDTH="<%=myApplication_Width%>"> 

<%
' HEADER
%>

<table border="0" cellpadding="3" cellspacing="0" WIDTH="<%=myApplication_Width%>"> 


<%
' TITLE
%>

<tr>
<td bgcolor="<%=myApplicationColor%>" align="Left" WIDTH="<%=myApplication_Width%>"> 
<font face="Arial, Helvetica, sans-serif" size="4"><b><font color="<%=myApplicationTextColor%>"> 
<%=myApplication_Title%>/<%=myProject_Name%></font></b></font>
</td>
</tr> 
</table>

<table border="0" cellpadding="3" cellspacing="1" WIDTH="<%=myApplication_Width%>"> 


<%
' Presentation
%>

<tr>
<td align="left" WIDTH="<%=myApplication_Width%>">
<font face="Arial, Helvetica, sans-serif" size="2"><b><%=myProject_Presentation%></b></font>
</td>
</tr>

</table>


<table border="1" cellpadding="0" cellspacing="0" bgcolor="<%=myBGColor%>"> 
<%
' Months
%> 

<tr> 
<td align="left" rowspan="2" WIDTH="<%=myLeft_Width%>" bgcolor="<%=myBorderColor%>" nowrap>&nbsp; 
</td>

<% 
' Months' names
myNumber_of_Months = 0
myIntMonth = Month(myDate_Beginning)
myIntYear1  = Year(myDate_Beginning)
myDate = myDate_Beginning

' Add a month 
myDate_End_Next_Month = DateAdd("m",1,myDate_End)

do while myDate < myDate_End_Next_Month
	myNumber_of_Months = myNumber_of_Months + 1
	' First three letters of the Month
	myIntMonth = Month(myDate)

If myIntMonth = 1 Then myMonthName = Left(myMessage_January,4)
If myIntMonth = 2 Then myMonthName = Left(myMessage_february,4)
If myIntMonth = 3 Then myMonthName = Left(myMessage_march,4)
If myIntMonth = 4 Then myMonthName = Left(myMessage_april,4)
If myIntMonth = 5 Then myMonthName = Left(myMessage_may,4)
If myIntMonth = 6 Then myMonthName = Left(myMessage_june,4)
If myIntMonth = 7 Then myMonthName = Left(myMessage_july,4)
If myIntMonth = 8 Then myMonthName = Left(myMessage_august,4)
If myIntMonth = 9 Then myMonthName = Left(myMessage_september,4)
If myIntMonth = 10 Then myMonthName = Left(myMessage_october,4)
If myIntMonth = 11 Then myMonthName = Left(myMessage_november,4)
If myIntMonth = 12 Then myMonthName = Left(myMessage_december,4)

	

	myNumber_of_Days = myDay_Width*GetLastDay(myDate) 
	if myIntYear = Year(myDate) then
			myStrYear = ""
	else
			myIntYear = Year(myDate)
			myStrYear = myIntYear
	end if
%> 

    <td valign="middle" align="center" height="20" width="<%=myNumber_of_Days%>" bgcolor="<%=myBorderColor%>"><img src="Images/OverApps-lenght_trans.gif" width="<%=myDay_Width*(myNumber_of_Days)%>" height="1"><b><font face="Arial, Helvetica, sans-serif" size="1" Color="<%=myBorderTextColor%>"> <%=myStrYear&"<br>"%>
	
	<%=myMonthName%>
	
	
	</font></b> </td>
<%
	' Next Month
	myDate = DateAdd("m",1,myDate)
loop 

%> 
</tr>
 <%
' Today's cursor
%> <tr><td align="left" height="20" colspan="<%=myNumber_of_Months %>"><% if myToday>=myDate_Beginning AND myToday=<myDate_End then %> 

<img src="Images/OverApps-lenght_trans.gif" width="<%=myDay_Width*((myToday-myDate_Beginning_First_Day)+((myToday-myDate_Beginning_First_Day)/30)+3)%>" height="6"><img src="Images/OverApps-arrow_down.gif" alt="<%=myToday%>" width="11" height="10"> 
<% end if %> &nbsp;</td></tr> 
<%
' PHASES 
%> 

<%	
do while not mySet_tb_Phases.eof 
	myPhase_ID = mySet_tb_Phases("Phase_ID")
	myPhase_Date_Beginning = DateSerial(myYear(mySet_tb_Phases("Phase_Date_Beginning")),myMonth(mySet_tb_Phases("Phase_Date_Beginning")),myDay(mySet_tb_Phases("Phase_Date_Beginning")))
	myPhase_Date_End  = DateSerial(myYear(mySet_tb_Phases("Phase_Date_End")),myMonth(mySet_tb_Phases("Phase_Date_End")),myDay(mySet_tb_Phases("Phase_Date_End")))
	
	myDate_Beginning_NumDay_in_Month = DatePart("d",myPhase_Date_Beginning)

	myPhase_Period_Beginning = myDay_Width*(myPhase_Date_Beginning - myDate_Beginning_First_Day)*(1+(1/15))

	myPhase_Length	= myDay_Width*(myPhase_Date_End+1 - myPhase_Date_Beginning)*(1+(1/15))

		
	' Corrected date
	myPhase_Date_Beginning2 = ""
	myPhase_Date_End2 = ""
	if len(mySet_tb_Phases("Phase_Date_Beginning2")) > 0 Then
	 myPhase_Date_Beginning2 = DateSerial(myYear(mySet_tb_Phases("Phase_Date_Beginning2")),myMonth(mySet_tb_Phases("Phase_Date_Beginning2")),MyDay(mySet_tb_Phases("Phase_Date_Beginning2")))
	 myPhase_Date_End2 = DateSerial(myYear(mySet_tb_Phases("Phase_Date_End2")),myMonth(mySet_tb_Phases("Phase_Date_End2")),MyDay(mySet_tb_Phases("Phase_Date_End2")))
    end if 
	if len(myPhase_Date_Beginning2)>0 AND len(myPhase_Date_End2)>0 then
		myPhase_Period_Beginning2=myDay_Width*(myPhase_Date_Beginning2 - myDate_Beginning_First_Day)*(1+(1/15))
		myPhase_Length2= myDay_Width*(myPhase_Date_End2+1 - myPhase_Date_Beginning2)*(1+(1/15))
	else
		myPhase_Period_Beginning2 = 0
		myPhase_Length2=0
	end if					

%> 

<tr>
<td height="30" align="left" bgcolor="<%=myBorderColor%>" VALIGN="MIDDLE" >
&nbsp;&nbsp;<a href="__Phase_Information.asp?Project_ID=<%=myProject_ID%>&Phase_ID=<%=mySet_tb_Phases("Phase_ID")%>"><font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><%=mySet_tb_Phases("Phase_Name")%></font></a>&nbsp;&nbsp;
</td>
<td height="30" align="left" valign="Middle" colspan="<%=myNumber_of_Months%>" bgcolor="#999999"> <font face="Arial, Helvetica, sans-serif" size="-7"><img src="Images/OverApps-lenght_trans.gif" width="<%=myPhase_Period_Beginning%>" height="6"><img src="Images/OverApps-lenght_scheduled.gif" width="<%=myPhase_Length%>" height="6" alt="<%=myPhase_Date_Beginning&"--"&myPhase_Date_End%>"><br> 
<% if (myPhase_Period_Beginning2>0 OR myPhase_Length2>0) then %><img src="Images/OverApps-lenght_trans.gif" width="<%=myPhase_Period_Beginning2%>" height="6"><img src="Images/OverApps-lenght_corrected.gif" width="<%=myPhase_Length2%>" height="6" alt="<%=myPhase_Date_Beginning2&"--"&myPhase_Date_End2%>"> 
<% end if %></font></td></tr><%
''''''''''''''''''''''''''''''
' Sub Phases
''''''''''''''''''''''''''''''
mySQL_Select_tb_Phases_SubPhases = "SELECT * FROM tb_phases WHERE (Project_ID = " & myProject_ID & " AND Phase_Parent_ID = " & myPhase_ID & ") ORDER BY Phase_Date_Beginning "
set mySet_tb_Phases_SubPhases = myConnection.execute(mySQL_Select_tb_Phases_SubPhases)

do while not mySet_tb_Phases_SubPhases.eof
	myPhase_ID= mySet_tb_Phases_SubPhases("Phase_ID")
	
	myPhase_Date_Beginning = DateSerial(myYear(mySet_tb_Phases_SubPhases("Phase_Date_Beginning")),myMonth(mySet_tb_Phases_SubPhases("Phase_Date_Beginning")),MyDay(mySet_tb_Phases_SubPhases("Phase_Date_Beginning")))
	myPhase_Date_End   = DateSerial(myYear(mySet_tb_Phases_SubPhases("Phase_Date_End")),myMonth(mySet_tb_Phases_SubPhases("Phase_Date_End")),MyDay(mySet_tb_Phases_SubPhases("Phase_Date_End")))
	
	myPhase_Period_Beginning	= myDay_Width*(myPhase_Date_Beginning - myDate_Beginning_First_Day)*(1+(1/15))
	myPhase_Length	= myDay_Width*(myPhase_Date_End+1 - myPhase_Date_Beginning)*(1+(1/15))
	' Corrected Date 
	
	myPhase_Date_Beginning2 =""
	myPhase_Date_End2  = ""
	If len(mySet_tb_Phases_SubPhases("Phase_Date_Beginning2")) >0 Then
	 myPhase_Date_Beginning2 =DateSerial(myYear(mySet_tb_Phases_SubPhases("Phase_Date_Beginning2")),myMonth(mySet_tb_Phases_SubPhases("Phase_Date_Beginning2")),MyDay(mySet_tb_Phases_SubPhases("Phase_Date_Beginning2")))
	 myPhase_Date_End2   =  DateSerial(myYear(mySet_tb_Phases_SubPhases("Phase_Date_End2")),myMonth(mySet_tb_Phases_SubPhases("Phase_Date_End2")),MyDay(mySet_tb_Phases_SubPhases("Phase_Date_End2")))
	end if
	
	if len(myPhase_Date_Beginning2) > 0 AND len(myPhase_Date_End2) > 0 then
		myPhase_Period_Beginning2	= myDay_Width*(myPhase_Date_Beginning2 - myDate_Beginning_First_Day)*(1+(1/15))
		myPhase_Length2= myDay_Width*(myPhase_Date_End2+1 - myPhase_Date_Beginning2)*(1+(1/15))
	else
		myPhase_Period_Beginning2= 0
		myPhase_Length2	= 0
	end if					
%> <tr> <td height="30" align="left" bgcolor="<%=myBorderColor%>" VALIGN="MIDDLE">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="__Phase_Information.asp?Project_ID=<%=myProject_ID%>&Phase_ID=<%=mySet_tb_Phases_SubPhases("Phase_ID")%>"> <font face="Arial, Helvetica, sans-serif" size="1" color="<%=myBorderTextColor%>"> <%=mySet_tb_Phases_SubPhases("Phase_Name")%></font></a></td><td height="30" align="left"  valign="MIDDLE" colspan="<%=myNumber_of_Months%>" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif" size="-7"> 
<img src="Images/OverApps-lenght_trans.gif" width="<%=myPhase_Period_Beginning%>" height="6"> 
<img src="Images/OverApps-lenght_scheduled.gif" width="<%=myPhase_Length%>" height="6" alt="<%=myPhase_Date_Beginning&" -- "&myPhase_Date_End %>"><br> 
<% if (myPhase_Period_Beginning2>0 OR myPhase_Length2>0) then %><img src="Images/OverApps-lenght_trans.gif" width="<% = myPhase_Period_Beginning2 %>" height="6"> 
<img src="Images/OverApps-lenght_corrected.gif" width="<% = myPhase_Length2 %>" height="6" alt="<% = " " & myPhase_Date_Beginning2 & " -- " & myPhase_Date_End2 %>"><% end if %></font>
</td></tr> <% 
			' NExt Sub Phase
			mySet_tb_Phases_SubPhases.movenext
		loop
	
		'Next Phase
		mySet_tb_Phases.movenext
	loop

 
	mySet_tb_Phases.close
	Set mySet_tb_Phases=Nothing
	myConnection.close
	set myConnection = nothing


%>
</table>


<%
' Legend
%>

<br>
<table border="1" cellpadding="3" cellspacing="0" width="200"> 
<tr>
<td WIDTH="100%" bgcolor="White">
<font color=black face="Arial, Helvetica, sans-serif" size="1">&nbsp;&nbsp;<img src="Images/OverApps-lenght_scheduled.gif" width="80" height="8">&nbsp;&nbsp; 
<%=myMessage_Scheduled%><br>&nbsp;&nbsp;<img src="Images/OverApps-lenght_corrected.gif" width="80" height="8">&nbsp;&nbsp; 
<%=myMessage_Revised%></font><br>
</td>
</tr>
</table>
<br>


<%
' Date Author
%>

<Table border="0" cellpadding="3" cellspacing="0" WIDTH="<%=myApplication_Width%>"> 
<tr>
<td bgcolor="<%=myApplicationColor%>" WIDTH="<%=myApplication_Width%>" align="Center"><p><font face="Arial, Helvetica, sans-serif" size="1" Color="<%=myApplicationTextColor%>"><%=myDate_Display(myProject_Date_Update,2)%> 
-- <%=myProject_Author_Update%> </font> </p>
</td>
</tr>
</table>


<%
' NAVIGATION
%>


<TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" CELLSPACING="0" CELLPADDING="3"> 
<tr>
<td><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="__Projects_List.asp?List=<%=myList%>&Numpage=<%=myNumPage%>&Search=<%=mySearch%>"><%=myMessage_Project%>s</a> </font>
</td>
</tr>
<% if myProject_Member_ID=myUser_ID or myProject_Leader_ID=myUser_ID or myPhase_Member_ID=myUser_ID or myPhase_Leader_ID=myUser_ID or myUser_type_ID=1 then %> 
	<TR>
	<TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> &nbsp;&nbsp;<A HREF="__Phase_Modification.asp?Project_ID=<%=myProject_ID%>"><%=myMessage_Add%>&nbsp;<%=myMessage_Phase%></A></FONT> 
	</TD>
	</TR>
	<% 
end if
%>
</TABLE>

<%
' /NAVIGATION
%>

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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> 
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors
</FONT>
</TD>
</TR>
</TABLE>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright												'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 

</BODY>
</HTML>

<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>