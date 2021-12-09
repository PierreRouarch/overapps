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
'
'-----------------------------------------------------------------------------
%>
<%
' ----------------------------------------------------------------------------
' Name 		: __Agenda_Box.asp
' Path    	: /
' Description 	: Calendar
' By 		: Pierre Rouarch	
' Company 	: OverApps
' Date		: December,10, 2001
' Versions : 1.15.0
'
' Contributions  : Jean-Luc Lesueur , Christophe Humbert, Dania Tcherkezoff 
'
' Modify by	:
' Company	:
' Date		: 	
' ----------------------------------------------------------------------------

' Current Date of the Agenda 
Dim myStrDate_Agenda     ' Date Agenda in String for request and input in DB
Dim myIntDate_Agenda     ' Date Agenda in Integer for Calculation

' Current Year, Month and Day of the Agenda
Dim myYear_Agenda
Dim myMonth_Agenda
Dim myDay_Agenda

Dim myIntFirst_Day_Of_Month
Dim myNumDay
Dim myDateNumVar
Dim myIntFirst_Day_Of_Week

Dim i, j, myColors, myTDBGColor, myString, myType, b, a


if myPage="__Agenda_Month.asp" then 
	myType=2
else
	myType=0
end if


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



 myIntFirst_Day_Of_Week = DatePart("w",myStrDate_Agenda) - 1 - myAgenda_Week_Start
	 
' First Day of Month
myIntFirst_Day_Of_Month = DateSerial(myYear_Agenda,myMonth_Agenda,1)
' Number of the First Day of Month
myNumDay=DatePart("w",myIntFirst_Day_Of_Month)-1


%>
<table align=left border=0 cellpadding=0 width="<%=myApplication_Width%>" cellspacing="0" HEIGHT="1%"> 
<%
' AGENDA TITLE
%> <tr bgcolor="<%=myApplicationColor%>"> <td ALIGN="Center" Colspan=2><b><font face="Arial, Helvetica, sans-serif"  color="<%=myApplicationTextColor%>"><%=myBox_Title%> 
</font> </b></td></tr> 
<%
' CALENDAR
%> <TR><TD vAlign=top width="1%" colspan=2> 

<TABLE border=1 cellPadding=0 cellSpacing=0 width="100%" > 

<%
' DAY, WEEK, MONTH, TODAY
%>
 
<TR> 

<%
' DAY
%>
<TD align=middle bgColor="<%=myBGColor%>"> 
<A HREF="__Agenda_Day.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda,myDay_Agenda-1)%>"> 
<IMG BORDER=0 HEIGHT=11 SRC="Images/OverApps-left_small.gif" WIDTH=11></A> <FONT color="<%= myBGTextColor %>" face=Arial,Helvetica size=1><a href="__Agenda_Day.asp?Date_Agenda=<%=myIntDate_Agenda%>">
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
	
 response.write Day(myStrDate_Agenda)& ", "
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


</A></FONT> 
<A HREF="__Agenda_Day.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda,myDay_Agenda+1)%>"> 
<IMG BORDER=0 HEIGHT=11 SRC="Images/OverApps-right_small.gif" WIDTH=11></A> </TD><TD align=middle bgColor=<%if mytype=1 then%>f2bfbf<%else%><%=myBGColor%><%end if%>><a href="__Agenda_Week.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda,myDay_Agenda-7)%>"> 
<img border=0 height=11 src="Images/OverApps-left_small.gif" width=11></a> <FONT color=<%= myBGTextColor %>
            face=Arial,Helvetica size=1> <a href="__Agenda_Week.asp?Date_Agenda=<%=myIntDate_Agenda%>"> 
&nbsp;<%
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
	
 response.write Day(myIntDate_Agenda-myIntFirst_Day_Of_Week)& ", "
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
<font size="1" face="Courier New, Courier, mono">--> </font>
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

</A></FONT> <a href="__Agenda_Week.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda,myDay_Agenda+7)%>"> 
<img border=0 height=11 src="Images/OverApps-right_small.gif" width=11></a> </td><TD align=middle bgColor=<%=myBGColor%>> 
<a href="__Agenda_Month.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda-1,1)%>"> 
<IMG border=0 height=11 src="Images/OverApps-left_small.gif" width=11></A> <FONT color=blue 
            face=Arial,Helvetica size=1><a href="__Agenda_Month.asp?Date_Agenda=<%=myIntDate_Agenda%>"><% Response.Write " "&myMonth_Agenda&" , "&myYear_Agenda&" " %></A></FONT> 
<a href="__Agenda_Month.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,myMonth_Agenda+1,1)%>"> 
<IMG border=0 height=11 src="Images/OverApps-right_small.gif" width=11></A> </TD><TD align=middle  bgcolor=#ffffff> 
<A href="__Agenda_Day.asp"><FONT size="1" face="Arial,Helvetica" Color="Blue"><B><%=myMessage_Today%></b></font></A> 
</TD>
</TR> 
</table>
</td>
</TR> 
<TR bgColor="<%=myBGColor%>"> 
<TD align=middle colspan=2>
 
 <% if mytype=0 or mytype=1 then %> 
 <table border=1 width=100% bgColor="#FFFFFF" cellPadding=0 cellSpacing=0><tr><td>
 <TABLE bgColor="#FFFFFF" border=0 cellPadding=1 cellSpacing=3 width=100%>
<%
' CALENDAR BY DAYS
%> <%
' Days Titles
%> <TR> <%For i=0 to 41 %> <TD align=center vAlign=top> <FONT face="Arial,Helvetica" size="1" Color="Black"> 

<%
If DatePart("w",myIntFirst_Day_Of_Month-myNumDay+i) = 1 Then 
 if myAgenda_Week_Start = 0 Then  
   response.write "<font color=red>" & Ucase(left(myMessage_Sunday,1)) & "</font>"
 else 
    response.write Ucase(left(myMessage_Sunday,1)) 
 end if
end if

If DatePart("w",myIntFirst_Day_Of_Month-myNumDay+i) = 2 Then 
 if myAgenda_Week_Start = 1 Then  
   response.write "<font color=red>" & Ucase(left(myMessage_Monday,1)) & "</font>"
 else 
    response.write Ucase(left(myMessage_Monday,1)) 
 end if
end if


If DatePart("w",myIntFirst_Day_Of_Month-myNumDay+i) = 3 Then response.write Ucase(left(myMessage_Tuesday,1))
If DatePart("w",myIntFirst_Day_Of_Month-myNumDay+i) = 4 Then response.write Ucase(left(myMessage_Wednesday,1))
If DatePart("w",myIntFirst_Day_Of_Month-myNumDay+i) = 5 Then response.write Ucase(left(myMessage_Thursday,1))
If DatePart("w",myIntFirst_Day_Of_Month-myNumDay+i) = 6 Then response.write Ucase(left(myMessage_Friday,1))
If DatePart("w",myIntFirst_Day_Of_Month-myNumDay+i) = 7 Then response.write Ucase(left(myMessage_Sunday,1))


%>
</FONT> </td>
<%Next%> 
</tr> 

<tr> 

<%
' DAYS
%> 


<% For j=0 to 41 %>  
<% 
myDateNumVar=myIntFirst_Day_Of_Month-myNumDay+j
mycolors="Blue"
myTDBGColor="White"
if Month(myDateNumVar)<>myMonth_Agenda then
	myTDBGColor="White"
	myColors="Black"
end if
If myDateNumVar=Cdate(myStrDate_Agenda) then
	myTDBGColor="White"
	mycolors="#ff0000"
end if
%> 
<td BGColor=<%=myTDBGColor%>>
<a href="__Agenda_Day.asp?Date_Agenda=<%=DateSerial(Year(myDateNumVar),Month(myDateNumVar),Day(myDateNumVar))%>"> 
<FONT color=<%=mycolors%> size="1" face="Arial,Helvetica">
<%=Day(myDateNumVar)%></font></a> 
</td><%Next%> </TR> </TABLE></td></tr><%end If%> 


 <% if mytype=2 then %> 
<TABLE  border=0 cellPadding=1 cellSpacing=3 width=100%>
<tr> 
<% For b=1 to 12 %> 
<td align=center
<%if myMonth_Agenda=b then%>
 bgcolor="White"
<%else%>
 bgcolor="<%= myBGColor %>" 
<%end if%>> 

<%if myMonth_Agenda=b then%>
 <table border=0 cellPadding=0 cellSpacing=0 width=100% bgColor="<%= myBorderColor %>"><tr><td align=center BGcolor="<%= myBorderColor %>"> <font Face="Arial" color="<%= myBorderTextColor %>"> 
<%else%> 
  <a href="__Agenda_Month.asp?Date_Agenda=<%=DateSerial(myYear_Agenda,b,1)%>"><font color="<%= myBGTextColor %>" Face="Arial">
<%end if%>
<%
 If b = 1 Then response.write   left(myMessage_January,20)
 If b = 2 Then response.write   left(myMessage_February,20)
 If b = 3 Then response.write   left(myMessage_March,20)
 If b = 4 Then response.write   left(myMessage_April,20)
 If b = 5 Then response.write   left(myMessage_May,20)
 If b = 6 Then response.write   left(myMessage_June,20)
 If b = 7 Then response.write   left(myMessage_July,20)
 If b = 8 Then response.write   left(myMessage_August,20)
 If b = 9 Then response.write   left(myMessage_September,20)
 If b = 10 Then response.write  left(myMessage_October,20)
 If b = 11 Then response.write  left(myMessage_November,20)
 If b = 12 Then response.write  left(myMessage_December,20)


%>
<%if myMonth_Agenda=b then%>
</font></font</a></td></tr></table>
<%end if%>
</font></font></td>
<%Next%> 

</tr> <%end if%> </table></TD></TR></TABLE>






