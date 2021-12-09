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
' ------------------------------------------------------------
' Name : __Events_Box.asp
' Path : /
' Description : Top 10 Events for home Page
' By : Pierre Rouarch	
' Company : OverApps
' Date : January, 4 2001
' Version : 1.15.0
' Contributor : Dania Tcherkezoff
'
' Modify by :
' Company :
' Date :
' ------------------------------------------------------------
 
' DB Variables
Dim  mySQL_select_tb_events, mySet_tb_events

' events Variables 
Dim myEvent_ID, myEvent_Name, myEvent_Date_Beginning, myEvent_Date_End, myEvent_Presentation

Dim myCalendar_ID

' Only One Events- Calendar in this Version 
myCalendar_ID=1


%>
<table border="0" CELLPADDING="0" CELLSPACING="0" > <TR> <TD><IMG SRC="Images/OverApps-transp.gif" WIDTH="<%=myApplication_Width%>" HEIGHT="1"></td></tr> 
<tr> <td align="center" bgcolor="<%=myApplicationColor%>"><b><font face="Arial, Helvetica, sans-serif" color="<%=myApplicationTextColor%>"><%=myBox_Title%> 
</font></b></td></tr> <tr ALIGN="CENTER"> <td> <%	
' DB connection 
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String


' Top 10 Events
mySQL_Select_tb_Events = "SELECT TOP 10 * FROM tb_Events INNER JOIN tb_Calendars_sites ON tb_Events.Calendar_ID=tb_Calendars_Sites.Calendar_ID WHERE tb_calendars_sites.Site_ID="&mySite_ID


' Coming Events
mySQL_Select_tb_events = mySQL_Select_tb_events &"  AND (tb_Events.Event_Date_Beginning >= '" & myDate_Now &"' OR  tb_Events.Event_Date_End >= '" & myDate_Now &"')  ORDER BY tb_Events.Event_Date_Beginning ASC"





	set mySet_tb_events = 	myConnection.Execute(mySQL_Select_tb_events)
if mySet_tb_events.eof then %> <table><tr ALIGN="CENTER"><td><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1"><div align=center><%=myMessage_No_Event%></div></FONT></td></tr></table><%else %> 
<Table width=<% =myApplication_Width %>> <%	
	do while not mySet_tb_events.eof
	myEvent_ID = mySet_tb_events("Event_ID")
	myEvent_Name = mySet_tb_events("Event_Name")
	myEvent_Date_Beginning = mySet_tb_events("Event_Date_Beginning")
	myEvent_Date_End = mySet_tb_events("Event_Date_End")
	myEvent_Presentation = mySet_tb_events("Event_Presentation")
 %> 
<TR>

          <TD valign="middle" width=120><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" align=left color=<%= myBGTextColor %>><%=myDate_Display(myEvent_Date_Beginning,1)%> 
            <%if myEvent_Date_End>myEvent_Date_Beginning then %>
            <font size="1" face="Courier New, Courier, mono">--&gt;</font> <%=myDate_Display(myEvent_Date_End,1)%> 
            <%end if%>
            </FONT></td>

<td><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1"  color=<%= myBGTextColor %>><%=myEvent_Name%></FONT></td>

          <td valign="middle"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1"  color=<%= myBGTextColor %>><%=myEvent_Presentation%></FONT></td>

</tr> 



<%	mySet_tb_events.movenext
	loop 
 %> </table><% end if ' Eof / not Eof
' Close Recordset and Connection
mySet_tb_events.close
Set mySet_tb_events = Nothing
myConnection.Close
set myConnection = Nothing
%> </td></tr> <tr ALIGN="RIGHT"><td> <A HREF="__Events_List.asp"><FONT SIZE="1" FACE="Arial, Helvetica, sans-serif"><%=myMessage_More%><font size="1" face="Courier New, Courier, mono">--&gt;</font> </FONT> 
</A></td></tr> </table>










<html></html>