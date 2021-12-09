<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   x- Ov er Ap ps -x http://www.overapps.com
'
' This program "__Events_list.asp" is free software; 
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
' Cache non géré par PWS
%>


<%
' ------------------------------------------------------------
' Name 			: __Events_list.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Events List
' by			: Pierre Rouarch
' Company		: OverApps
' Date			: December,10 , 2001
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage
myPage="__Events_List.asp"

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


''''''''''''''''''''''''''''''''''''''''''''''
' VARIABLES
''''''''''''''''''''''''''''''''''''''''''''''
Dim mySearch, myMaxRspByPage, mySortEvent_Date_Beginning, mySortEvent_Name, mySortEvent_Presentation, myOrder, myRs,  myNumPage, myNbrPage, indice,  myRole, myInfo, myModif

Dim  myEvent_Site_ID, myEvent_Member_ID, myEvent_ID, myEvent_Name, myEvent_Presentation, myEvent_Date_Beginning, myEvent_Date_End

Dim myCalendar_ID, myCalendar_Name

Dim i, j

Dim mySQL_Select_tb_Events, mySet_tb_Events, mySQL_Select_tb_Calendars, mySet_tb_Calendars

Dim myList

%>
<html>

<head>
<title><%=mySite_Name%> - Events List</title>
</head>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">

<%
' TOP
%> 

<!-- #include file="_borders/Top.asp" --> 

<%
' CENTER
%> 


<TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> 

<%
' CENTER LEFT
%> 

<TD WIDTH="<%=myLeft_Width%>">
<!-- #include file="_borders/Left.asp" -->
</td>

<%
' CENTER Application
%> 

<%
mySearch=Replace(Request.querystring("search"),"'","''")



if mySearch="" then
	mySearch=Replace(Request.form("search"),"'","''")
end if

myNumPage=Request("Page")
if Len(myNumPage)=0 then
	myNumPage=1
end if
myMaxRspByPage=10

myList=Request("List")


' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


mySortEvent_Date_Beginning = "<a href=""__Events_List.asp?order=Event_Date_Beginning&Page="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Date&"</font></a>"

mySortEvent_Name = "<a href=""__Events_List.asp?order=Event_Name&Page="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Event&"</font></a>"

mySortEvent_Presentation = "<a href=""__Events_List.asp?order=Event_Presentation&Page="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</font></a>"

' Get Sort MEthod
myOrder = Request.QueryString("order")
	Select case myOrder
		case "Event_Name"
			mySortEvent_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Event&"</FONT>"
		case "Event_Presentation"
			mySortEvent_Presentation = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</FONT>"
		case "Event_Date_Beginning"
			mySortEvent_Date_Beginning = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Date&"</FONT>"

		case else
			myOrder="Event_Date_Beginning ASC"
			mySortEvent_Date_Beginning = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Date&"</FONT>"
	End Select

' Read in Database - CALENDAR_ID=1 in this Version


mySQL_Select_tb_Events = "SELECT * FROM tb_Events INNER JOIN tb_Calendars_sites ON tb_Events.Calendar_ID=tb_Calendars_Sites.Calendar_ID WHERE tb_calendars_sites.Site_ID="&mySite_ID



if myList<>"All" then 
	mySQL_Select_tb_events = mySQL_Select_tb_events &"  AND	 (tb_Events.Event_Date_Beginning >= '" & myDate_Now() &"'  OR tb_Events.Event_Date_End >= '" & myDate_Now() & "' ) "
end if 



if mySearch<>"" then
	mySQL_Select_tb_Events=mySQL_Select_tb_Events & " AND (tb_Events.Event_Name LIKE '%"&mySearch&"%' OR tb_Events.Event_Presentation LIKE '%"&mySearch&"%')"
end if
if myOrder <> "Event_Name" Then
 mySQL_Select_tb_Events=mySQL_Select_tb_Events & " ORDER BY " & myOrder &", tb_Events.Event_Name"
else 
 mySQL_Select_tb_Events=mySQL_Select_tb_Events & " ORDER BY " & myOrder 
end if  

'response.write mySQL_Select_tb_events
'response.end

set mySet_tb_Events = myConnection.Execute(mySQL_Select_tb_Events)
%> 



<TD VALIGN="top" ALIGN="left" bgcolor="<%=mybgColor%>" Width="<%=myApplication_Width%>"> 

<%
' APPLICATION TITLE
%>

<TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" BGCOLOR="<%=myApplicationColor%>"> 
<TR>
<TD><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font>
</TD>
</TR> 
</TABLE>

<br> 

<%
' Search Box
%>


<table border="0" Width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0"> 
<tr ALIGN="CENTER">
<td>
<form method="post" action="<%=myPage%>" id=form1 name=form1> 
<input type="text" name="search" size="30" value="<%=mySearch%>"> &nbsp; <INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"></form>
</td>
</tr> 
</table>


<BR> 



<%
i=0
myRs=(myNumPage-1)*myMaxRspByPage
j=0
if not mySet_tb_Events.bof then mySet_tb_Events.MoveFirst
do while not mySet_tb_Events.eof
i=i+1
mySet_tb_Events.movenext
loop
if not mySet_tb_Events.bof then
mySet_tb_Events.MoveFirst
mySet_tb_Events.Move(myRs)
end if

%> 



<table border="0" cellpadding="5" cellspacing="1" Width="<%=myApplication_Width%>"> 
<tr Width="<%=myApplication_Width%>">

<td valign="top" align="left" bgcolor="<%=myBorderColor%>" >
<b><font face="Arial, Helvetica, sans-serif" size="2"><%=mySortEvent_Date_Beginning%> 
</font></b>
</td>

<td valign="top" align="left" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2"><%=mySortEvent_Name%> 
</font></b>
</td>

<td valign="top" align="left" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2"><%=mySortEvent_Presentation%> </font></b>
</td>

<td valign="top" align="left" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_More%> 
</font></b></td></tr> 

<%

IF mySet_tb_Events.eof Then%>

<tr Width="<%=myApplication_Width%>">
	<td  colspan=4 align=center>
	<font face="Arial, Helvetica, sans-serif" size="2"><i><%= myMessage_No_Event %>	</i>
</td></tr> 



<%
end if
	
do while not mySet_tb_Events.eof AND (myMaxRspByPage>j)
	j=j+1
	myEvent_Site_ID = mySet_tb_Events("Site_ID")
	myEvent_Member_ID = mySet_tb_Events("Member_ID")
	myCalendar_ID   = mySet_tb_Events("Calendar_ID")
	myEvent_ID        = mySet_tb_Events("Event_ID")
	myEvent_Name         = mySet_tb_Events("Event_Name")
	myEvent_Presentation  = mySet_tb_Events("Event_Presentation")
	myEvent_Date_Beginning =  myDate_Display(mySet_tb_Events("Event_Date_Beginning"),1)
	myEvent_Date_End =  myDate_Display(mySet_tb_Events("Event_Date_End"),1)


' myInfo Not Used in this version - All the information is on the List Page

	myInfo  = "<a href=""__Event_Information.asp?Event_ID=" & myEvent_ID & """>" & "<img border=""0"" src=""images/overapps-info.gif"" WIDTH=""20"" HEIGHT=""20"" " & " alt="" " & myEvent_Name & """></a>"

	myModif = "<a href=""__Event_Modification.asp?Action=Update&List=" & myList&"&Event_ID=" & myEvent_ID & """>" 	& "<img border=""0"" src=""images/overapps-update.gif"" WIDTH=""20"" HEIGHT=""22"" " & " alt="" " & myEvent_Name & """></a>"

	%> 

	<tr Width="<%=myApplication_Width%>">
	<td align="left">
	<font face="Arial, Helvetica, sans-serif" size="2"><strong><%=myEvent_Date_Beginning%><%if mySet_tb_Events("Event_Date_End")>mySet_tb_Events("Event_Date_Beginning") then %> <font size="2" face="Courier New, Courier, mono">--&gt;</font>  <%=myEvent_Date_End%><%end if%></strong></font>
	</td>
	<td align="left">
	<font face="Arial, Helvetica, sans-serif" size="2"><strong><% = myEvent_Name %></strong></font>
	</td>
	<td align="left" >
	<P align="Justify"><font face="Arial, Helvetica, sans-serif" size="2"><%=myEvent_Presentation%></font></p>
	</td><td align="right" > <%=myModif%> </td></tr> 
	<% 	
	mySet_tb_Events.movenext
loop 
%>

<tr Width="<%=myApplication_Width%>"> <td align="left" valign="middle" colspan="4" bgcolor="<%=myApplicationColor%>"> 
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myApplicationTextColor%>"><b>PAGE(S) 
:&nbsp; <%myNbrPage=int((i+myMaxRspByPage-1)/myMaxrspbyPage)
          indice=1
          do While not indice>myNbrPage 
			if CInt(indice)=CInt(myNumPage) then
          %>[<%=indice%>]&nbsp; <%else%> <a href="__Events_List.asp?Page=<%=indice%>&search=<%=mySearch%>&oder=<%=myOrder%>">[<%=indice%>]</a>&nbsp; 
<%
			end if
			indice=indice+1
          loop
          %>&nbsp;</b></FONT> </td></tr> </table>

<%
' ADMINISTRATION
' EveryBody Can Add And Event in this Version
%> 

<table border="0">
<tr>
<td>
<font face="Arial, Helvetica, sans-serif" size="2"> 
<a href="__Event_Modification.asp?Action=New&Calendar_ID=<%=myCalendar_ID%>"><%=myMessage_Add%>&nbsp;<%=myMessage_Event%></a> 
<%
if myUser_type_ID=1 then 
%> 
	<% 
	if myList<>"All" then 
	%>
	 , <a href="__Events_List.asp?List=All"><font face="Arial, Helvetica, sans-serif" size="2"><%=myMessage_All%>&nbsp;<%=myMessage_Event%>s</font></a> 
	<% 
	Else 
	%>
	 , <a href="__Events_List.asp"><font face="Arial, Helvetica, sans-serif" size="2"> <%=myMessage_Event%>s&nbsp;<%=myMessage_To_Come%></font></a> 
	<%
	end if
	%>
<%
end if
%>
</font>
</td>
</tr>
</table>
</td>
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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> 
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> OverApps</font></A> & contributors
</FONT>
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
<%
	myConnection.Close
	set myConnection = Nothing
%>
