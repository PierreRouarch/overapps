<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001  - OverApps - http://www.overapps.com
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
' 	" Copyright (C) 2001 OverApps & contributors "
' at the bottom of the page with an active link from the name "OverApps" to 
' the Address http://www.overapps.com where the user (netsurfer) could find the 
' appropriate information about this license. 
'-----------------------------------------------------------------------------
%>
<%
' ------------------------------------------------------------
' Name 		: Top.asp
' Path 		: /_borders
' Description 	: Header Site
' By 		: Pierre Rouarch
' Company 	: OverApps
' Date 		: November, 29, 2001
' Version : 1.16.0
'
' Modify by 	:
' Company 	:
' Date 		:
' ------------------------------------------------------------
Dim myTop_Right_Width
myTop_Right_Width=myGlobal_Width-myBanner_Width-myLeft_Width
if myTop_Right_Width<0 then
	myTop_Right_Width=0
end if

Dim myBrowser, myBrowser_Type
Dim Pos

%>
<table width="<%=myGlobal_Width%>" border="0" cellspacing="0" cellpadding="0"> 
<tr>


<%
' LOGO 
%>



<td bgcolor="<%=myBorderColor%>"   VALIGN="Middle" colspan=3 align=left>
<A HREF="http://www.overapps.com"><IMG SRC="Images/OverApps-LogOverappsCom.gif"  BORDER="0"></A>
</td></tr> 

<TR> <td bgcolor="<%=myBorderColor%>" ALIGN="CENTER" width="<%= myLeft_Width %>">&nbsp; 
</TD><TD bgcolor="<%=myBorderColor%>" align=left><A HREF="__Home.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR="<%=myBorderTextColor%>"><B><%=mySite_Name%></B></FONT></a> 
</TD><TD bgcolor="<%=myBorderColor%>" ALIGN="RIGHT" WIDTH="<%=myTop_Right_Width%>"> 
<b><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" Color="<%=myBorderTextColor%>"> 
<%=myMessage_Hello%>&nbsp;<%=myUser_Pseudo%> </font></b> <br> <B><FONT FACE="Arial, Helvetica, sans-serif" Size=1 Color="<%=myBorderTextColor%>"> 
<%
If myDate_Format <> 1 Then
  response.write Day(Now) & "&nbsp;"
end if 
%>
<%
 If Month(Now) = 1 Then response.write   myMessage_January
 If Month(Now) = 2 Then response.write   myMessage_February
 If Month(Now) = 3 Then response.write   myMessage_March
 If Month(Now) = 4 Then response.write   myMessage_April
 If Month(Now) = 5 Then response.write   myMessage_May
 If Month(Now) = 6 Then response.write   myMessage_June
 If Month(Now) = 7 Then response.write   myMessage_July
 If Month(Now) = 8 Then response.write   myMessage_August
 If Month(Now) = 9 Then response.write   myMessage_September
 If Month(Now) = 10 Then response.write   myMessage_October
 If Month(Now) = 11 Then response.write   myMessage_November
 If Month(Now) = 12 Then response.write   myMessage_December
response.write "&nbsp;"
 
If myDate_Format = 1 Then
  response.write Day(Now) & "&nbsp;"
end if 

response.write Year(Now)
%> 

<%'
'=Mid(FormatDateTime(now(),1),instr(FormatDateTime(now(),1)," "))
%>


</FONT></B> 
<% If  myUser_Type_ID=1 then %><br>
<A HREF="__Administration_Site.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" Color="<%=myBorderTextColor%>"> <%=myMessage_Administration%></font>
</A> <%end if %></TD></TR> </TABLE>