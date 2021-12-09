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

<% 	Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Cache non géré par PWS
'Dim mySQL_Dont_Connect 
mySQL_Dont_Connect = 1
%>
<!-- #include file="_INCLUDE/Global_Parameters.asp" -->


<%
' ------------------------------------------------------------
' Name		: __Administration_SQL.asp
' Path		: /
' Description 	: SQL SERVER Administration Home Page
' By		: Pierre Rouarch, Dania Tcherkezoff	
' Company 	: OverApps
' Date		: December, 11, 2001
' Version   : 1.17.0
'
' Modify by	: 
' Company	:
' Date
' ------------------------------------------------------------

Dim myPage
myPage = "__Administration_Site.asp"

%>


<!-- #include file="_INCLUDE/DB_Environment.asp" -->

<%
If myUser_Type_ID<>1 then
	response.redirect("__Home.asp")
End if
%>



<HTML><HEAD></HEAD><TITLE><%=mySite_Name%> - Site Administration </TITLE>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%> <!-- #include file="_borders/Top.asp" --> <%
' CENTER
%> <TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> <%
' CENTER LEFT
%> <TD WIDTH="<%=myLeft_Width%>"> <!-- #include file="_borders/Left.asp" --> </TD><%
'CENTER APPLICATION
%> <TD WIDTH="<%=myApplication_Width%>" BGCOLOR="<%=myBGCOLOR%>"> 
<%''''''''''''''''''''''''''''''''''START OF SCRIPT''''''''''''''''%>



<table width=100%>
 <tr> 
   <td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"><b><font face="Arial, Helvetica, sans-serif" size="3" color="<%=myApplicationTextColor%>">SQL SERVER</font></b> </td>
  </tr>
</table>	
<br>
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=Mymessage_server%> : "<%= myServer %>"</font><br>
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=mymessage_Database%> : "<%= myDatabase %>"</font><br>
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%= myMessage_Connection%>&nbsp;<%=mymessage_login%> : "<%= myLogin %>"</font><br>
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=mymessage_Password%> : ********</font>
<br>

<%
on error resume next
Dim myconnection_SQL
set myConnection_SQL = Server.CreateObject("ADODB.Connection")
myConnection_SQL.Open myConnection_String_SQL
IF myConnection_SQL.state = 1 Then
%>
<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><b>&nbsp;&nbsp;<%=myMessage_Connection_established%></b><br>
<br>
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><b><%=mymessage_Choose%> : </b><br>
&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question2%>'))document.location='__Administration_Site_SQL3.asp?action=New>';"<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%= myMessage_Create_Table %></font></a><br>
&nbsp;&nbsp;<a href=__Administration_Site.asp><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Leave%></font></a><br>
<%else
 myConnection_SQL.close()
 set myConnection_SQL = NOTHING
 response.redirect "__Administration_Site_SQL.asp?Error=1&server=" & myserver &  "&login=" & mylogin & "&database=" & mydatabase & "&password=" & mypassword
end if
myConnection_SQL.close()
set myConnection_SQL = NOTHING
%>
<br><br><%=myMessage_Warning2%>
<br><br><br><br>

<%''''''''''''''''''''''''''''''''''END OF SCRIPT''''''''''''''''''''''''''''%>
</TD></TR>
<%
'CENTER APPLICATION
%>
</TABLE>
<%
'CENTER
%>
<%
'DOWN
%>
<!-- #include file="_borders/Down.asp" --> 
<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.							                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright			                                	'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</BODY>
</HTML>

<html><script language="JavaScript"></script></html>