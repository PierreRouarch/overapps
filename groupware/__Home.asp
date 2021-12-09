<%@ LANGUAGE="VBScript" ENABLESESSIONSTATE=TRUE %>

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
'
'-----------------------------------------------------------------------------
%>


<% 	Option Explicit
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" ' do not work with PWS ?????
%>
<!-- #include file="_INCLUDE/Global_Parameters.asp" -->
<!--#include file="_INCLUDE/Environment_tools.asp"-->
<% 
'-----------------------------------------------------------------------------

' Name 		: __home.asp
' Path    : /
' Description 	: Home Page
' By 	: Pierre Rouarch
' Company 	: OverApps
' Date	: January, 03, 2001
' Version : 1.14.2 DEMO VERSION
'
' Modify by :
' Company
' Date
' ------------------------------------------------------------

Dim myPage
myPage = "__home.asp"
%>
<%
 if len(session("user_ID"))=0 then
	mySite_ID=1
	session("Site_ID")=mySite_ID
	myUser_type_ID = 7
	session("User_type_ID")
	response.Redirect("__Identification_site.asp")
end if 
%>

<!-- #include file="_INCLUDE/DB_Environment.asp" -->



<HTML>
<HEAD></HEAD>
<TITLE><%=mySite_Name%> - Home Page</TITLE>


<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>" marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%> <!-- #include file="_borders/Top.asp" --> <%
' CENTER
'%> 

<TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> 

<%
' LEFT
%> 


<TD Width="<%=myLeft_Width%>"><!-- #include file="_borders/Left.asp" --> </TD>
<%
' APPLICATION
%> 

<TD> 



<TABLE BORDER="1"  Width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" CELLPADDING="0" CELLSPACING="0"> 




<%

myNumber_User_Applications=0
Do While myNumber_User_Applications<=myMax_User_Applications

myBox_Title=myUser_Application_Title(myNumber_User_Applications)
Select Case myUser_Application_Name(myNumber_User_Applications)

		Case "Agenda" 
		
%>

		<TR ALIGN="CENTER"> 
		<TD VALIGN="TOP"> 
		<!-- #include file="__Agenda_Box.asp" --> 
		</TD>
		</TR> 

<%
		Case "Projects"		
%>
		<TR ALIGN="CENTER"> 
		<TD VALIGN="TOP"> 
		<!-- #include file="__Projects_Box.asp" --> 
		</TD>
		</TR> 

<%
		Case "Members"		
%>
		<TR ALIGN="CENTER"> 
		<TD VALIGN="TOP"> 
		<!-- #include file="__Sites_Members_Box.asp" --> 
		</TD>
		</TR> 

<%
		Case "News"		
%>
		<TR ALIGN="CENTER"> 
		<TD VALIGN="TOP"> 
		<!-- #include file="__News_Box.asp" --> 
		</TD>
		</TR> 


<%
		Case "Contacts"		
%>
		<TR ALIGN="CENTER"> 
		<TD VALIGN="TOP"> 
		<!-- #include file="__Contacts_Box.asp" --> 
		</TD>
		</TR> 


<%
		Case "Events"		
%>
		<TR ALIGN="CENTER"> 
		<TD VALIGN="TOP"> 
		<!-- #include file="__Events_Box.asp" --> 
		</TD>
		</TR> 


<%
		Case "Webs"		
%>
		<TR ALIGN="CENTER"> 
		<TD VALIGN="TOP"> 
		<!-- #include file="__Webs_Box.asp" --> 
		</TD>
		</TR> 




 <% 

End Select


myNumber_User_Applications=myNumber_User_Applications+1

Loop 
%>


 </TABLE></TD><%
' RIGHT - Not Used
%> </TR> </TABLE><%
' /CENTER
%> <%
' DOWN
%> <!-- #include file="_borders/Down.asp" --> 
<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.							'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001 <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</BODY>
</HTML>

<html><script language="JavaScript"></script></html>