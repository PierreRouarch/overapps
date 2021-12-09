<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001  - OverApps - http://www.overapps.com
'
' This program "__Styles_list.asp" is free software; 
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
'	Response.CacheControl = "no-cache" 
' Doesn't Work With PWS ????
mySQL_Dont_Connect = 1
%>

<%
' ------------------------------------------------------------
' Name 			: __DB_Update_Message.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Message during the db importation
' By			: Dania Tcherkezoff
' Company		: OverApps
' Date			: Novermber 21, 2001		
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Styles_List.asp"

Dim myPage_Application
myPage_Application="Styles"
	
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

%>
<html>

<head>
<title><%=mySite_Name%>  Styles - List -</title>
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

<TD WIDTH="<%=myLeft_Width%>"><!-- #include file="_borders/Left.asp" --></td>


<%
' CENTER APPLICATION
%> 

<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 
<table border="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="0" CELLSPACING="0"> 
<tr> 
<td align="center" bgcolor="<%=myApplicationColor%>" > 
<font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><B><%=myDB_Message_Importation%> </b> </font>
</td></tr> </table>
<br>
<br>
<br>
<div align=center><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myBGTextColor%>"><b><%=myDB_Message_Copy%></b></font></div>

<br>
<br>
<br>
<div align=center><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myBGTextColor%>"><%=myDB_Message_Patient%></font></div>
<br><br><br><br><br><br><br><br>
 
<table border="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="0" CELLSPACING="0"> 
<tr> <td align="center" bgcolor="<%=myApplicationColor%>" heigth="100%" WIDTH="<%=myApplication_Width%>"> 
<font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>">&nbsp; </font></td></tr> </table>

</td>
</tr>
</table>
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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001 <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors</FONT>
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