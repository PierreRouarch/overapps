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
%>
<!-- #include file="_INCLUDE/Global_Parameters.asp" -->


<%
' ------------------------------------------------------------
' Name		: __Administration_Site.asp
' Path		: /
' Description 	: Site Administration Home Page
' By		: Pierre Rouarch, Dania Tcherkezoff	
' Company 	: OverApps
' Date		: December, 11, 2001
' Version : 1.15.0
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
%> <TD WIDTH="<%=myApplication_Width%>" BGCOLOR="<%=myBGCOLOR%>"> <table border="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="0" CELLSPACING="0"> 
<tr> <td align="center" bgcolor="<%=myApplicationColor%>" heigth="100%" WIDTH="<%=myApplication_Width%>"> 
<font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><B><%=myMessage_Administration%> 
: <%=mySite_Name%> </b> </font></td></tr> 

<tr BGCOLOR="<%=myBGColor%>" ALIGN="CENTER"> 
<td WIDTH="<%=myApplication_Width%>"><P>&nbsp; </P><P><A HREF="__Administration_Site_General.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR=<%=myBGTextColor%>> 
<%=myMessage_General_Parameters%></font></A> <BR> </P><P><A HREF="__Sites_Members_list.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR=<%=myBGTextColor%>><%=myMessage_Members_Administration%> 
</FONT> </A> </P>
<P><A HREF="__Files_Administration.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR=<%=myBGTextColor%>><%= myFile_Message_Files_Administration %>
</FONT> </A> </P>
<P><A HREF="__Styles_list.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR=<%=myBGTextColor%>><%= myStyles_Administration %>
</FONT> </A> </P>
<P><A HREF="__Administration_Site_SQL.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR=<%=myBGTextColor%>><%=myMessage_SQL_Administration%> </A>
<br><font size=1>
<%If mySQL_Enabled = 1 then%>
<b>(<%=myMessage_sql_mode%>)</b>
<%else%>
<b>(Access Mode)</b>
<%end if%>
</font>
</FONT> </P>
<P><A HREF="__DB_Update_Adv.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR=<%=myBGTextColor%>> Access mode : <%= myDB_Message_Importation %>
</FONT> </A> </P>
<P>&nbsp;</P>





     <P><A HREF="__Administration_Site_About-en.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR=<%=myBGTextColor%>>Copyright
</FONT> </A><BR> </P><P>&nbsp;</P>



</td></tr> 


</table></TD></TR> <%
' / CENTER APPLICATION
%> </TABLE><%
' /CENTER
%> <%
' DOWN
%> <!-- #include file="_borders/Down.asp" --> 
<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.							'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</BODY>
</HTML>

<html></html>