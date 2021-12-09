<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001  - OverApps - http://www.overapps.com
'
' This program "__Contact_Information.asp" is free software; 
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

<% 
	Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1 
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Cache non géré par PWS
mySQL_Dont_Connect = 1
%>

<%
' ------------------------------------------------------------
' Name			   : __DB_Update_Begin.asp
' Path   		   : /
' Vertsion 	   : 1.15.0
' Description  : Tools for updating old DB to version 1.11.0
' By			     : Dania TCHERKEZOFF												
' Company		   : OverApps
' Date			   : Novermber 21, 2001
' ------------------------------------------------------------

Dim myPage
myPage = "__DB_Update_Begin.asp"

	
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
if myUser_Type_ID <> 1 then
	Response.redirect("__Quit.asp")
end if

Dim MyDone, myVersion

myDone = Request.QueryString("done") 
myVersion = Request.QueryString("Version")
%>


<frameset rows=100%,0 border=0>
<frame src="__DB_Update_Message.asp" name=message>
<frame src="__DB_Update.asp?done=<%=myDone%>&version=<%= myVersion %>" name=work>
</frameset>

<html><script language="JavaScript"></script></html>