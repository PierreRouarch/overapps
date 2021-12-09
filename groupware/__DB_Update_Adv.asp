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
%>

<%
' ------------------------------------------------------------
' Name 			: __DB_Updtad_Adve.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Cuation message before importation  
' By			: Dania TCHERKEZOFF
' Company		: OverApps
' Date			: November 21, 2001		
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__DB_Update.asp"

Dim myPage_Application
myPage_Application="DB Importation"

Dim myFile_System_Object, myFile_Test , myVersion
	
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
if myUser_Type_ID <>  1 then
	Response.redirect("__Quit.asp")
end if


%>
<html>

<head>
<title><%=mySite_Name%></title>
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
<tr> <td align="center" bgcolor="<%=myApplicationColor%>" heigth="100%" WIDTH="<%=myApplication_Width%>"> 
<font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><B><%=myDB_Message_Importation%></b> </font></td></tr> </table>

<br>



<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myDB_Message_Description%></font>
<br>
<br>

<div align=center><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><b><%=myDB_Message_Adv%></b></font></div>

<br>







<br>
<%
'TEST ON DATABASES FILES
myFile_Test = 1
set myFile_System_Object=server.createobject("scripting.FileSystemObject")
myVersion=0

if myFile_System_Object.FileExists(myNew_Database_Path & ".bak")  AND  myFile_System_Object.FileExists(myNew_Database_Path )  Then
  myFile_System_Object.deletefile(myNew_Database_Path)
end if


if myFile_System_Object.FileExists(myNew_Database_Path & ".bak") = True Then
myFile_Test = -1
%> 
<div align=center><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myDB_Message_Done%></font></div><br><br>
<%

end if



if myFile_System_Object.FileExists(myOld_Database_Path_OS & "overapps-software.mdb") Then
 %> 
 <div align=center><a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?done=<%=myFile_Test%>&Version=1';"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.9.5)</b></font></A></div><br>
 <%
 myVersion   = 1 
end if 

if myFile_System_Object.FileExists(myOld_Database_Path_OS & "overapps-software_V1110.mdb") OR myFile_System_Object.FileExists(myOld_Database_Path_OS & "overapps-software_V1111.mdb") Then 
 myVersion   = 2 

 if myFile_System_Object.FileExists(myOld_Database_Path_OS & "overapps-software_V1111.mdb") Then 
  myVersion=3
 end if  
 %> 
 <div align=center><a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?done=<%=myFile_Test%>&Version=<%= myVersion %>';"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.10.x, 1.11.x)</b></font></A></div>
<br>
<%
end if

if myFile_System_Object.FileExists(myOld_Database_Path_GW & "overapps_V113X.mdb") Then
 %> 
 <div align=center><a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?done=<%=myFile_Test%>&Version=4';"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.13.X)</b></font></A></div><br>
 <%
 myVersion   = 4
end if 




if myFile_System_Object.FileExists(myOld_Database_Path_GW & "overapps_V114X.mdb") Then
 %> 
 <div align=center><a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?done=<%=myFile_Test%>&Version=5';"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.14.X)</b></font></A></div><br>
 <%
 myVersion   = 5
end if 

if myFile_System_Object.FileExists(myOld_Database_Path_GW & "overapps_V115X.mdb") Then
 %> 
 <div align=center><a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?done=<%=myFile_Test%>&Version=6';"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.15.X)</b></font></A></div><br>
 <%
 myVersion   = 6
end if 

if myFile_System_Object.FileExists( myOld_Database_Path_GW &  "overapps_V116X.mdb") Then
 %> 
 <div align=center><a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?done=<%=myFile_Test%>&Version=7';"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.16.X)</b></font></A></div><br>
 <%
 myVersion   = 7
end if 

set myFile_System_Object=nothing
%>


<%
If myFile_Test <> 0 Then
%>

<br>

<%
else
%>
<div align=center><font face="Arial, Helvetica, sans-serif" size="" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Cant_Proceed%></b></font></div>


<%
end if
%>

<br><br><br><br>
<table border="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="0" CELLSPACING="0"> 
<tr> <td align="center" bgcolor="<%=myApplicationColor%>"  WIDTH="<%=myApplication_Width%>"> 
&nbsp;</td></tr> </table>
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