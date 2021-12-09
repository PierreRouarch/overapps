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
 	Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Cache non géré par PWS

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
            <td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"><b><font face="Arial, Helvetica, sans-serif" size="3" color="<%=myApplicationTextColor%>"> 
              SQL SERVER</font></b> </td>
          </tr>
</table>	
<form action="__Administration_Site_SQL.asp" method="post">	  
<%

Dim myFSO,myTxtFile


function encrypt(mystring)
 Dim myChar
 Dim mycode(27),myEntry(27)
 dim i , j
 mycode(1) = "b"
 mycode(2) = "c"
 mycode(3) = "d"
 mycode(4) = "e"
 mycode(5) = "g"
 mycode(6) = "h"
 mycode(7) = "i"
 mycode(8) = "j"
 mycode(9) = "k"
 mycode(10) = "u"
 mycode(11) = "a"
 mycode(12) = "v"
 mycode(13) = "w"
 mycode(14) = "x"
 mycode(15) = "y"
 mycode(16) = "z"
 mycode(17) = "p"
 mycode(18) = "q"
 mycode(19) = "t"
 mycode(20) = "r"
 mycode(21) = "s"
 mycode(22) = "o"
 mycode(23) = "n"
 mycode(24) = "m"
 mycode(25) = "l"
 mycode(26) = "f"

 
 myEntry(1) = "a"
 myEntry(2) = "b"
 myEntry(3) = "c"
 myEntry(4) = "d"
 myEntry(5) = "e"
 myEntry(6) = "f"
 myEntry(7) = "g"
 myEntry(8) = "h"
 myEntry(9) = "i"
 myEntry(10) = "j"
 myEntry(11) = "k"
 myEntry(12) = "l"
 myEntry(13) = "m"
 myEntry(14) = "n"
 myEntry(15) = "o"
 myEntry(16) = "p"
 myEntry(17) = "q"
 myEntry(18) = "r"
 myEntry(19) = "s"
 myEntry(20) = "t"
 myEntry(21) = "u"
 myEntry(22) = "v"
 myEntry(23) = "w"
 myEntry(24) = "x"
 myEntry(25) = "y"
 myEntry(26) = "z"
   
 IF len(mystring) > 0 Then
 
 i = 1
 do while i < len(myString) + 1
 
 j = 1

 do while myEntry(j) <> lcase(mid(myString,i,1)) 
  j = j +1
  if j > 26 Then exit do  
 loop 
 
  myChar=  replace( lcase(mid(myString,i,1)) ,myEntry(j),myCode(j))
  
 
 
  encrypt = encrypt & myChar 

  i = i + 1
 loop
 
end if
end function


Dim myError

set myFSO = Server.CreateObject("Scripting.FileSystemObject")

'GET PARAMETERS
IF Request.form("Validation")=myMessage_Go Then
 myServer   = encrypt(Request.Form("Server"))
 myDatabase = encrypt(Request.Form("Database"))
 myLogin    = encrypt(Request.Form("Login"))
 myPassword = encrypt(Request.Form("Password"))
end if

 myError    = Request.QueryString("Error")
 
'IF FILE HAS TO BE  CREATED or CHANGED 
If not myFSO.FileExists(server.MapPath(".") & "\_include\config.asp") OR Request.form("Validation")=myMessage_Go then 
 set  myTxtFile = myFSO.CreateTextFile(server.MapPath(".") & "\_include\config.asp",TRUE,FALSE)
 myTxtFile.WriteLine("<%")
 myTxtFile.WriteLine("Dim myServer, myLogin, myPassword, myDatabase,mySQL_Enabled ")
 myTxtFile.WriteLine("myServer   = """ & myServer & """" )
 myTxtFile.WriteLine("mydatabase = """ & mydatabase & """" )
 myTxtFile.WriteLine("mylogin    = """ & mylogin    & """" )
 myTxtFile.WriteLine("mypassword = """ & mypassword & """" )
 myTxtFile.WriteLine("mySQL_Enabled = 1 " )
 myTxtFile.WriteLine("%" & ">")
 myTxtFile.close()
 response.redirect("__Administration_Site_SQL2.asp")
end if
%>



<%
If len(myError) > 0 Then %>
&nbsp;<font FACE="Arial, Helvetica, sans-serif" SIZE="2" color="<%= myBGTextColor %>"><b><%= mymessage_error %></b>, <%= myMessage_error_parameters %></font>
<%
myServer = Request.QueryString("Server")
myDatabase = Request.QueryString("Database")
myLogin = Request.QueryString("Login")
myPassword = Request.QueryString("Password")
else %>
<Font face="Arial, Helvetica, sans-serif" size="3" color="<%=myBGTextColor%>">&nbsp;&nbsp;<b><%=myMessage_Parameters%> :</b><br><br><%=myMessage_Warning1  %></font> <br><br>
<%
end if
%>

<table border=0 colspan=0 celspan=1>
<tr><td bgcolor="<%= myBorderColor%>" align=right width=200><font FACE="Arial, Helvetica, sans-serif" SIZE="2" color="<%= myBorderTextColor %>"><b><%=mymessage_server%> :</b>&nbsp;</font><td align=left bgcolor="<%=myBGColor%>">&nbsp;<input type=text name=Server value="<%= myServer %>"></tr>
<tr><td bgcolor="<%= myBorderColor%>" align=right><font FACE="Arial, Helvetica, sans-serif" SIZE="2" color="<%= myBorderTextColor %>"><b><%=myMessage_database%> :</b>&nbsp;</b></font><td align=left bgcolor="<%=myBGColor%>">&nbsp;<input type=text name=Database value="<%= myDatabase %>"></tr>
<tr><td bgcolor="<%= myBorderColor%>" align=right><font FACE="Arial, Helvetica, sans-serif" SIZE="2" color="<%= myBorderTextColor %>"><b><%=mymessage_connection%>&nbsp;<%=mymessage_login%> :</b>&nbsp;</b></font><td align=left bgcolor="<%=myBGColor%>">&nbsp;<input type=text name=Login value="<%= myLogin %>"></tr>
<tr><td bgcolor="<%= myBorderColor%>" align=right><font FACE="Arial, Helvetica, sans-serif" SIZE="2" color="<%= myBorderTextColor %>"><b><%=mymessage_password%> :</b>&nbsp;</font><td align=left bgcolor="<%=myBGColor%>">&nbsp;<input type=password name=Password value="<%= myPassword %>"></tr>
<tr><td bgcolor="<%= myBorderColor%>" align=right><input type=submit name=validation value="<%= myMessage_Go %>"><td bgcolor="<%=myBGColor%>">&nbsp;</td></tr>
</table>



<%''''''''''''''''''''''''''''''''''END OF SCRIPT''''''''''''''''''''%>
</TD></TR> <%
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

<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>