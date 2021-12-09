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

<% 	'Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Cache non géré par PWS
%>
<!-- #include file="_INCLUDE/global_parameters.asp" -->
<%''''''''''''''''''''''''''''''''''START OF SCRIPT''''''''''''''''%>
<%
Dim myConnection_type, myError, myFSO,myTxtFile
 
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

 myConnection_type = Request.QueryString("Connection_Type")
 
 'Access Connection Not Working
 if Request.QueryString("Error") = 1 and myConnection_String = myConnection_String_Access then myError = 1 
 'SQL Server Connection Not Working
 if Request.QueryString("Error") = 1 and myConnection_String = myConnection_String_SQL  then myError = 2
 
 'tb_sites is empty is access base 
  if Request.QueryString("Error") = 2 and myConnection_String = myConnection_String_Access then myError = 3 
 'tb_sites is empty is sql server base
  if Request.QueryString("Error") = 2 and myConnection_String = myConnection_String_SQL  then myError = 4 
 

 

 
 'IF THERE IS A CONNECTION ERROR , ACCES CONECTION IS ACTIVATED
 If myConnection_type = "access" Then
 
  myServer   = encrypt(myServer)
  myDatabase = encrypt(myDatabase)
  myLogin    = encrypt(myLogin)
  myPassword = encrypt(myPassword)
  
  set myFSO = Server.CreateObject("Scripting.FileSystemObject")
  set  myTxtFile = myFSO.CreateTextFile(server.MapPath(".") & "\_include\config.asp",TRUE,FALSE)
  
  myTxtFile.WriteLine("<%")
  myTxtFile.WriteLine("Dim myServer, myLogin, myPassword, myDatabase,mySQL_Enabled ")
  myTxtFile.WriteLine("myServer   = """ & myServer & """" )
  myTxtFile.WriteLine("mydatabase = """ & mydatabase & """" )
  myTxtFile.WriteLine("mylogin    = """ & mylogin    & """" )
  myTxtFile.WriteLine("mypassword = """ & mypassword & """" )
  myTxtFile.WriteLine("mySQL_Enabled = 0 " )
  myTxtFile.WriteLine("%" & ">")
  myTxtFile.close() 
  response.redirect ("__Administration_Site.asp")
 end if
 
%>


<font size=3 face=Arial> <b>Connection ERROR</b></font>
<br><br>
<font face=Arial size=2>
<% If myError = 1 then %>
<b>Can't Connect to your Access Database : </b><br>
The configured path for your database is <i><%= myDatabase_Path %></i><br> You can change this in <i>...\Groupware\_Include\Global_Parameters.asp (line 180)</i>
<%end if%>
<% If myError = 2 then %>
<b>Can't Connect to your SQL Server Database : </b><br>
Your server is maybe down, you can use the back button and retry to connect<br>
If you have to re-configure you SQL Server parameters, you have to turn to access db mode, and go to SQL server Administration, for this <a href=__Connection_Error.asp?connection_type=access>click here</a>
<%end if%>
<% If myError = 3 then %>
<b>Access Database is empty :</b><br>

<%end if%>
<% If myError = 4 then %>
<b>SQL Server Database is empty : </b><br>
Your database seems to be empty, maybe the importation has not been done or has not worked
<br>you have to turn to access db mode, and go to SQL server Administration and select a database to be imported, for this <a href=__Connection_Error.asp?connection_type=access>click here</a>

<%end if%>
</font>
<br><br><br><br><br><br>
<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.							'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0" align=right><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</BODY>
</HTML>

<html><script language="JavaScript"></script></html>
<html></html>
<html><script language="JavaScript"></script></html>