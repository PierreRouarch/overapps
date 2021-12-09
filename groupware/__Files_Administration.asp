<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Styles_Modification.asp" is free software; 
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
' at the bottom of the page with an active link from the name "OverApps"  
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
' Name 		       	: __Files_Administration.asp
' Path   	      	: /
' Version 	    	: 1.15.0
' Description   	: Choose File Extensions to be authorised
' By		        	: Dania TCHERKEZOFF
' Company	      	: OverApps
' Date			      : December 10, 2001
'
' Modify by		    :
' Company		      :
' Date			      :
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Files_Administration.asp"

Dim myPage_Application
myPage_Application="Files"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' INCLUDES 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<!-- #include file="_INCLUDE/Global_Parameters.asp" -->

<!-- #include file="_INCLUDE/Form_validation.asp" -->

<!-- #include file="_INCLUDE/DB_Environment.asp" -->

<!-- #include file="_INCLUDE/Environment_Tools.asp" -->

<!-- #include file="_INCLUDE/Files_Upload_Class.asp" -->

<%

Dim mySet_tb_Files_Extensions, mySQL_Select_tb_Files_Extensions, myCounter, myError, myAction, myFile_Extension_ID


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET APPLICATION TITLE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myApplication_Title = Get_Application_Title(myPage_Application)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CHECK IF THE USER CAN ENTER IN THIS APPLICATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myUser_Type_ID > 1 Then
  Response.Redirect "__Identification_Site.asp"
end if



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET PARAMETERS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

myAction = Request.QueryString("Action")
myFile_Extension_ID = Request.QueryString("File_Extension_ID")



'Open Connection, will serve in any case
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Delete a file extension
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myAction="Delete" Then

myConnection.Execute("Delete From tb_Files_Extensions where File_Extension_Id="& myFile_Extension_ID)

end if


'''''''''''''''''''''''''''''''''''''''''''''''''''
'UPDATING FILES PARAMETERS
'''''''''''''''''''''''''''''''''''''''''''''''''''

if Request.Form("submit") = myMessage_Go Then
Dim mySet_tb_Site_Maximum_Files_Size, mySQL_Select_tb_Site_Maximum_Files_Size




'Updating maximum file size
if myValueIsGood (myerr_numerical, Request.Form("Maximum_File_Size")) Then
mySQL_Select_tb_Site_Maximum_Files_Size = "Select * from tb_Sites"
set mySet_tb_Site_Maximum_Files_Size = Server.CreateObject("ADODB.RecordSet")
mySet_tb_Site_Maximum_Files_Size.open mySQL_Select_tb_Site_Maximum_Files_Size, myConnection,3,3
mySet_tb_Site_Maximum_Files_Size.fields("Site_Maximum_Files_Size") = Request.Form("Maximum_File_Size")
myMaximum_File_Size = Request.Form("Maximum_File_Size")
mySet_tb_Site_Maximum_Files_Size.update
mySet_tb_Site_Maximum_Files_Size.close
set mySet_tb_Site_Maximum_Files_Size = Nothing
else
myError = "A"
end if

'Updating File Extension autorisation

set mySet_tb_Files_Extensions = Server.CreateObject("ADODB.RecordSet")
mySQL_Select_tb_Files_Extensions = "Select * from tb_Files_Extensions"
mySet_tb_Files_Extensions.open mySQL_Select_tb_Files_Extensions, myConnection,3,3

do while not mySEt_tb_Files_Extensions.eof
 if Request.Form(mySet_tb_Files_Extensions.fields("File_Extension")) = "on" Then
 mySet_tb_Files_Extensions.fields("File_Extension_Autorised") = 1
else  
 mySet_tb_Files_Extensions.fields("File_Extension_Autorised") = 0
end if 
 mySet_tb_Files_Extensions.Update
 mySet_tb_Files_Extensions.MoveNext
loop

mySet_tb_Files_Extensions.close
set mySet_tb_Files_Extensions = Nothing


'Add a new Extension
if len(Request.Form("New_File_Extension") ) > 0 Then 
 set mySet_tb_Files_Extensions = Server.CreateObject("ADODB.RecordSet")
 mySQL_Select_tb_Files_Extensions = "Select * from tb_Files_Extensions"
 mySet_tb_Files_Extensions.open mySQL_Select_tb_Files_Extensions, myConnection,3,3
  mySet_tb_Files_Extensions.AddNew
 mySet_tb_Files_Extensions.fields("File_Extension_Autorised") = 1
 mySet_tb_Files_Extensions.fields("File_Extension") =Request.Form("New_File_Extension")
 mySet_tb_Files_Extensions.fields("Site_ID") = mySite_ID
 mySet_tb_Files_Extensions.update
 mySet_tb_Files_Extensions.close
 set  mySet_tb_Files_Extensions = Nothing

end if

'Close Conection
myConnection.close
set myConnection = Nothing 


end if
%>
<html>

<head>
<title><%=mySite_Name%>  </title>
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


<%

%> 



<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 

<%
' APPLICATION TITLE
%>

<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></TD></TR> 
</table>


<%
' FORM BOXS
%>

<form action="__Files_Administration.asp" method=post >
<table  cellpadding="3" cellspacing="1">
<%
'MAXIMUM FILE SIZE
%>
<tr >
<td  bgcolor="<%= myBorderColor %>" align=right width=50%><font face="Arial, Helvetica, sans-serif" size=2 color="<% =myBorderTextColor %>"><b> <%= myFile_Message_Maximum_Size %></b></font></td>
<td ><input type=text name="Maximum_File_Size" value="<%= myMaximum_File_Size %>">
<font face="Arial, Helvetica, sans-serif" size=2 color="<% =myBGTextColor%>"><b>(
<%'
'DISPLAY MAXIMUM SIZE IN Ko OR Mo
 If myMaximum_File_Size <  1048576 Then %>
 
  <%= (int(( myMaximum_File_Size / 1024)*100)) / 100 %> Ko )
   
<%
else
%>

 <%= (int(( myMaximum_File_Size / (1024*1024))*100)) / 100 %> Mo )
 
<%
end if
%>
</b></font>
</td>
</tr>
<%
'FILE EXTENSION
%>

<tr><td bgcolor="<%= myBorderColor %>" align=right>
&nbsp;&nbsp;&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size=2 color="<% =myBorderTextColor %>"><b><%=myFile_Message_Extension%>&nbsp;</b></font></td>
<td>
<%


set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

set mySet_tb_Files_Extensions = Server.CreateObject("ADODB.RecordSet")
mySQL_Select_tb_Files_Extensions = "Select * from tb_Files_Extensions where Site_ID = "& mySite_ID&" order by File_Extension"
set mySet_tb_Files_Extensions = myConnection.Execute(mySQL_Select_tb_Files_Extensions)
myCounter = 0
%>
<table align=center>
<%
Do while not mySet_tb_Files_Extensions.eof
if myCounter = 0 Then response.write "<tr valign=middle>"
%> 
 
<td  ><font face="Arial, Helvetica, sans-serif" size=2><b><%= mySet_tb_Files_Extensions("File_Extension")%></b></font>
</td><td align=left  >
<% If  mySet_tb_Files_Extensions("File_Extension_Autorised") = 1 Then %>
<input type=checkbox value=on checked name="<%= mySet_tb_Files_Extensions("File_Extension")%>">
<%else%>
<input type=checkbox value=on name="<%= mySet_tb_Files_Extensions("File_Extension")%>">
<%end if%>
<a href="__Files_Administration.asp?Action=Delete&File_Extension_ID=<%= mySet_tb_Files_Extensions("File_Extension_ID") %>"><font color="<%= myBGTextColor %>" face="Arial" size=2>X</font></a>&nbsp;&nbsp;</td>        
<%
if myCounter = 2 Then 
 Response.Write "</tr>"
 myCounter = -1
end if
 
myCounter = myCounter + 1
mySet_tb_Files_Extensions.MoveNext
loop		
%>		  

</table>
</td>
</tr>
<tr>
<td bgcolor="<%= myBorderColor %>" align=right width=50%><font face="Arial, Helvetica, sans-serif" size=2 color="<% =myBorderTextColor %>"><b><%= myFile_Message_Add_Extension %></b></font></td>
<td>&nbsp;<input type=text name=New_File_Extension></td>
</tr>



<tr>
<td bgcolor="<%= myBorderColor %>" align=right width=50%><font face="Arial, Helvetica, sans-serif" size=2 color="<% =myBorderTextColor %>"><input type=submit name=submit value="<%= myMessage_Go %>">&nbsp;</font></td>
<td>&nbsp;</td>
</tr>


</table>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>">&nbsp;</font></TD></TR> 
<TR><TD bgcolor="<%=myBGColor%>">&nbsp;
</td></tr>
</table>
</form>
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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors</FONT>
</TD>
</TR>
</TABLE>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' End Copyright									'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 

</body>
</html>

<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>