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
<!-- #include file="_INCLUDE/Form_validation.asp" -->
<%
' ------------------------------------------------------------
' Name 		: __identification_Site.asp
' Path    : /
' Description	: Identification procedure to enter a site
' By 	: Pierre Rouarch
' Company : OverApps
' Date : September, 20, 2001
' Versions 1.15.0
'Contributor : Dania Tcherkezoff
'
' Modify by :
' Company :
' Date :
' ------------------------------------------------------------
Dim myPage
myPage="__Identification_Site.asp"

%>

<!-- #include file="_INCLUDE/DB_Environment.asp" -->

<%



Dim myMember_Login, myMember_Password, myError


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myError=0

if Request.form("Validation")=myMessage_Go then


myUser_Login = Request.Form("Member_Login")
myUser_Password = Request.Form("Member_Password")

' Validation du formulaire 
Call myFormSetEntriesInString

' récupération des variables saisies :

myFormCheckEntry null, "Member_Login",true,null,null,0,100
myFormCheckEntry null, "Member_Password",true,null,null,0,100


if not myform_entry_error then

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Site Member 								 	'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Connection
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String
	' Select
	mySQL_Select_tb_Sites_Members = "SELECT * FROM  tb_Sites_Members WHERE Site_ID="&mySite_ID&" AND Member_Login='"&myUser_Login&"' AND Member_Password='"&myUser_Password&"'"

	myConnection.Execute(mySQL_select_tb_Sites_Members)
	set mySet_tb_Sites_Members = myConnection.Execute(mySQL_select_tb_Sites_Members)
	' if it's ok
	if not mySet_tb_Sites_Members.eof then
		myUser_ID = mySet_tb_Sites_Members("Member_ID")
	else
		myError=1
	end if
' Close Recordset and connection
mySet_tb_Sites_Members.close
Set mySet_tb_Sites_Members = Nothing
myConnection.Close
set myConnection = Nothing

End if



if myError=0 and mySite_ID<>0 and myUser_ID<>0 then

'if it s all OK


	Session("Site_ID") = mySite_ID
	session("User_ID") = myUser_ID
	' go in site
	response.redirect("__Home.asp")

end if

End if

%>


<html>

<head>

<title><%=mySite_Name%> - Site Identification</title>
</head>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' BODY TOP
%>

<!-- #include file="_borders/Top.asp" -->

<%
' BODY CENTER
%>

<table WIDTH="<%=myGlobal_Width%>" CELLPADDING="0" CELLSPACING="0" BORDER="0">
<%
' BODY CENTER LEFT
%>

<tr>
<td bgcolor="<%=myBorderColor%>" Width="<%=myLeft_Width%>"> 
<IMG SRC="Images/OverApps-transp.gif" WIDTH="<%=myLeft_Width%>" HEIGHT="1" BORDER="0"></td>

<%
' BODY CENTER APPLICATION
%>


<td With="<%=myApplication_Width%>"> 

<table border="0" WIDTH="100%" CELLPADDING="0" CELLSPACING="1"> 
<form method="POST" 
action="<%=myPage%>"> 

<%
' BODY CENTER APPLICATION TITLE
%>

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><%=myMessage_Identification%></font>
</td>
</tr>
<tr>
<td align="center" colspan="2" bgcolor="<%=myBorderColor%>"><font face="Arial, Helvetica, sans-serif" size="4" Color="<%=myBorderTextColor%>"> 
<%=myMessage_Site%> : <%=mySite_Name%></font>
</td>
</tr> 

<%
' BODY CENTER APPLICATION FIELDS
%>

<tr> <td align="right" bgcolor="<%=myBorderColor%>" width="50%"><font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b>
<%=myMessage_Login%>*</b></font><BR> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><FONT Color="<%=myBorderTextColor%>"><%=myFormGetErrMsg("Member_Login")%></FONT></FONT></B> 
</td><td bgcolor="<%=myBGColor%>" width="50%" ALIGN="CENTER"> <input Type="Text" name="Member_Login" size="14" VALUE="<%=myMember_Login%>"> 
</td></tr> 

<tr> <td align="right" bgcolor="<%=myBorderColor%>" width="50%"><font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Password%>*</b></font><BR> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><FONT Color="<%=myBorderTextColor%>"><%=myFormGetErrMsg("Member_Password")%></FONT></FONT></B> 
</td><td bgcolor="<%=myBGColor%>" width="50%" ALIGN="CENTER"> <input type="Password" name="Member_Password" size="14" VALUE="<%=myMember_Password%>"> 
</td></tr>

<%
' BODY CENTER APPLICATION VALIDATION
%>

<tr> 
<td nowrap bgcolor="<%=myBorderColor%>" width="50%">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" Color="<%=myBorderTextColor%>">&nbsp;* = <%=myMessage_Required%></Font>
</td>
<td nowrap bgcolor="<%=myBGColor%>" width="50%" ALIGN="CENTER"> 
<input type="submit" value="<%=myMessage_Go%>" name="Validation">
</td>
</tr>


<%
' /FORM
%>
</form>


<%
' Separator
%>

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<IMG SRC="Images/OverApps-transp.gif" WIDTH="1" HEIGHT="20">
</td>
</tr> 
<%
' BODY CENTER APPLICATION ERROR LOGIN
%>

<tr ALIGN="CENTER">
<td colspan=2>
<%If myError<>0 then %> 
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><FONT COLOR="#ff0000">
<%=myError_Message_Login_Password%></FONT></FONT></B> <br> 
<% end if %>
</td>
</tr>

<%
' BODY CENTER APPLICATION inscription message
%>



<tr ALIGN="CENTER">
<td colspan=2>
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">&nbsp;</FONT></B><br> 
</td>
</tr>


<tr> 
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"><IMG SRC="Images/OverApps-transp.gif" WIDTH="1" HEIGHT="20">
</td>
</tr> 
</table>
<%
' /BODY CENTER APPLICATION
%>
</td></tr> </table>
<%
' /BODY CENTER
%>
<%
' BODY DOWN
%>
<!-- #include file="_borders/Down.asp" --> 

<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.							'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</body>
</html>

<html><script language="JavaScript"></script></html>