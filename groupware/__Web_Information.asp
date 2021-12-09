<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - O.verApps - http://www.overapps.com
'
' This program "__Web_Information.asp" is free software; 
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
' Does n't Work with PWS ?
%>

<%
' ------------------------------------------------------------
' Name 			: __Web_Information.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Information about a Web Site
' By			: Pierre Rouarch
' Company		: OverApps
' Date			: February, 2, 2001
' 
' Modify by		: 
' Company		:
' Date			:
' ------------------------------------------------------------



Dim myPage
myPage = "__Web_Information.asp"


Dim myPage_Application
myPage_Application="Webs"
	
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

myApplication_Public_Type_ID = Get_Application_Public_Type_ID(myPage_Application)
if myApplication_Public_type_ID<myUser_type_ID then
	Response.redirect("__Quit.asp")
end if
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET APPLICATION TITLE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myApplication_Title = Get_Application_Title(myPage_Application)


Dim myWebDirectory_ID, myCategory_ID, myWeb_ID, myWeb_Name, myWeb_URL, myWeb_Description_Short, myWeb_Description_Long,  myWeb_Public, myWeb_Top, myWeb_Author_Update, myWeb_Date_Update

Dim  mySQL_Select_tb_Webs, mySet_tb_Webs

' Get Web_ID
myWeb_ID = Request.QueryString("Web_ID")
if len(myWeb_ID & " ") = 1 then
	Response.Redirect("__Webs_List.asp")
end if


%>
<html>

<head>
<title><%=mySite_Name%> - Web Information</title>
</head>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%>

<!-- #include file="_borders/top.asp" --> 
<%
' CENTER
%>


<TABLE WIDTH="<%=myGlobal_Width%>" BGColor="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP">
<%
' CENTER LEFT
%>

 <TD WIDTH="<%=myLeft_Width%>"> <!-- #include file="_borders/left.asp" --> 
</td>

<%
' CENTER APPLICATION
%> 



<%
' DB Connection 
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

	
mySQL_Select_tb_Webs = "SELECT *,Web_Author_Update FROM tb_Webs WHERE Site_ID = "&session("Site_ID")&" AND Web_ID = " & myWeb_ID
set mySet_tb_Webs = myConnection.Execute(mySQL_Select_tb_Webs)

' if nothing go Back
if mySet_tb_Webs.eof then
	' Close Recordset
	mySet_tb_Webs.close
	Set mySet_tb_Webs=Nothing
	' Close Connection
	myConnection.close
	set myConnection = nothing
	Response.Redirect("__Webs_List.asp")
end if


' Not Used in this Version 
myWebDirectory_ID = mySet_tb_Webs("WebDirectory_ID")
myCategory_ID = mySet_tb_Webs("Category_ID")

myWeb_ID = mySet_tb_Webs("Web_ID")
myWeb_Name  = mySet_tb_Webs("Web_Name")
myWeb_URL   = mySet_tb_Webs("Web_URL")	
myWeb_Description_Short = mySet_tb_Webs("Web_Description_Short")	
myWeb_Description_Long = mySet_tb_Webs("Web_Description_Long")
myWeb_Date_Update     = mySet_tb_Webs("Web_Date_Update")
myWeb_Author_Update   = mySet_tb_Webs("Web_Author_Update")

' Not Used in this version
myWeb_Top = mySet_tb_Webs("Web_Top")
myWeb_Public = mySet_tb_Webs("Web_Public")
 
myWeb_Date_Update     = mySet_tb_Webs("Web_Date_Update")
myWeb_Author_Update   = mySet_tb_Webs("Web_Author_Update")

' Close Recordset
mySet_tb_Webs.close
Set mySet_tb_Webs=Nothing
' Close Connection
myConnection.close
set myConnection = nothing

%> 


<td valign="top" Width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" > 

<table border="0" cellpadding="5" cellspacing="1" width="<%=myApplication_Width%>"> 

<%
' Application Tuitle
%>

<tr> 
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" Color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></td>
</tr> 


<%
' Web Address (URL)
%>


<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_Address%>&nbsp;(URL)</font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><a href="<%=myWeb_URL%>"><%=myWeb_URL%></a></font>
</td>
</tr> 

<%
' Web Name
%>

<tr> 
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_Name%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><strong><%=myWeb_Name%></strong></font>
</td>
</tr> 


<%
' Web Presentation
%>


<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_Presentation%></font></b>
</td>
<td align="left" >
<font face="Arial, Helvetica, sans-serif" size="2"><%=myWeb_Description_Short%></font>
</td>
</tr> 

<%
' Details
%>

<tr>
<td valign="top" align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_More%>
</font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myWeb_Description_Long%></font>
</td>
</tr>

<%
' Date and Author
%>


<tr>
<td valign="top" align="right" colspan="2" bgcolor="<%=myApplicationColor%>"> 
<p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="<%=myApplicationTextColor%>"><% if len(myWeb_Date_Update) > 0 then %> 
<% = myDate_Display(myWeb_Date_Update,2) %> -- <% = myWeb_Author_Update %><% end if %></font></p>
</td>
</tr>

</table>

<%
' ADMINISTRATION - EveryBody Can Add, Delete or Modify in this Version
%> 
<TABLE BORDER="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="3" CELLSPACING="0"> 
<TR>
<TD WIDTH="1%">&nbsp;

</TD>
<TD>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> 
<A HREF="__Web_Modification.asp?WebDirectory_ID=<%=myWebDirectory_ID%>&Action=Update&Web_ID=<%=myWeb_ID%>"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Modify%></font></A>
,
<A HREF="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Web_modification.asp?Action=Delete&amp;Web_ID=<%=myWeb_ID%>';"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Delete%></font></A>  
</FONT> 
 </TD>
</TR>
<TR>
<TD WIDTH="1%">&nbsp;

</TD>
<TD>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> 
<A HREF="__Web_Modification.asp?WebDirectory_ID=<%=myWebDirectory_ID%>&Action=New"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Add%>&nbsp;<%=myMessage_Web%></font></A>
</FONT>
</TD>
</TR>

</TABLE>

<%
' /ADMINISTRATION
%>


</td>
</TR> 
</TABLE>

<%
' /Center
%>


<!-- #include file="_borders/down.asp" --> 


<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.														'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0">
<TR ALIGN="RIGHT">
<TD>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> 
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</FONT>
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


<html></html>
<html><script language="JavaScript"></script></html>