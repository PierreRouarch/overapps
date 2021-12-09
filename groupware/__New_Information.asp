<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   kkk OverApps kkk http://www.overapps.com
'
' This program "__New_Information.asp" is free software; 
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
'Doesn't Work with PWS ????
%>

<%
' ------------------------------------------------------------
' Title 		: __New_Information.asp
' Path    		: /
' Version 		: 1.15.0
' Description 	: Article
' by			: Pierre Rouarch	
' Company		: OverApps
' Date			: December, 10, 2001
' Contributions : Dania Tcherkezoff
'
' Modify by 	:
' Company		:
' Date			:
' ------------------------------------------------------------

Dim myPage
myPage = "__New_Information.asp"
Dim myPage_Application
myPage_Application="News"
	
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim myNew_Site_ID, myNew_Member_ID, myNew_ID, myNew_Title,  myNew_Description_Short, myNew_Description_Long,  myNew_Date, myNew_Date_Update, myNew_Author_Update  

Dim   myNewsWire_ID, myNewsWire_Name

Dim  mySQL_Select_tb_News, mySet_tb_News, mySQL_Select_tb_NewsWires, mySet_tb_NewsWires

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Parameters
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Force to NewsWire 1 in this Version 
myNewsWire_ID=1

' Get Parameters
myNew_ID = Request.QueryString("New_ID")
' If Nothing Go Back
if len(myNew_ID & " ") = 1 or myNew_ID<1 then
		Response.Redirect("__News_List.asp")
end if

%>

<html>
<head>
<title><%=mySite_Name%> - <%=myMessage_Article%> </title>
</head>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">


<%
' TOP
%>

<!-- #include file="_borders/Top.asp" --> 

<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>"  BGColor="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 

<TR VALIGN="TOP">

<%
' CENTER - LEFT
%>

<TD WIDTH="<%=myLeft_Width%>">

<!-- #include file="_borders/Left.asp" --> 

</td>

<%
' CENTER - APPPLICATION
%>


<%
	
' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


' Select Article
mySQL_Select_tb_News = "SELECT *,New_Description_Long,New_Date_Update,New_Author_Update FROM tb_News INNER JOIN tb_NewsWires_Sites ON tb_News.NewsWire_ID=tb_Newswires_Sites.NewsWire_ID WHERE tb_NewsWires_Sites.Site_ID="&mySite_ID&" AND NEW_ID="&myNew_ID 

set mySet_tb_News = myConnection.Execute(mySQL_Select_tb_News)

' If Nothing go back
if mySet_tb_News.eof then
	' Close Recordset
	mySet_tb_News.close
	set mySet_tb_News = nothing
	' Close Connection
	myConnection.close
	set myConnection = nothing
	Response.Redirect("__News_List.asp")
end if

' Get Informations
myNew_Site_ID = mySet_tb_News("Site_ID")
myNew_Member_ID = mySet_tb_News("Member_ID")
myNewsWire_ID = mySet_tb_News("NewsWire_ID")
myNew_ID = mySet_tb_News("New_ID")
myNew_Title  = mySet_tb_News("New_Title")
myNew_Description_Short = mySet_tb_News("New_Description_Short")	
myNew_Description_Long = mySet_tb_News("New_Description_Long")
myNew_Date     = myDate_Display(mySet_tb_News("New_Date"),2)
myNew_Date_Update     = myDate_Display(mySet_tb_News("New_Date_Update"),2)
myNew_Author_Update   = mySet_tb_News("New_Author_Update")


' Close Recordset
mySet_tb_News.close
set mySet_tb_News = nothing		
' Close Connection
myConnection.close
set myConnection = nothing

%> 

<td valign="top" WIDTH="<%=myApplication_Width%>" BGColor="<%=myBGColor%>"> 

<%
' Application Title
%>

<table WIDTH="<%=myApplication_Width%>" border="0" cellpadding="5" cellspacing="1">

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></FONT>
</TD>
</TR> 


<%
' Date
%>

<tr> 
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" 
Color="<%=myBorderTextColor%>"><%=myMessage_Date%></font></b></td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><strong><%=myNew_Date%></strong></font>
</td>
</tr> 

<%
' New 's Title
%>

<tr> 
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Title%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><strong><%=myNew_Title%></strong></font>
</td>
</tr> 

<%
' Presentation 
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Presentation%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myNew_Description_Short%></font>
</td>
</tr>

<%
' Article
%>
 
<tr>
<td valign="top" align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" COLOR="<%=myBorderTextColor%>"><%=myMessage_Article%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myNew_Description_Long%></font>
</td>
</tr>

<%
' Date - author
%>

 
<tr>
<td valign="top" align="right" colspan="2" bgcolor="<%=myApplicationColor%>"> 
<p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="<%=myApplicationTextColor%>"><%=myNew_Date_Update%> -- <%=myNew_Author_Update%></font></p>
</td>
</tr> 

</table>

<%
' ADMINISTRATION - Everybody can add/modify/delete in this version
%>
<table border="0" width="90%" cellpadding="3" cellspacing="0">
 
<tr>
<td>&nbsp;<a href="__New_Modification.asp?action=Update&amp;New_ID=<%=myNew_ID%>"><font face="Arial, Helvetica, sans-serif" size="2"><%=myMessage_Modify%></font></a> 
, <a href="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__New_Modification.asp?action=Delete&amp;New_ID=<%=myNew_ID%>';"><font face="Arial, Helvetica, sans-serif" size="2"><%=myMessage_Delete%></font></a>
</td>
</tr>

 </table>

<%
' /Administration
%>

</td></TR></TABLE>

<%
' /CENTER
%>



<%
' DOWN
%>


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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>" >OverApps</font></A> & contributors</FONT>
</TD>
</TR>
</TABLE>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 

</body>
</html>

<% 

%>

<html><script language="JavaScript"></script></html>