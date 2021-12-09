<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Webs_list.asp" is free software; 
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
' Doesn't Work With PWS ????
%>

<%
' ------------------------------------------------------------
' Name 			: __Webs_list.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Webs List
' By			: Pierre Rouarch
' Company		: OverApps
' Date			: February 1, 2001 		
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Webs_List.asp"

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

Dim mySearch, myMaxRspByPage, mySortWebDirectory_ID, myWebDirectory_ID, myWebDirectory_Name, mySortWeb_Name, mySortWeb_Description_Short, myOrder, myRs,  myNumPage, myNbrPage, indice, myWeb_URL, myWeb_Name, myWeb_ID, myWeb_Description_Short,  myInfo, myModif

Dim mySQL_Select_tb_Webs, mySet_tb_Webs, mySQL_Select_tb_WebDirectories, mySet_tb_WebDirectories 

Dim i, j


%>
<html>

<head>
<title><%=mySite_Name%>  Webs - List -</title>
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
' Get Parameters

mySearch=Replace(Request.querystring("Search"),"'","''")
if mySearch="" then
		mySearch=Replace(Request.form("Search"),"'","''")
end if

myNumPage=Request("Page")
if Len(myNumPage)=0 then 
	myNumPage=1
end if

myOrder = Request.QueryString("order")
		
myMaxRspByPage=10



' Prepare Sort 

mySortWeb_Name = "<a href=""__Webs_List.asp?order=Web_Name&Page="&myNumPage&"&search="&mySearch&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Web&"</font></a>"

mySortWeb_Description_Short = "<a href=""__Webs_List.asp?order=Web_Description_Short&Page="&myNumPage&"&search="&mySearch&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</font></a>"

Select case myOrder
	case "Web_Name"
		mySortWeb_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Web&"</FONT>"
	case "Web_Description_Short"
		mySortWeb_Description_Short = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</FONT>"
	case else
		myOrder="Web_Name"
		mySortWeb_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Web&"</FONT>"
End Select

' dbConnection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

		
' Get Webs Informations
mySQL_Select_tb_Webs = "SELECT * FROM tb_webs INNER JOIN tb_WebDirectories_sites ON tb_Webs.WebDirectory_ID = tb_WebDirectories_sites.WebDirectory_Id WHERE tb_WebDirectories_sites.Site_ID="&mySite_ID

' Search Purposes
if mySearch<>"" then
	mySQL_Select_tb_Webs=mySQL_Select_tb_Webs & " AND (Web_URL LIKE '%"&mySearch&"%' OR Web_Name LIKE '%"&mySearch&"%' OR Web_Description_Short LIKE '%"&mySearch&"%' OR Web_Description_Long LIKE '%"&mySearch&"%')"
end if


' ORDER
if myOrder <> "Web_Name" Then 
 mySQL_Select_tb_Webs=mySQL_Select_tb_Webs & " ORDER BY " & myOrder &", Web_Name"
else 
 mySQL_Select_tb_Webs=mySQL_Select_tb_Webs & " ORDER BY Web_Name"
end if	 


' Execute
set mySet_tb_Webs = myConnection.Execute(mySQL_Select_tb_Webs)



%> 



<%
' SEARCH BOX
%>


<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 

<%
' APPLICATION TITLE
%>

<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></TD></TR> 
</table>

<%
' SEARCH BOX
%>


<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0"> 
<tr ALIGN="CENTER">
<td> 
<form method="post" action="__Webs_List.asp" id=form1 name=form1> 
<br> &nbsp; <input type="text" name="search" size="30" VALUE="<%=mySearch%>"> &nbsp; <INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"> 
</form>
</td>
</tr>
</table>


<BR> 

<%
' LIST
%>



<%
' Go to the current record
i=0
myRs=(myNumPage-1)*myMaxRspByPage
j=0
if not mySet_tb_Webs.bof then mySet_tb_Webs.MoveFirst
do while not mySet_tb_Webs.eof 
i=i+1
mySet_tb_Webs.movenext
loop 
if not mySet_tb_Webs.bof then 
mySet_tb_Webs.MoveFirst
mySet_tb_Webs.Move(myRs) 
end if
%> 

<%
' ROW TITLES
%>


<table border="0" Width="<%=myApplication_Width%>" cellpadding="5" cellspacing="1">
<tr> 
<td valign="top" align="left" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2"><%=mySortWeb_Name%></font></b>
</td>

<td align="left" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2"><%=mySortWeb_Description_Short%>
</font></b>
</td>

<td valign="top" align="left" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_More%></font></b>
</td>

</tr> 

<%
' LISTING
%>

<%	
do while not mySet_tb_Webs.eof AND (myMaxRspByPage>j)

	j=j+1
	myWebDirectory_ID	= mySet_tb_Webs("WebDirectory_ID")
	myWeb_URL	= mySet_tb_Webs("Web_URL")
	myWeb_Name         = mySet_tb_Webs("Web_Name")
	myWeb_ID        = mySet_tb_Webs("Web_ID")
	myWeb_Description_Short  = mySet_tb_Webs("Web_Description_Short")
	
	' For Multi Web Directories Purpose
	'mySQL_Select_tb_WebDirectories = "SELECT * FROM tb_WebDirectories WHERE	 WebDirectory_ID="&myWebDirectory_ID
	'set mySet_tb_WebDirectories = myConnection.Execute(mySQL_Select_tb_WebDirectories)
	'myWebDirectory_Name = mySet_tb_WebDirectories("WebDirectory_Name")


	myInfo  = "<a href=""__Web_Information.asp?Web_ID=" & myWeb_ID & """>" & "<img border=""0"" src=""images/overapps-info.gif"" WIDTH=""20"" HEIGHT=""20"" " & " alt="" " & myWeb_Name & """></a>"
	myModif = "<a href=""__Web_Modification.asp?Web_ID=" & myWeb_ID & """>" 	& "<img border=""0"" src=""images/overapps-update.gif"" WIDTH=""20"" HEIGHT=""22"" " & " alt="" " & myWeb_Name & """></a>"
	%> 
	<tr> 
	<td align="left" valign="middle">
	<font face="Arial, Helvetica, sans-serif" size="2"><a href="<%=myWeb_URL%>"><strong><%=myWeb_Name%></strong></a></font>
	</td>
	<td valign="middle" align="left">
	<font face="Arial, Helvetica, sans-serif" size="2"><%=myWeb_Description_Short%></font>
	</td>
	<td align="right" valign="middle">
	<% = myInfo %> &nbsp;&nbsp;<% = myModif %>
	</td>
	</tr> 
	<%
	mySet_tb_Webs.movenext
loop
 
' Close Recordset
mySet_tb_Webs.close
Set mySet_tb_Webs=Nothing
' Close Connection 	
myConnection.Close
set myConnection = Nothing

%> 

<%
' PAGES LIST
%>

<tr>
<td align="left" valign="middle" colspan="3" bgcolor="<%=myApplicationColor%>"> 
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myApplicationTextColor%>"><b><%=myMessage_Page%>(S) :&nbsp; 
<%
myNbrPage=int((i+myMaxRspByPage-1)/myMaxrspbyPage)
indice=1
do While not indice>myNbrPage 
	if CInt(indice)=CInt(myNumPage) then
		%>
		[<%=indice%>]&nbsp; 
		<%
	else
		%>
		<a href="__Webs_List.asp?page=<%=indice%>&search=<%=mySearch%>&order=<%=myOrder%>"><Font Color="<%=myApplicationTextColor%>">[<%=indice%>]</FONT></a>&nbsp;
	<%
	end if	
	indice=indice+1
loop
%>
&nbsp;</b></FONT>
</td>
</tr>

</table>

<%
' ADMINISTRATION - ADD WEB for EveryBody
%>
 
<br>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="2"><a href="__Web_Modification.asp?Action=New"><%=myMessage_Add%>&nbsp;<%=myMessage_Web%></a></font><br><br>

</td>
</TR>

</TABLE>

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
'				    End Copyright												'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 

</body>
</html>


<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>