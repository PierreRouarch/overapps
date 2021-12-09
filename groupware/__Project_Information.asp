<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - O&v&erA&pps - http://www.overapps.com
'
' This program "__Project_Information.asp" is free software; 
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
'-----------------------------------------------------------------------------
%>
<% 	Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Does n't Work With PWS ???
%>

<%
'------------------------------------------------------------
' Name			: __Project_Information.asp
' Path		    : /Projects
' Version 		: 1.15.0
' Description 	: Project Presentation
' By			: Pierre Rouarch
' Company		: OverApps
' Date			: October 4, 2001
'
' Modify by		:
' Company		:
' Date			: 	
'------------------------------------------------------------
Dim myPage
myPage = "__Project_Information.asp"


Dim myPage_Application
myPage_Application="Projects"
	
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


'-------------------------------------------------------
' LOCAL VARIABLES DEFINITIONS
'-------------------------------------------------------

Dim myProject_Site_ID, myProject_Site_URL, myProject_Member_ID, myProject_Member_Pseudo, myProject_Parent_ID, myProject_ID, myProject_Name,  myProject_Presentation,  myProject_Date_Beginning, myProject_Date_End, myProject_Date_Beginning2, myProject_Date_End2,  myProject_Status_ID, myProject_Status_Name, myProject_Leader_ID, myProject_Leader_Pseudo, myProject_Priority_ID, myProject_Priority_Name, myProject_Progress, myProject_Personnal, myProject_Date_Update, myProject_Author_Update

Dim myProject_Public, myProject_Site_Top

Dim  mySQL_Select_tb_Projects, mySet_tb_Projects, mySQL_Select_tb_Projects_sites, mySet_tb_Projects_sites, mySQL_Select_tb_Projects_Types, mySet_tb_Projects_Types, mySQL_Select_tb_Projects_themes, mySet_tb_Projects_themes, mySQL_Select_tb_Projects_status, mySet_tb_Projects_Status, mySQL_Select_tb_Projects_Priorities, mySet_tb_Projects_Priorities


Dim myList, myAction, myNumPage, mySearch


''''''''''''''''''''''''''''''''''''''''' 
' Get Parameters						'
'''''''''''''''''''''''''''''''''''''''''

myList = ""
myList = request.querystring("List")
myAction = request.querystring("Action")
myProject_ID = request.querystring("Project_ID")
myNumPage=Request("NumPage")
if Len(myNumPage)=0 then 
		myNumPage=1
end if
mySearch=Replace(Request.querystring("search"),"'","''")
if mySearch="" then
	mySearch=Replace(Request.form("search"),"'","''")
end if
		
' Get Project ID
myProject_ID = Request.QueryString("Project_ID")
' if nothing go Back
if len(myProject_ID & " ") = 1 then
	Response.Redirect("__Projects_List.asp?List="&myList&"&Numpage="&myNumPage&"&Search="&mySearch&"")
end if

%>


<html>

<head>
<title><%=mySite_Name%> - Project - Information </title>
</head>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%>

<!-- #include file="_borders/Top.asp" -->

<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BGColor="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP">

<%
' CENTER LEFT
%>

<TD WIDTH="<%=myLeft_Width%>">
<!-- #include file="_borders/Left.asp" -->
</td>

<%
' CENTER APPLICATION
%>

<%
' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


' READ Projects INFORMATION IN DB
mySQL_Select_tb_Projects = "SELECT * FROM tb_Projects WHERE Project_ID="&myProject_ID
set mySet_tb_Projects = myConnection.Execute(mySQL_Select_tb_Projects)

' if eof go back
if mySet_tb_Projects.eof then
	' Close Recordset
	mySet_tb_Projects.close
	Set mySet_tb_Projects=Nothing
	' Close Connection
	myConnection.close
	set myConnection = nothing
	Response.Redirect("__Projects_List.asp?List="&myList&"&Numpage="&myNumPage&"&Search="&mySearch&"")
end if

' Read Information
myProject_Site_ID = mySet_tb_Projects("Site_ID")
myProject_Member_ID = mySet_tb_Projects("Member_ID")
myProject_ID = mySet_tb_Projects("Project_ID")
myProject_Parent_ID = mySet_tb_Projects("Project_Parent_ID")
myProject_ID = mySet_tb_Projects("Project_ID")
myProject_Name  = mySet_tb_Projects("Project_Name")
myProject_Presentation = mySet_tb_Projects("Project_Presentation")
myProject_Date_Beginning = mySet_tb_Projects("Project_Date_Beginning")	
myProject_Date_End = mySet_tb_Projects("Project_Date_End")	
myProject_Date_Beginning2 = mySet_tb_Projects("Project_Date_Beginning2")	
myProject_Date_End2 = mySet_tb_Projects("Project_Date_End2")	
myProject_Status_ID = mySet_tb_Projects("Project_Status_ID")
myProject_Leader_ID = mySet_tb_Projects("Project_Leader_ID")
if len(myProject_Leader_ID)=0 then 
	myProject_Leader_ID=0
end if 
myProject_Priority_ID = mySet_tb_Projects("Project_Priority_ID")
if len(myProject_Priority_ID)=0 then
	myProject_Priority_ID=0
end if
myProject_Progress = mySet_tb_Projects("Project_Progress")
if len(myProject_Progress)=0 then 
	myProject_Progress=0
end if 
myProject_Personnal = mySet_tb_Projects("Project_Personnal")
if len(myProject_Personnal)=0 then 
	myProject_Personnal=False
end if 
myProject_Date_Update     = mySet_tb_Projects("Project_Date_Update")
myProject_Author_Update   = mySet_tb_Projects("Project_Author_Update")
	

''''''''''''''''''''''''''''''''''''''''''
' Author Pseudo							 '
''''''''''''''''''''''''''''''''''''''''''

mySQL_Select_tb_sites_members = "SELECT * FROM tb_Sites_members WHERE Member_ID = "&myProject_Member_ID 
set mySet_tb_sites_members = myConnection.Execute(mySQL_Select_tb_sites_members)
if not mySet_tb_sites_members.eof then
	myProject_Member_Pseudo = mySet_tb_sites_members("Member_Pseudo")
end if  

'''''''''''''''''''''''''''''''''
'	Status						'	
'''''''''''''''''''''''''''''''''

mySQL_Select_tb_Projects_Status = "SELECT * FROM tb_Projects_Status WHERE Site_ID="&myProject_Site_ID 

set mySet_tb_Projects_Status = myConnection.Execute(mySQL_Select_tb_Projects_Status)

if  not mySet_tb_Projects_Status.eof then
	myProject_Status_Name = mySet_tb_Projects_Status("Project_Status_Name")
else
	myProject_Status_Name = ""
end if

'''''''''''''''''''''''''''''''''
'	Leader's Pseudo			 	'	
'''''''''''''''''''''''''''''''''

mySQL_Select_tb_sites_members = "SELECT * FROM tb_Sites_members WHERE Member_ID = "&myProject_Leader_ID 
set mySet_tb_sites_members = myConnection.Execute(mySQL_Select_tb_sites_members)
if not mySet_tb_sites_members.eof then
	myProject_Leader_Pseudo = mySet_tb_sites_members("Member_Pseudo")
else 
	myProject_Leader_Pseudo = ""
end if


'''''''''''''''''''''''''''''''''
'	Priority					'	
'''''''''''''''''''''''''''''''''
mySQL_Select_tb_Projects_Priorities = "SELECT * FROM tb_Projects_Priorities WHERE Site_ID = "&myProject_Site_ID 

set mySet_tb_Projects_Priorities = myConnection.Execute(mySQL_Select_tb_Projects_Priorities)
if  not mySet_tb_Projects_Priorities.eof then
	myProject_Priority_Name = mySet_tb_Projects_Priorities("Project_Priority_Name")
else
	myProject_Priority_Name = ""
end if
	 
' Close Connection
myConnection.close
set myConnection = nothing



	
%> <td valign="top" Width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" > 
<table border="0" cellpadding="5" cellspacing="1"> <%
' TITLE
%> <tr> <td align="center" colspan="2" bgcolor="<%=myApplicationColor%>" Width="<%=myApplication_Width%>"><b><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><%=myMessage_Project%></font></b></td></tr> 
<%
' TITLE
%> 

<tr>
<td align="right" bgcolor="<%=myBorderColor%>"  >
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Title%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><strong><%=myProject_Name%>
</strong>
</font>
</td>
</tr> 

<%
' Pseudo
%>


<tr><td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Author%> 
</font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myProject_Member_Pseudo%></font>
</td>
</tr> 

<%
' Presentation
%>


<tr>
<td valign="top" align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Presentation%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myProject_Presentation%></font>
</td>
</tr>
 
<%
' Date Beginning
%>

<tr>
<td valign="top" align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Beginning%> 
</font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myDAte_Display(myProject_Date_Beginning,1)%></font>
</td>
</tr>


<%
' Date End
%>
 
<tr>
<td valign="top" align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_end%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myDate_Display(myProject_Date_End,1)%></font>
</td>
</tr>


<%
' Leader
%>

<tr>
<td valign="top" align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Leader%></font></b>
</td>
<td valign="top" align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myProject_Leader_Pseudo%></font>
</td>
</tr>

<%
' Date Author
%>
 
<tr bgcolor="<%=myApplicationColor%>">
<td  colspan="2" align="center">
<font face="Arial, Helvetica, sans-serif" size="1" Color="<%=myApplicationTextColor%>"> 
<%=myDate_Display(myProject_Date_Update,2)%> -- <%=myProject_Author_Update%></font>
</td>
</tr>
 
</table>

<%
' NAVIGATION
%>

<table border="0" width="90%" cellpadding="3" cellspacing="0">
<tr>
<td>
<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="__Projects_List.asp?List=<%=myList%>&Numpage=<%=myNumPage%>&Search=<%=mySearch%>"><%=myMessage_Project%>s</a>,&nbsp;<a href="__Phases_List.asp?Project_ID=<%=myProject_ID%>"><%=myMessage_Phase%>s</a></font>
</td>
</tr> 
<tr>
<td>
<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="__Project_Modification.asp?Action=New"><%=myMessage_Add%>&nbsp;<%=myMessage_Project%></a></font> 
</td>
</tr>

<%
' Can be Modify or Delete by Author, Leader or Administrator
%>

<% if myProject_Member_ID=myUser_ID or myProject_Leader_ID=myUser_ID or myUser_type_ID=1 then %> 
	<tr>
	<td>
	<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="__Project_Modification.asp?Project_ID=<%=myProject_ID%>"><%=myMessage_Modify%></a> 
, <a href="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Project_Modification.asp?action=Delete&Project_ID=<%=myProject_ID%>'"><%=myMessage_Delete%></a></font> 
	</td>
	</tr>
	<% 
End IF
%>
</table>
</td>
</TR>

</TABLE>

<!-- #include file="_borders/Down.asp" --> 

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
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</FONT></A> & contributors
</FONT>
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