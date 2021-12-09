<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   s- O v e r A p p s - http://www.overapps.com
'
' This program "__Projects_List.asp" is free software; you can redistribute it 
' and/or modify
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
' Doesn't work with PWS ????
%>



<%
' ----------------------------------------------------------------------------
' Name 			: 	__Projects_list.asp
' Path   		: 	/
' Version		: 1.15.0
' Description 	: Projects List
' By			: Pierre Rouarch		
' Company		: OverApps	
' Date			: October, 3, 2001
'
' Modify by		:
' Company		:
' Date			:	
' -----------------------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Projects_List.asp"

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
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET APPLICATION TITLE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myApplication_Title = Get_Application_Title(myPage_Application)



Dim myAction

Dim mySearch, myMaxRspByPage, mySortProject_Type_ID, mySortProject_Name, mySortProject_Presentation, mySortProject_Leader_Pseudo, myOrder, myRs,  myNumPage, myNbrPage, indice  


Dim  MyProject_Site_ID, myProject_Member_ID, myProject_Parent_ID, myProject_ID, myProject_Path, myProject_Name, myProject_Presentation, myProject_Leader_ID, myProject_Leader_Pseudo



Dim myProject_Type_ID, myProject_Type_Name, myProject_Theme_ID 
Dim  myInfo, myModif, myList


Dim mySQL_Select_tb_Projects, mySet_tb_Projects,myMembers_Public_Type_ID
' For extensions Purpose
Dim  mySQL_Select_tb_Projects_Members, mySet_tb_Projects_Members
Dim  mySQL_Select_tb_Projects_Themes, mySet_tb_Projects_Themes
Dim  mySQL_Select_tb_Projects_types, mySet_tb_Projects_types

Dim i, j

''''''''''''''''''''''''''''''''''''''''' 
' Get Parameters 						'
'''''''''''''''''''''''''''''''''''''''''


myAction = request.querystring("Action")
myProject_ID = request.querystring("Project_ID")


myNumPage=Request("NumPage")
if Len(myNumPage)=0 then 
		myNumPage=1
end if
mySearch=Replace(Request.querystring("Search"),"'","''")
if mySearch="" then
	mySearch=Replace(Request.form("Search"),"'","''")
end if

myOrder = Request.QueryString("Order")

%>






<html>

<head>
<title><%=mySite_Name%> - Projects List</title>
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
' CENTER - LEFT
%>

<TD WIDTH="<%=myLeft_Width%>">
<!-- #include file="_borders/Left.asp" -->
</td>

<%
' CENTER - APPLICATION
%> 

<%

myMaxRspByPage=10

mySortProject_Name = "<a href=""__Projects_List.asp?order=Project_Name&NumPage="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Name&"</font></a>"



mySortProject_Presentation = "<a href=""__Projects_List.asp?order=Project_Presentation&NumPage="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</font></a>"

myMembers_Public_Type_ID = Get_Application_Public_Type_ID("Members")

 
mySortProject_Leader_Pseudo = "<a href=""__Projects_List.asp?order=tb_Sites_Members.Member_Pseudo&NumPage="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Leader&"</font></a>"


' Sort Method
Select case myOrder
	case "Project_Name"
		mySortProject_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Name&"</FONT>"


	case "Project_Presentation"
		mySortProject_Presentation = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</FONT>"

	case "tb_Sites_Members.Member_Pseudo"
		mySortProject_Leader_Pseudo = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Leader&"</FONT>"

	case else
		myOrder="Project_Name"
		mySortProject_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Name&"</FONT>"
End Select

' DB CONNECTION
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

' Read tb_projects
mySQL_Select_tb_Projects = "SELECT tb_projects.*,  tb_sites_Members.Member_Pseudo as Project_Leader_Pseudo FROM tb_Projects INNER JOIN tb_Sites_Members on tb_Projects.Project_leader_ID=tb_Sites_Members.Member_ID WHERE tb_Projects.Site_ID ="& mySite_ID


if mySearch<>"" then
	mySQL_Select_tb_Projects=mySQL_Select_tb_Projects & " AND (Project_Name LIKE '%"&mySearch&"%' OR Project_Presentation LIKE '%"&mySearch&"%'  OR  tb_Sites_Members.Member_Pseudo LIKE '%"&mySearch&"%')"
end if


If myOrder <> "Project_Name" Then
 mySQL_Select_tb_Projects=mySQL_Select_tb_Projects & " ORDER BY " & myOrder &", tb_Projects.Project_Name"
else
 mySQL_Select_tb_Projects=mySQL_Select_tb_Projects & " ORDER BY tb_Projects.Project_Name"
end if


	set mySet_tb_Projects = myConnection.Execute(mySQL_Select_tb_Projects)

%> 

<%
' SEARCH BOX
%>


 <TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></TD></TR> 
</table><table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0"> 
<tr ALIGN="CENTER"> <td> <form method="post" action="<%=myPage%>" id=form1 name=form1> 
<br> &nbsp; <input type="text" name="Search" size="30" value="<%=mySearch%>"> &nbsp; <INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"> 
</form></td></tr> </table><%
' /SEARCH BOX
%> <BR> <%
' LIST
%> <%
i=0
myRs=(myNumPage-1)*myMaxRspByPage
j=0
if not mySet_tb_Projects.bof then mySet_tb_Projects.MoveFirst
do while not mySet_tb_Projects.eof 
i=i+1
mySet_tb_Projects.movenext
loop 
if not mySet_tb_Projects.bof then 
mySet_tb_Projects.MoveFirst
mySet_tb_Projects.Move(myRs) 
end if
%> <table border="0" cellpadding="5" cellspacing="1" WIDTH="<%=myApplication_Width%>"> 
<%
' LIST Header
%> <tr> <td valign="top" align="left" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2"><% = mySortProject_Name %> 
</font></b></td><td align="left" valign="top" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2"><% = mySortProject_Presentation %> 
</font></b></td><td align="left" valign="top" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2">

<% = mySortProject_Leader_Pseudo %> 


</font></b></td><td valign="top" align="left" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
<FONT color="<%=myBorderTextColor%>"><%=myMessage_More%></font> </font> </b></td></tr> 
<%
' /LIST Header
%> <%
' LIST Informations 

IF mySet_tb_Projects.eof Then
%> 

<tr><td align="center" valign="middle" colspan=4>
<font face=Arial size=2><i><%= myMessage_No_project %></i></font>
</td></tr>
<%	
end if
do while not mySet_tb_Projects.eof AND (myMaxRspByPage>j)

        j=j+1
		myProject_Site_ID	= mySet_tb_Projects("Site_ID")
		myProject_Member_ID	= mySet_tb_Projects("Member_ID")
		myProject_ID   = mySet_tb_Projects("Project_ID")
		myProject_Name = mySet_tb_Projects("Project_Name")
		myProject_Presentation = mySet_tb_Projects("Project_Presentation")
		myProject_Leader_ID	= mySet_tb_Projects("Project_Leader_ID")	
		myProject_Leader_Pseudo	= mySet_tb_Projects("Project_Leader_Pseudo")	




'  INFORMATION


		myInfo  = "<a href=""__Project_Information.asp?Project_ID=" & myProject_ID & """>" & "<img border=""0"" src=""images/overapps-info.gif"" WIDTH=""20"" HEIGHT=""20"" " & " alt="" " & myProject_Name & """></a>"


' Can be Modify or Delete by Author, Leader or Administrator


if myProject_Member_ID=myUser_ID or myProject_Leader_ID=myUser_ID or myUser_type_ID=1 then 

		myModif = "<a href=""__Project_Modification.asp?Project_ID=" & myProject_ID & """>" 	& "<img border=""0"" src=""images/overapps-update.gif"" WIDTH=""20"" HEIGHT=""22"" " & " alt="" " & myProject_Name & """></a>"

else 
		myModif=""
end if 


%> <tr><td align="left" valign="middle"> <p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><a href="__Phases_List.asp?Project_ID=<%=myProject_ID%>"><%=myProject_Name %></a> 
</font></td><td valign="middle" align="left"> <p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><% = myProject_Presentation %> 
</font> </td><td valign="middle" align="left"> <p align="left"><font face="Arial, Helvetica, sans-serif" size="2">
<%if myUser_type_ID <= myMembers_Public_type_ID then%>
<a href="__Site_Member_Information.asp?Site_ID=<%=myProject_Site_ID%>&Member_ID=<%=myProject_Member_ID%>"><% = myProject_Leader_Pseudo %></A>
<%else%>
<% = myProject_Leader_Pseudo %>
<%end if%>

 
</font> </td><td valign="middle" align="right"> <font face="Arial, Helvetica, sans-serif" size="2"><%=myInfo%> 
&nbsp;&nbsp;<%=myModif%> </font> </td><% 		
mySet_tb_Projects.movenext
	loop %> <%
' PAGES LIST
%> <tr><td align="left" valign="middle" colspan="4" bgcolor="<%=myApplicationColor%>"> 
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2"><b><Font Color="<%=myApplicationTextColor%>">PAGE(S) 
:&nbsp;</FONT> <%
myNbrPage=int((i+myMaxRspByPage-1)/myMaxrspbyPage)
          indice=1
          do While not indice>myNbrPage 
			if CInt(indice)=CInt(myNumPage) then
          %><Font Color="<%=myApplicationTextColor%>">[<%=indice%>]&nbsp; </FONT><%else%> 
<a href="__Projects_List.asp?List=<%=myList%>&NumPage=<%=indice%>&search=<%=mySearch%>&oder=<%=myOrder%>"><Font Color="<%=myApplicationTextColor%>">[<%=indice%>]</FONT></a>&nbsp; 
<%
			end if	
			indice=indice+1
          loop
          %>&nbsp;</b></FONT> </td></tr> </table>

<%
' ADMINISTRATION - ADD Project for EveryBody
%> 


<Table cellpadding="3">
<tr>
<td>
<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="__Project_Modification.asp?Action=New"><%=myMessage_Add%>&nbsp;<%=myMessage_Project%></a></font> 

</td>
</TR>

</TABLE>


</td>
</TR>

</TABLE>

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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> 
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</FONT></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' End Copyright																	'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</body>
</html>

<% 
	myConnection.Close
	set myConnection = Nothing
%>
