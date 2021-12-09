<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - O!ver!App!s - http://www.overapps.com
'
' This program "__Sites_Members_List.asp" is free software; 
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
' Does n't Work with PWS ???
%>

<%
' ------------------------------------------------------------
' Name		: __Sites_Members_List.asp
' Path		: /
' Description 	: Site Members List
' By : Pierre Rouarch	
' Company : OverApps
' Date : January, 16 , 2001
' Version : 1.15.0
' Modify by	:
' Company	:
' Date 		:
' ------------------------------------------------------------
Dim myPage
myPage = "__Sites_Members_List.asp"

Dim myPage_Application
myPage_Application="Members"
	
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


Dim mySearch, myNumPage, myMaxRspByPage, indice, myNbrPage, mySortMember_Pseudo, mySortMember_Company, myOrder, myRs

Dim myMember_ID, myMember_Pseudo, myMember_Email,  myMember_Company, myMember_Company_type,  myMember_Company_Phone, myMember_Company_Mobile, myMember_Home_type, myMember_Home_Phone, myMember_Home_Mobile

Dim i, j

%>


<HTML>
<HEAD>
</HEAD>
<TITLE>
<%=mySite_Name%> - Members List</TITLE>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">

<%
'  TOP
%>


<!-- #include file="_borders/Top.asp" --> 

<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> 
<%
' Navigation
%>

 <TD WIDTH="<%=myLeft_Width%>">
<!-- #include file="_borders/Left.asp" -->
</TD>
<%
' Application
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

		
' Get Sort Method
myOrder = Request.QueryString("Order")
if len(myOrder)=0 then 
	myOrder="Member_Pseudo"
end if

	
myMaxRspByPage=10



' Prepare Sort 

mySortMember_Pseudo = "<a href=""__Sites_Members_List.asp?order=Member_Pseudo&Page="&myNumPage&"&search="&mySearch&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Member&"</font></a>"

mySortMember_Company = "<a href=""__Sites_Members_List.asp?order=Member_Company&Page="&myNumPage&"&search="&mySearch&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Company&"</font></a>"




Select case myOrder
	case "Member_Pseudo"
		mySortMember_Pseudo = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Member&"</FONT>"
	case "Member_Company"
		mySortMember_Company = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Company&"</FONT>"
	case else
		myOrder="Member_Pseudo"
		mySortMember_Pseudo = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Member&"</FONT>"
End Select



		
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String



' Global SELECT		
mySQL_Select_tb_Sites_Members = "SELECT * FROM tb_Sites_Members WHERE Site_ID = "&mySite_ID&" "

' Search if done
if mySearch<>"" then
	mySQL_Select_tb_Sites_Members=mySQL_Select_tb_Sites_Members & " AND (Member_Name LIKE '%"&mySearch&"%'  OR Member_FirstName LIKE '%"&mySearch&"%' OR Member_Pseudo LIKE '%"&mySearch&"%' OR Member_Company LIKE '%"&mySearch&"%')"
end if

' Sort 
If myOrder <> "Member_Pseudo" Then
	mySQL_Select_tb_Sites_Members=mySQL_Select_tb_Sites_Members & " ORDER BY " & myOrder &", Member_Pseudo"
else
		mySQL_Select_tb_Sites_Members=mySQL_Select_tb_Sites_Members & " ORDER BY  Member_Pseudo"
end if		



set mySet_tb_Sites_Members = myConnection.Execute(mySQL_Select_tb_Sites_Members)
%> 

<td valign="top"  bgcolor="<%=mybgColor%>" Width="<%=myApplication_Width%>"> 
<%
i=0
myRs=(myNumPage-1)*myMaxRspByPage

j=0
if not mySet_tb_Sites_Members.bof then mySet_tb_Sites_Members.MoveFirst
do while not mySet_tb_Sites_Members.eof 
i=i+1
mySet_tb_Sites_Members.movenext
loop 
if not mySet_tb_Sites_Members.bof then
mySet_tb_Sites_Members.MoveFirst
mySet_tb_Sites_Members.Move(myRs) 
end if
%> 



<%
' Members Title
%> 

<TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" BGCOLOR="<%=myApplicationColor%>">
<TR>
<TD><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font>
</TD>
</TR> 
</TABLE>

<%
' Search Form
%> 

<TABLE WIDTH="<%=myApplication_Width%>" bgcolor="<%=myBGColor%>" BORDER="0"> 
<TR ALIGN="CENTER"> <TD> <form method="post" action="<%=myPage%>" id=form1 name=form1><br>&nbsp; 
<input type="text" name="search" size="30">&nbsp; <INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"> 
</form></TD></TR> </TABLE><%
' Listing 
%> <table WIDTH="<%=myApplication_Width%>" bgcolor="<%=myBGColor%>" border="0" cellpadding="5" cellspacing="1"> 
<%
' Listing Header
%>

<tr>

<td valign="top" align="left" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2"><%=mySortMember_Pseudo%></font></b>
</td>

<td align="left" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2"><%=mySortMember_Company%></font></b>
</td>

<td align="left" valign="top" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_Phone%>s</font></b>
</td>
<td valign="top" align="left" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_More%></font></b>
</td>
</tr> 
<%
' Listing values
%> <%	
do while not mySet_tb_Sites_Members.eof AND (myMaxRspByPage>j)
	j=j+1
	myMember_Pseudo = mySet_tb_Sites_Members("Member_Pseudo")
	myMember_ID  = mySet_tb_Sites_Members("Member_ID")
	myMember_Email = mySet_tb_Sites_Members("Member_Email")
	myMember_Company_Type  = mySet_tb_Sites_Members("Member_Company_type")	
	myMember_Company  = mySet_tb_Sites_Members("Member_Company")
	myMember_Company_Phone  = mySet_tb_Sites_Members("Member_Company_Phone")
	myMember_Company_Mobile  = mySet_tb_Sites_Members("Member_Company_Mobile")
	myMember_Home_Type  = mySet_tb_Sites_Members("Member_Home_Type")
	myMember_Home_Phone = mySet_tb_Sites_Members("Member_Home_Phone")
	myMember_Home_Mobile = mySet_tb_Sites_Members("Member_Home_Mobile")
%> <tr> <td align="left" valign="middle" > <p> <font face="Arial, Helvetica, sans-serif" size="2"> 
<a href="mailto:<%=myMember_Email%>"><img src="Images/OverApps-mail1.gif" alt="<%=myMember_Email%>" 
              align="absmiddle" border="0" WIDTH="22" HEIGHT="22"></a>&nbsp; <strong><%=myMember_Pseudo%></strong> 
</font></p></td><td valign="middle" align="left" > <font face="Arial, Helvetica, sans-serif" size="2"><%=myMember_Company%></font> 
</td><td valign="middle" align="left" > <% if myMember_Company_Type then %> <%=myMessage_Office%> 
:&nbsp; <% if len(myMember_Company_Phone)>0 then %> <font face="Arial, Helvetica, sans-serif" size="2"><%=myMember_Company_Phone%>&nbsp;&nbsp;</font> 
<%end if%> <% if len(myMember_Company_Mobile)>0 then %> <font face="Arial, Helvetica, sans-serif" size="2"><%=myMember_Company_Mobile%></font> 
<%end if%> <br> <%end if%> <% if myMember_Home_Type then %> <%=myMessage_Home%> 
:&nbsp; <% if len(myMember_Home_Phone)>0 then %> <font face="Arial, Helvetica, sans-serif" size="2"><%=myMember_Home_Phone%>&nbsp;&nbsp;</font> 
<%end if%> <% if len(myMember_Home_Mobile)>0 then %> <font face="Arial, Helvetica, sans-serif" size="2"><%=myMember_Home_Mobile%></font> 
<%end if%> <%end if%> </td><td align="right" valign="middle"> <font face="Arial, Helvetica, sans-serif" size="2"><a href="__Site_Member_Information.asp?Member_ID=<%=myMember_ID%>"><img border="0" src="images/OverApps-info.gif" WIDTH="20" HEIGHT="20" alt="<%=myMember_Pseudo%>"></a>	
&nbsp;&nbsp; <%
if (myUser_ID=myMember_ID or myUSer_Type_ID=1) then
%> <a href="__Site_Member_Modification.asp?Action=Update&Member_ID=<%=myMember_ID%>"><img border="0" src="images/OverApps-update.gif" WIDTH="20" HEIGHT="20" alt="<%=myMember_Pseudo%>"></a>	
</font> <%end if%> </td></tr> <% 		mySet_tb_Sites_Members.movenext
	loop %> <%
' Other Pages
%> <tr> <td align="left" valign="middle" colspan="4" bgcolor="<%=myApplicationColor%>">&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"><b>PAGE(S) 
:&nbsp; <%
myNbrPage=int((i+myMaxRspByPage-1)/myMaxrspbyPage)
indice=1
do While not indice>myNbrPage 
	if CInt(indice)=CInt(myNumPage) then
%> [<%=indice%>]&nbsp; <%else%> <a href="__Sites_Members_List.asp?page=<%=Indice%>&search=<%=mySearch%>&oder=<%=myOrder%>">[<%=indice%>]</a>&nbsp; 
<%
	end if	
	indice=indice+1
    loop
%> &nbsp;</b></FONT> </td></tr> </table>

<%
' ADD MEMBER
%> 
<table border="0" width="<%=myApplication_Width%>" cellpadding="5" cellspacing="0"> 
<% if (myUser_Type_ID=1) then %>
<tr> 
<td> 
<a href="__Site_Member_Modification.asp?Action=New"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Add%>&nbsp;<%=myMessage_Member%></font></a> 
</td>
</tr> 
<%end if%> 

</table>

</td>
</TR>
</TABLE>

<!-- #include file="_borders/Down.asp" --> 
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
</body>
</html>
<% 
	myConnection.Close
	set myConnection = Nothing
%>




<html><script language="JavaScript"></script></html>