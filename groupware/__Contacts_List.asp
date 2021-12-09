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
' Doesn't work with PWS
%>


<%
' ------------------------------------------------------------------
' Name 			: __Contacts_List.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Contacts List, Only One Directory in this Version
' By			: Pierre Rouarch
' Company		: OverApps
' Date 			: September, 20, 2001
'
' Modify by		:
' Company		:
' Date			:
' --------------------------------------------------------------------

Dim myPage
myPage = "__Contacts_List.asp"

Dim myPage_Application
myPage_Application="Contacts"
	
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


Dim myMaxRspByPage, mySortDirectory_ID, mySortContact_Name, mySortContact_Company, myOrder, myRs,  myNumPage, myNbrPage, indice, myInfo, myModif

Dim myContact_Site_ID, myContact_Member_ID, myDirectory_ID, myDirectory_Name, myContact_ID, myContact_Name, myContact_FirstName, myContact_Company, myContact_Company_Activity, myContact_Company_Activity_ID, myContact_Company_Zip, myContact_Company_City, myContact_Company_Phone, myContact_Company_Mobile, myContact_Company_Email,  myContact_Home_Zip, myContact_Home_City, myContact_Home_Phone, myContact_Home_Mobile, myContact_Home_Email

Dim myContact_City


Dim  myMethod_Search, mySearch_Contact_Company, mySearch_Contact_Company_Activity, mySearch_Contact_Company_Activity_ID, mySearch_Contact_Name, mySearch_Contact_City

Dim mySQL_Select_tb_Contacts, mySet_tb_Contacts, mySQL_Select_tb_Directories, mySet_tb_Directories

Dim i, j

myMaxRspByPage=10


' FORCE TO DIRECTORY 1
myDirectory_ID=1

%>
<html>

<head>
<title><%=mySite_Name%> - Contacts List</title>
</head>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%> 

<!-- #include file="_borders/Top.asp" --> 

<%
' CENTER
%> 

<TABLE WIDTH="<%=myGlobal_Width%>" BGColor=<%=myBorderColor%> BORDER="0" CELLPADDING="0" CELLSPACING="0">
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
' Company or Individual or Both
myMethod_Search = Replace(Request("Method_Search"),"'","''")
if len(myMethod_Search)=0 then
	myMethod_Search="Both"
end if
mySearch_Contact_Company=Replace(Request("Contact_Company"),"'","''")
mySearch_Contact_Name=Replace(Request("Contact_Name"),"'","''")
mySearch_Contact_City=Replace(Request("Contact_City"),"'","''")
mySearch_Contact_Company_Activity_ID=Replace(Request("Contact_Company_Activity_ID"),"'","''")
myNumPage=Request("NumPage")
if Len(myNumPage)=0 then
	myNumPage=1
end if
myOrder = Request("order")
' For Future Extension
mySearch_Contact_Company_Activity=Replace(Request("Contact_Company_Activity"),"'","''")
	if len(mySearch_Contact_Company_Activity)>0 then
Response.redirect("__Contacts_Activities_List.asp?Method_Search="&myMethod_Search&"&Contact_Company="&myContact_Company&"&&Contact_Name="&myContact_Name&"&Contact_City="&myContact_City&"&Contact_Company_Activiy="&myContact_Company_Activity&"&NumPage="&myNumPage&"&Order="&myOrder&"")
		end if



set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

' Not Used
mySortDirectory_ID = "<a href=""__Contacts_List.asp?Method_Search="&myMethod_Search&"&Contact_Company="&mySearch_Contact_Company&"&Contact_Name="&mySearch_Contact_Name&"&Contact_City="&mySearch_Contact_City&"&Contact_Company_Activiy_ID="&mySearch_Contact_Company_Activity_ID&"&NumPage="&myNumPage&"&Order=Contact_Directory_ID""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Directory&"</Font></a>"


mySortContact_Name = "<a href=""__Contacts_List.asp?Method_Search="&myMethod_Search&"&Contact_Company="&mySearch_Contact_Company&"&Contact_Name="&mySearch_Contact_Name&"&Contact_City="&mySearch_Contact_City&"&Contact_Company_Activiy_ID="&mySearch_Contact_Company_Activity_ID&"&NumPage="&myNumPage&"&Order=Contact_Name""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Contact&"</font></a>"

mySortContact_Company = "<a href=""__Contacts_List.asp?Method_Search="&myMethod_Search&"&Contact_Company="&mySearch_Contact_Company&"&Contact_Name="&mySearch_Contact_Name&"&Contact_City="&mySearch_Contact_City&"&Contact_Company_Activiy_ID="&mySearch_Contact_Company_Activity_ID&"&NumPage="&myNumPage&"&Order=Contact_Company""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Company&"</font></a>"

'  Sort Method

Select case myOrder

' Not USed
	case "Directory_ID"

		myOrder="tb_contacts.Directory_ID"
		mySortDirectory_ID = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Directory&"</FONT>"

	case "Contact_Name"

		myOrder="Contact_Name"
		mySortContact_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Contact&"</FONT>"

	case "Contact_Company"

		myOrder="Contact_Company"
		mySortContact_Company = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Company&"</FONT>"

	case else

		myOrder="Contact_Name"
		mySortContact_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Contact&"</FONT>"

End Select

		' Reading Contact List




' tb_Directories_Sites for multi sites and multi directories Extensions Purpose


mySQL_Select_tb_Contacts = "SELECT * FROM tb_contacts INNER JOIN tb_Directories_sites ON tb_contacts.Directory_ID = tb_Directories_sites.Directory_Id WHERE tb_Directories_sites.Site_ID="&mySite_ID



' Avoid Error Condition
if mySearch_Contact_Company<>"" AND myMethod_Search="Home" then
	myMethod_Search="Both"
end if



' Search Company,  Individual or both ?

''''''''''''''''''''''''''''''''''''''''
' COMPANY
'''''''''''''''''''''''''''''''''''''''

if myMethod_Search="Company" then

	mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " AND Contact_Company_Type = 1 "

		if mySearch_Contact_Company<>"" then
			mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " AND Contact_Company LIKE '%"&mySearch_Contact_Company&"%'"
		end if

		if mySearch_Contact_City<>""  then
				mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " AND Contact_Company_City LIKE '%"&mySearch_Contact_City&"%'"
		end if

end if


''''''''''''''''''''''''''''''''''''''''
' Home
'''''''''''''''''''''''''''''''''''''''

if myMethod_Search="Home" then
	mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " AND Contact_Home_Type = 1 "

		if mySearch_Contact_City<>""  then
			mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " AND Contact_Home_City LIKE '%"&mySearch_Contact_City&"%'"
		end if

end if

''''''''''''''''''''''''''''''''''''''''
' BOTH
'''''''''''''''''''''''''''''''''''''''

if myMethod_Search="Both"  then
	mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts 
	'& " AND( Contact_Company_Type = 1 OR Contact_Home_Type = 1) "

		if mySearch_Contact_Company<>"" then
			mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " AND Contact_Company LIKE '%"&mySearch_Contact_Company&"%'"
		end if


		if mySearch_Contact_City<>"" then
				mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " AND (Contact_Company_City LIKE '%"&mySearch_Contact_City&"%' OR Contact_Home_City LIKE '%"&mySearch_Contact_City&"%')"
		end if


end if



		if mySearch_Contact_Name<>"" then
			mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " AND Contact_Name LIKE '%"&mySearch_Contact_Name&"%'"
		end if


If myOrder <> "Contact_Name" Then
	mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " ORDER BY " & myOrder &", Contact_Name"
else
	mySQL_Select_tb_Contacts=mySQL_Select_tb_Contacts & " ORDER BY  Contact_Name"
end if	
	
	
	'response.write mySQL_Select_tb_Contacts
	'response.end

		set mySet_tb_Contacts = myConnection.Execute(mySQL_Select_tb_Contacts)

%> 

<TD BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left">

<TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" BGCOLOR="<%=myApplicationColor%>">
<TR>
<TD>

<font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font>
</TD>
</TR> 
</TABLE>



<%
' CENTER APPLICATION SEARCH FORM
%> 
<table border="0" WIDTH="<%=myApplication_Width%>" >
<form method="post" action="<%=myPage%>" id=form1 name=form1>
<tr BGCOLOR="<%=myBgColor%>" ALIGN="Left">
<td WIDTH="89">
<font face="Arial,Helvetica" size="1" color="<%=myBGTextColor%>"><b><%=myMessage_Company%>
:</b></font>
</td>
<td WIDTH="139">
<input type="text" name="Contact_Company" size="20" VALUE="<%=mySearch_Contact_Company%>">
</td>
<td WIDTH="60">
<FONT FACE="Arial,Helvetica" SIZE="1" color="<%=myBGTextColor%>"><B><%=myMessage_City%>
:</B></FONT>
</td>

<td WIDTH="37">
<INPUT TYPE="text" NAME="Contact_City" SIZE="20" VALUE="<%=mySearch_Contact_City%>" >
</td>
</tr>


<tr BGCOLOR="<%=myBGColor%>" ALIGN="LEFT">
<td WIDTH="89">
<font face="Arial,Helvetica" size="1" color="<%=myBGTextColor%>"><B><%=myMessage_Name%>
:</B></font>
</td>

<td WIDTH="139">
<INPUT TYPE="text" NAME="Contact_Name" SIZE="20" VALUE="<%=mySearch_Contact_Name%>">
</td>
<td WIDTH="60">
</td>
<td WIDTH="37">
</td>
</tr> 


<tr BGCOLOR="<%=myBGColor%>" ALIGN="right">
<td colspan="4"> 
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" color="<%=myBGTextColor%>"><%=myMessage_You_Search_For%>
: <%=myMessage_A_Professionnal%><INPUT TYPE="radio" NAME="Method_Search" VALUE="Company" <% if myMethod_Search="Company" then %>  CHECKED  <%end if%> >,
<%=myMessage_An_Individual%> <INPUT TYPE="radio" NAME="Method_Search" VALUE="Home"  <% if myMethod_Search="Home" then %>  CHECKED  <%end if%>>
, <%=myMessage_Both%> <INPUT TYPE="radio" NAME="Method_Search" VALUE="Both"  <% if myMethod_Search="Both" then %>  CHECKED  <%end if%>></FONT>
</td>
</tr>

<tr BGCOLOR="<%=myBGColor%>" ALIGN="right">
<td colspan="4">
 &nbsp; <INPUT TYPE="submit" NAME="Submit2" VALUE="<%=myMessage_Go%>">
</td>
</tr>
</form>
</table>

<%
' CENTER APPLICATION SEARCH RESULTS
%> 


<%
i=0
myRs=(myNumPage-1)*myMaxRspByPage
j=0
if not mySet_tb_Contacts.bof then mySet_tb_Contacts.MoveFirst
do while not mySet_tb_Contacts.eof
i=i+1
mySet_tb_Contacts.movenext
loop
if not mySet_tb_Contacts.bof then
mySet_tb_Contacts.MoveFirst
mySet_tb_Contacts.Move(myRs)
end if
%> 

<table border="0" cellpadding="5" cellspacing="1" WIDTH="<%=myApplication_Width%>">
<tr> 

<% if myMethod_Search<>"Home" then %>
<td valign="top" align="left" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=mySortContact_Name %>
</font></b>
</td>
<td align="left" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=mySortContact_Company %></font></b></td>

<% else %>
<td valign="top" align="left" bgcolor="<%=myBorderColor%>" colspan="2">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><% = mySortContact_Name %>
</font></b>
</td>
<% end if %>

 <td align="left" valign="top" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>">

<% if myMethod_Search="Home" then %> <%=myMessage_Phone%>s&nbsp;(<%=myMessage_Home%>)
<% else %> <%=myMessage_Phone%>s&nbsp;(<%=myMessage_Office%>)<% end if %> </font></b>
</td>
<td valign="top" align="left" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_More%></font></b>
</td>
</tr>
<%	do while not mySet_tb_Contacts.eof AND (myMaxRspByPage>j)
        j=j+1

		myDirectory_ID = mySet_tb_Contacts("Directory_ID")
		myContact_Site_ID= mySet_tb_Contacts("Site_ID")
		myContact_Member_ID=mySet_tb_Contacts("Member_ID")
		myContact_ID = mySet_tb_Contacts("Contact_ID")
		myContact_Name = mySet_tb_Contacts("Contact_Name")
		myContact_FirstName = mySet_tb_Contacts("Contact_FirstName")
		myContact_Company = mySet_tb_Contacts("Contact_Company")
		myContact_Company_Zip = mySet_tb_Contacts("Contact_Company_Zip")
		myContact_Company_City = mySet_tb_Contacts("Contact_Company_City")
		myContact_Company_Phone   = mySet_tb_Contacts("Contact_Company_Phone")
		myContact_Company_Mobile = mySet_tb_Contacts("Contact_Company_Mobile")
		myContact_Company_Email = mySet_tb_Contacts("Contact_Company_Email")
		myContact_Home_Zip = mySet_tb_Contacts("Contact_Home_Zip")
		myContact_Home_City = mySet_tb_Contacts("Contact_Home_City")
		myContact_Home_Phone   = mySet_tb_Contacts("Contact_Home_Phone")
		myContact_Home_Mobile = mySet_tb_Contacts("Contact_Home_Mobile")
		myContact_Home_Email = mySet_tb_Contacts("Contact_Home_Email")



		myInfo  = "<a href=""__Contact_Information.asp?Contact_ID="&myContact_ID&""">" & "<img border=""0"" src=""images/overapps-info.gif"" WIDTH=""20"" HEIGHT=""20"" " & " alt=""  " & myContact_Name & """></a>"
		myModif = "<a href=""__Contact_Modification.asp?Action=Update&Contact_ID=" & myContact_ID & """>" 	& "<img border=""0"" src=""images/overapps-update.gif"" WIDTH=""20"" HEIGHT=""22"" " & " alt="" " & myContact_Name & """></a>"
%> 

<tr> 
<%if myMethod_Search<>"Home" then %> 
<td align="left" valign="middle">
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2">
<a href="mailto:<% = myContact_Company_Email %>"><img src="Images/OverApps-mail1.gif" alt="<%= myContact_FirstName %>&nbsp;  <%= myContact_Name %>" align="absmiddle" border="0" WIDTH="22" HEIGHT="22"></a>
<strong><%=myContact_FirstName%>&nbsp;<%=myContact_Name%></strong></font>
</td>
<td valign="middle" align="left">
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><%= myContact_Company %>
</font>
</td>
<td align="left" valign="middle" nowrap>
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><%=myMessage_Office%>
: <% = myContact_Company_Phone %><br> <%=myMessage_Mobile%> : <% = myContact_Company_Mobile %>
</font>
</td>
<%else%> 
<td align="left" valign="middle" >
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2">
<a href="mailto:<% = myContact_Home_Email %>"><img src="Images/OverApps-mail1.gif" alt=" <%=myContact_Name%>,&nbsp; <%=myContact_FirstName%>" align="absmiddle" border="0" WIDTH="22" HEIGHT="22"></a>
<strong><%=myContact_Name%>,&nbsp;<%=myContact_FirstName%></strong></font>
</td>
<td align="left" valign="middle" Colspan="2" nowrap>
<p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><%=myMessage_Home%> : <% = myContact_Home_Phone %><br> <%=myMessage_Mobile%> : <% = myContact_Home_Mobile %></font> 
</td>
<% end if %>
<td align="right" valign="middle" >
<% = myInfo %> &nbsp;&nbsp;<%=myModif%>
</td>
</tr>

<% 		mySet_tb_Contacts.movenext
	loop %> 

<tr>
<td align="left" valign="middle" colspan="5" bgcolor="<%=myApplicationColor%>">
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myApplicationTextColor%>"><b>PAGE(S)
:&nbsp; <%myNbrPage=int((i+myMaxRspByPage-1)/myMaxrspbyPage)
          indice=1
          do While not indice>myNbrPage 
			if CInt(indice)=CInt(myNumPage) then
          %>[<%=indice%>]&nbsp; <%else%> <a href="__Contacts_List.asp?Method_Search=<%=myMethod_Search%>&Contact_Company=<%=mySearch_Contact_Company%>&Contact_Name=<%=mySearch_Contact_Name%>&Contact_City=<%=mySearch_Contact_City%>&NumPage=<%=indice%>&Order=<%=myOrder%>">[<%=indice%>]</a>&nbsp;
<%
			end if
			indice=indice+1
          loop
          %>&nbsp;</b></FONT> 

</td>
</tr>
</table>


<%
' ADMINISTRATION FOR EVERYBODY
%> 

<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0">
<tr>
<td>
<font face="Arial, Helvetica, sans-serif" size="2"><a href="__Contact_Modification.asp?Action=New"><%=myMessage_Add%>&nbsp;<%=myMessage_Contact%></a></font>
</td>
</tr>
</table>

<%
' /ADMINISTRATION
%>

</td>

<%
' /CENTER APPLICATION
%> 

</TR> 
</TABLE>

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
' license's compliances.														'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 

<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0">
<TR ALIGN="RIGHT">
<TD>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</FONT>
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
<%
	myConnection.Close
	set myConnection = Nothing
%>

<html><script language="JavaScript"></script></html>