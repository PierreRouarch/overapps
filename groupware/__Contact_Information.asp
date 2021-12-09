<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Contact_Information.asp" is free software; 
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
' Cache non géré par PWS
%>

<%
' ------------------------------------------------------------
' Name			: __Contact_Information.asp
' Path   		: /
' Vertsion 		: 1.15.0
' Description 	: Contact Information
' By			: Pierre Rouarch												
' Company		: OverApps
' Date			: December 10, 2001
' Contributor : Dania Tcherkezoff
' ------------------------------------------------------------

Dim myPage
myPage = "__Contacts_Information.asp"

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


Dim myContact_Site_ID, myContact_Member_ID, myDirectory_ID, myDirectory_Name, myContact_ID,  myContact_Company_Activity,  myContact_Title_ID, myContact_Title,  myContact_Name, myContact_FirstName, myContact_Company_Type, myContact_Company, myContact_Company_Activity_ID, myContact_Company_Address, myContact_Company_Zip, myContact_Company_City, myContact_Company_State, myContact_Company_Country, myContact_Company_Country_ID, myContact_Company_Phone, myContact_Company_Mobile, myContact_Company_Fax, myContact_Company_Email, myContact_Company_Web, myContact_Company_Fonction, myContact_Home_Type, myContact_Home_Address, myContact_Home_Zip, myContact_Home_City, myContact_Home_State, myContact_Home_Country, myContact_Home_Country_ID, myContact_Home_Phone, myContact_Home_Mobile, myContact_Home_Fax, myContact_Home_Email, myContact_Home_Web, myContact_Comments, myContact_Author_Update, myContact_Date_Update

Dim  myMethod_Search, mySearch_Contact_Company, mySearch_Contact_Company_Activity, mySearch_Contact_Company_Activity_ID, mySearch_Contact_Name, mySearch_Contact_City

Dim mySQL_Select_tb_Contacts, mySet_tb_Contacts, mySQL_Select_tb_Directories, mySet_tb_Directories, mySQL_Select_tb_Contacts_Activities, mySet_tb_Contacts_Activities, mySQL_Select_tb_Countries, mySet_tb_Countries


' Get Parameters

myContact_ID = Request.QueryString("Contact_ID")
if len(myContact_ID & " ") = 1 then
		Response.Redirect("__Contacts_List.asp")
end if

' NOT USED
'myDirectory_ID = Request.QueryString("Directory_ID")
'if len(myDirectory_ID & " ") = 1 then
'		Response.Redirect("__Contacts_List.asp")
'end if
' Force to Directory 1
myDirectory_ID=1


%>
<html>

<head>
<title><%=mySite_Name%> - Contact Information</title>
</head>
<BODY BackGround="<%=myBGImage%>" bgColor="<%=myBGColor%>" Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">

<%
' TOP
%>

<!-- #include file="_borders/Top.asp" --> 

<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR="<%=myBorderCOLOR%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
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

mySQL_Select_tb_Contacts = "SELECT * FROM tb_contacts WHERE Contact_ID = "&myContact_ID
set mySet_tb_Contacts = myConnection.Execute(mySQL_Select_tb_Contacts)
	
if not mySet_tb_Contacts.eof then

	' Get Data
	myContact_Site_ID = mySet_tb_Contacts("Site_ID")
	myContact_Member_ID = mySet_tb_Contacts("Member_ID")
	myDirectory_ID = mySet_tb_Contacts("Directory_ID")
	myContact_Title_ID = mySet_tb_Contacts("Contact_Title_ID")
	myContact_Name = mySet_tb_Contacts("Contact_Name")
	myContact_FirstName = mySet_tb_Contacts("Contact_FirstName")
	myContact_Company_Type = mySet_tb_Contacts("Contact_Company_Type")
	myContact_Company = mySet_tb_Contacts("Contact_Company")
	myContact_Company_Activity_ID = mySet_tb_Contacts("Contact_Company_Activity_ID")
	myContact_Company_Address = mySet_tb_Contacts("Contact_Company_Address")
	myContact_Company_Zip = mySet_tb_Contacts("Contact_Company_Zip")
	myContact_Company_City = mySet_tb_Contacts("Contact_Company_City")
	myContact_Company_State = mySet_tb_Contacts("Contact_Company_State")
	myContact_Company_Country_ID = mySet_tb_Contacts("Contact_Company_Country_ID")
	myContact_Company_Phone = mySet_tb_Contacts("Contact_Company_Phone")
	myContact_Company_Mobile = mySet_tb_Contacts("Contact_Company_Mobile")
	myContact_Company_Fax = mySet_tb_Contacts("Contact_Company_Fax")
	myContact_Company_Email = mySet_tb_Contacts("Contact_Company_Email")
	myContact_Company_Web = mySet_tb_Contacts("Contact_Company_Web")
	myContact_Company_Fonction = mySet_tb_Contacts("Contact_Company_Fonction")
	myContact_Home_Type = mySet_tb_Contacts("Contact_Home_Type")
	myContact_Home_Address = mySet_tb_Contacts("Contact_Home_Address")
	myContact_Home_Zip = mySet_tb_Contacts("Contact_Home_Zip")
	myContact_Home_City = mySet_tb_Contacts("Contact_Home_City")
	myContact_Home_State = mySet_tb_Contacts("Contact_Home_State")
	myContact_Home_Country_ID = mySet_tb_Contacts("Contact_Home_Country_ID")
	myContact_Home_Phone = mySet_tb_Contacts("Contact_Home_Phone")
	myContact_Home_Mobile = mySet_tb_Contacts("Contact_Home_Mobile")
	myContact_Home_Fax = mySet_tb_Contacts("Contact_Home_Fax")
	myContact_Home_Email = mySet_tb_Contacts("Contact_Home_Email")
	myContact_Home_Web = mySet_tb_Contacts("Contact_Home_Web")
	myContact_Comments = mySet_tb_Contacts("Contact_Comments")
	myContact_Author_Update = mySet_tb_Contacts("Contact_Author_Update")
	myContact_Date_Update = mySet_tb_Contacts("Contact_Date_Update")

else 
	' Close Recordset
	mySet_tb_Contacts.close
	Set mySet_tb_Contacts=Nothing
	' Close Connection
	myConnection.close
	set myConnection = nothing	
	Response.Redirect("__Contacts_List.asp")
end if


' Title
if myContact_Title_ID = 1 then myContact_Title=myMessage_Mister end if
if myContact_Title_ID = 2 then myContact_Title=myMessage_Misses end if
if myContact_Title_ID = 3 then myContact_Title=myMessage_Miss end if 

		
' Country
if myContact_Company_Type and (len(myContact_Company_Country_ID)>0 and myContact_Company_Country_ID<>0) then 
	mySQL_Select_tb_countries = "Select * FROM tb_Countries WHERE Country_ID = "&myContact_Company_Country_ID
set mySet_tb_Countries = myConnection.Execute(mySQL_Select_tb_Countries)
myContact_Company_Country = mySet_tb_Countries("Country")
' Close Recordset
mySet_tb_Countries.close
Set mySet_tb_Countries=Nothing
else 
	myContact_Company_Country_ID = 0
	myContact_Company_Country = ""	
end if

if myContact_Home_Type and (len(myContact_Home_Country_ID)>0 and myContact_Home_Country_ID<>0) then 
	mySQL_Select_tb_countries = "Select * FROM tb_Countries WHERE Country_ID = "&myContact_Home_Country_ID
set mySet_tb_Countries = myConnection.Execute(mySQL_Select_tb_Countries)
myContact_Home_Country = mySet_tb_Countries("Country")
' Close Recordset
mySet_tb_Countries.close
Set mySet_tb_Countries=Nothing
else 
	myContact_Home_Country_ID = 0
	myContact_Home_Country = ""	
end if




' Close Recordset
mySet_tb_Contacts.close
Set mySet_tb_Contacts=Nothing
' Close Connection
myConnection.close
set myConnection = nothing


%> 

<td valign="top" bgcolor="<%=mybgColor%>" Width="<%=myApplication_Width%>"> 

<table border="0" bgcolor=<%=myBGCOLOR%> cellpadding="5" cellspacing="1" WIDTH="<%=myApplication_Width%>">
 
<%
' Application Title
%>

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"> 
<%=myApplication_Title%></font></b>
</td>
</tr> 

<%
' Title - Civility
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Title%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><strong>
<%=myContact_Title%></strong></font>
</td>
</tr> 


<%
' FirstName
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_FirstName%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><strong><%=myContact_Firstname%></strong></font>
</td>
</tr>


<%
' Name
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Name%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><strong><%=myContact_Name%></strong></font>
</td>
</tr> 
 
<%
' Company
%>

<% 
if myContact_Company_Type=1 then 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 	Professionnal Contact 								'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 
<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"> 
<%=myMessage_Office%></font></b>
</td>
</tr> 


<%
' Fonction
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Fonction%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company_Fonction%></font>
</td>
</tr> 

<%
' Company
%>


<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Company%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company%></font>
</td>
</tr> 


<%
' Company Address
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Address%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company_Address%></font>
</td>
</tr> 


<%
' Company Zip Code
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Zip_Code%></font></b>
</td>
<td align="left"><font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company_Zip%></font>
</td>
</tr>
 

<%
' Company City
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_City%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company_City%></font>
</td>
</tr> 


<%
' Company State
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_State%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company_State%></font>
</td>
</tr> 


<%
' Country
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Country%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2">
<%=myContact_Company_Country%></font>
</td>
</tr> 

<%
' Company Phone
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Phone%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company_Phone%></font>
</td>
</tr> 


<%
' Company Mobile
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Mobile%></font></b></td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company_Mobile%></font>
</td>
</tr> 


<%
' Company Fax
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Fax%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Company_Fax%></font>
</td>
</tr> 


<%
' Company Email
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Email%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"> 
<a  HREF="mailto:<%=myContact_Company_Email%>"><%=myContact_Company_Email%></A></font>
</td>
</tr> 


<%
' Company Web
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Web%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><A HREF="<%=myContact_Company_Web%>"><%=myContact_Company_Web%></A></font>
</td>
</tr>
 
<% end If ' Company %> 



<%
' HOME
%>

<% 
if myContact_Home_Type=1 then 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 	Home Contact			'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"><%=myMessage_Home%></font></b>
</td>
</tr>

<%
' Home Address
%>
 
<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Address%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Home_Address%></font>
</td>
</tr>
 

<%
' Home Zip Code
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Zip_Code%></font></b></td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Home_Zip%></font>
</td>
</tr>
 
<%
' Home City
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_City%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><% = myContact_Home_City %></font>
</td>
</tr> 


<%
' Home State
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_State%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Home_State %></font>
</td>
</tr>


<%
'Home Country
%>
 
<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Country%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Home_Country %></font>
</td>
</tr>
 

<%
' Home Phone
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Phone%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Home_Phone%></font>
</td>
</tr>
 

<%
' Home Mobile
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Mobile%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Home_Mobile%></font>
</td>
</tr>
 

<%
' Home Fax
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Fax%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><%=myContact_Home_Fax%></font>
</td>
</tr>
 

<%
' Home Email
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Email%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><A HREF="mailto:<%=myContact_Home_Email%>"><%=myContact_Home_Email%></A></font>
</td>
</tr> 


<%
' Home Web
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Web%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><A HREF="<%=myContact_Home_Web%>"><%=myContact_Home_Web%></A></font>
</td>
</tr>

<%
' Separator
%>

<TR>
<TD VALIGN="top" ALIGN="right" COLSPAN="2" BGCOLOR="<%=myApplicationColor%>">&nbsp;

</TD>
</TR> 

<% end If ' Contact Domicile %> 


<%
' Comments
%>

<TR>
<TD ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Comments%> 
</FONT></B>
</TD>
<TD ALIGN="left">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myContact_Comments%></FONT>
</TD>
</tr>



<%
' Date and Author
%>

<TR>
<TD VALIGN="top" ALIGN="right" COLSPAN="2" BGCOLOR="<%=myApplicationColor%>"> 
<P ALIGN="center"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" Color="<%=myApplicationTextColor%>">
<% if len(myContact_Date_Update) > 0 then %>
 <% = myDate_Display(myContact_Date_Update,2) %> -- <% = myContact_Author_Update %>
<% end if %></FONT></P>
</TD>
</TR> 


</table>

<%
' NAVIGATION
%>


<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0"> 
<td>&nbsp;<a href="__Contact_Modification.asp?Action=New&Directory_ID=<%=myDirectory_ID%>"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Add%>&nbsp;<%=myMessage_Contact%></font></a>
</td>
</tr>
<tr> 
<td>&nbsp;<a href="__Contact_Modification.asp?Action=Update&Contact_Id=<%=myContact_ID%>"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Modify%></font></a> 
, <a href="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Contact_Modification.asp?Action=Delete&amp;Contact_ID=<%=myContact_ID%>';"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Delete%></font></a>
</td>
</tr>
 
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
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> OverApps</font></A> & contributors</FONT>
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