<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   _ OverApps - http://www.overapps.com
'
' This program "__Contact_Modification.asp" is free software; you can 
' redistribute it and/or modify
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
' Does n't Work with PWS ?
%>

<%
' ------------------------------------------------------------
' 
' Name			: __Contact_Modification.asp
' Path   		: /
' Version		: 1.15.0
' Description 	: Contact Modification
' By		 	: Pierre Rouarch	
' Company		: OverApps
' Date			:December,10, 2001
' Contributions : Jean-Luc Lesueur, Christophe Humbert, Dania Tcherkezoff
' 
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------

Dim myPage
myPage = "__Contact_Modification.asp"

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

Dim myDirectory_ID, myDirectory_Name, myContact_ID,  myContact_Company_Activity, myContact_Site_ID, myContact_Member_ID, myContact_Title_ID, myContact_Title,  myContact_Name, myContact_FirstName, myContact_Company_Type, myContact_Company, myContact_Company_Activity_ID, myContact_Company_Activity_ID_Selected, myContact_Company_Activity_Name, myContact_Company_Address, myContact_Company_Zip, myContact_Company_City, myContact_Company_State, myContact_Company_Country_ID, myContact_Company_Country_ID_Selected, myContact_Company_Country, myContact_Company_Phone, myContact_Company_Mobile, myContact_Company_Fax, myContact_Company_Email, myContact_Company_Web, myContact_Company_Fonction, myContact_Home_Type, myContact_Home_Address, myContact_Home_Zip, myContact_Home_City, myContact_Home_State, myContact_Home_Country_ID,   myContact_Home_Country_ID_Selected, myContact_Home_Country, myContact_Home_Phone, myContact_Home_Mobile, myContact_Home_Fax, myContact_Home_Email, myContact_Home_Web, myContact_Comments, myContact_Author_Update, myContact_Date_Update

Dim  myMethod_Search, mySearch_Contact_Company, mySearch_Contact_Company_Activity, mySearch_Contact_Company_Activity_ID, mySearch_Contact_Name, mySearch_Contact_City

Dim mySQL_Select_tb_Contacts, mySet_tb_Contacts, mySQL_Select_tb_Directories, mySet_tb_Directories, mySQL_Select_tb_Contacts_Activities,  mySet_tb_Contacts_Activities 


Dim mySQL_Select_tb_Countries, mySet_tb_Countries

Dim myMember_ID,  myAction

Dim mySQL_Delete_tb_Contacts, mySQL_Insert_tb_Contacts,  mySQL_Update_tb_Contacts

''''''''''''''''''''''''''''''''''''
' Initialization and Get Parameters
''''''''''''''''''''''''''''''''''''
' Get Parameters
myAction = Request("Action")
if len(myAction)=0 then
	myAction="New"
end if
myContact_ID = Request("Contact_ID") 
if len(myContact_Id)=0 then 
	myAction="New"
end if

' Force to Directory 1 in this version
' myDirectory_ID  = Request("Directory_ID")
myDirectory_ID=1

myPage=myPage&"?Action="&myAction&"&Contact_ID="&myContact_ID&"&Directory_ID="&myDirectory_ID


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Delete
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myAction = "Delete" then

	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String
	' Everybody Can delete a contact in this version
	mySQL_Delete_Tb_Contacts = "DELETE FROM tb_contacts WHERE Contact_ID="&myContact_ID&" AND Directory_ID="&myDirectory_ID&" AND Site_ID="&mySite_ID
	myConnection.Execute(mySQL_Delete_Tb_Contacts)
	myConnection.Close
	set myConnection = Nothing
	' And Go Back
	Response.Redirect("__Contacts_list.asp")

end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation
'''''''''''''''''''''''''''''''''''''''''''''''''


if Request.form("Validation")=myMessage_Go then

	' For multi-sites Purpose
	myContact_Site_ID  = Request.Form("Contact_Site_ID")
	' First Author
	myContact_Member_ID  = Request.Form("Contact_Member_ID")
	myContact_ID  = Request.Form("Contact_ID")
	myContact_Title_ID = Request.Form("Contact_Title_ID")
	myContact_FirstName = Replace(Request.Form("Contact_FirstName"),"'"," ")
	myContact_Name = Replace(Request.Form("Contact_Name"),"'"," ")

	myContact_Company_Type = Request.Form("Contact_Company_Type")
	if len(myContact_Company_Type)>0 then 
		myContact_Company_Type=1
	else
		myContact_Company_Type=0
	end if 
	myContact_Company_Fonction = Replace(Request.Form("Contact_Company_Fonction"),"'"," ")

	myContact_Company = Replace(Request.Form("Contact_Company"),"'"," ")

	myContact_Company_Activity_ID_Selected=0

	myContact_Company_Address = Replace(Request.Form("Contact_Company_Address"),"'"," ")
	myContact_Company_Zip = Replace(Request.Form("Contact_Company_Zip"),"'"," ")
	myContact_Company_City = Replace(Request.Form("Contact_Company_City"),"'"," ")
	myContact_Company_State = Replace(Request.Form("Contact_Company_State"),"'"," ")
	myContact_Company_Country_ID_Selected = Request.Form("Contact_Company_Country_ID")

	if len(myContact_Company_Country_ID_Selected  & " ")= 1 then 
		myContact_Company_Country_ID_Selected=0
	end if
	myContact_Company_Phone = Replace(Request.Form("Contact_Company_Phone"),"'"," ")
	myContact_Company_Mobile = Replace(Request.Form("Contact_Company_Mobile"),"'"," ")
	myContact_Company_Fax = Replace(Request.Form("Contact_Company_Fax"),"'"," ")
	myContact_Company_Email = Replace(Request.Form("Contact_Company_Email"),"'","")
	myContact_Company_Web = Replace(Request.Form("Contact_Company_Web"),"'"," ")
	myContact_Company_Fonction = Replace(Request.Form("Contact_Company_Fonction"),"'"," ")



	myContact_Home_Type = Request.Form("Contact_Home_Type")
	if len(myContact_Home_Type)>0 then 
		myContact_Home_Type=1
	else
		myContact_Home_Type=0
	end if 
	myContact_Home_Address = Replace(Request.Form("Contact_Home_Address"),"'"," ")
	myContact_Home_Zip = Replace(Request.Form("Contact_Home_Zip"),"'"," ")
	myContact_Home_City = Replace(Request.Form("Contact_Home_City"),"'"," ")
	myContact_Home_State = Replace(Request.Form("Contact_Home_State"),"'"," ")
	myContact_Home_Country_ID_Selected = Request.Form("Contact_Home_Country_ID")
	if len(myContact_Home_Country_ID_Selected & " ")= 1  then 
		myContact_Home_Country_ID_Selected=0
	end if
	myContact_Home_Phone = Replace(Request.Form("Contact_Home_Phone"),"'"," ")
	myContact_Home_Mobile = Replace(Request.Form("Contact_Home_Mobile"),"'"," ")
	myContact_Home_Fax = Replace(Request.Form("Contact_Home_Fax"),"'"," ")
	myContact_Home_Email = Replace(Request.Form("Contact_Home_Email"),"'"," ")
	myContact_Home_Web = Replace(Request.Form("Contact_Home_Web"),"'"," ")
	myContact_Comments = Replace(Request.Form("Contact_Comments"),"'"," ")



' Test Entries 
Call myFormSetEntriesInString

myFormCheckEntry null, "Contact_Name",true,null,null,0,100


if not myform_entry_error then

' db Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

myContact_Author_Update = myUser_Pseudo
myContact_Date_Update = myDate_Now()


if myAction = "New" then

myContact_Site_ID=mySite_ID
myContact_Member_ID = myUser_ID

' INSERT

	mySQL_Select_tb_Contacts = "SELECT * FROM tb_contacts"			
	Set mySet_tb_Contacts = server.createobject("adodb.recordset")
   	mySet_tb_Contacts.open mySQL_Select_tb_Contacts, myConnection, 3,3
   	mySet_tb_Contacts.AddNew

		mySet_tb_Contacts.fields("Site_ID") = myContact_Site_ID
		mySet_tb_Contacts.fields("Member_ID") = myContact_Member_ID
		mySet_tb_Contacts.fields("Directory_ID") = myDirectory_ID
		mySet_tb_Contacts.fields("Contact_Title_ID") = myContact_Title_ID
		mySet_tb_Contacts.fields("Contact_Name") = myContact_Name
		mySet_tb_Contacts.fields("Contact_FirstName") = myContact_Firstname
		mySet_tb_Contacts.fields("Contact_Company_Type") = myContact_Company_Type
		mySet_tb_Contacts.fields("Contact_Company_Fonction") = myContact_Company_Fonction
		mySet_tb_Contacts.fields("Contact_Company") = myContact_Company
		mySet_tb_Contacts.fields("Contact_Company_Activity_ID") = myContact_Company_Activity_ID_Selected
		mySet_tb_Contacts.fields("Contact_Company_Address") = myContact_Company_Address
		mySet_tb_Contacts.fields("Contact_Company_Zip") = myContact_Company_Zip
		mySet_tb_Contacts.fields("Contact_Company_City") = myContact_Company_City
		mySet_tb_Contacts.fields("Contact_Company_State") = myContact_Company_State
		mySet_tb_Contacts.fields("Contact_Company_Country_ID") = myContact_Company_Country_ID_Selected
		mySet_tb_Contacts.fields("Contact_Company_Phone") = myContact_Company_Phone
		mySet_tb_Contacts.fields("Contact_Company_Mobile") = myContact_Company_Mobile
		mySet_tb_Contacts.fields("Contact_Company_Fax") = myContact_Company_Fax
		mySet_tb_Contacts.fields("Contact_Company_Email") = myContact_Company_Email
		mySet_tb_Contacts.fields("Contact_Company_Web") = myContact_Company_Web
		mySet_tb_Contacts.fields("Contact_Home_Address") = myContact_Home_Address
		mySet_tb_Contacts.fields("Contact_Home_Zip") = myContact_Home_Zip
		mySet_tb_Contacts.fields("Contact_Home_City") = myContact_Home_City
		mySet_tb_Contacts.fields("Contact_Home_State") = myContact_Home_State
		mySet_tb_Contacts.fields("Contact_Home_Country_ID") = myContact_Home_Country_ID_Selected
		mySet_tb_Contacts.fields("Contact_Home_Phone") = myContact_Home_Phone
		mySet_tb_Contacts.fields("Contact_Home_Mobile") = myContact_Home_Mobile
		mySet_tb_Contacts.fields("Contact_Home_Fax") = myContact_Home_Fax
		mySet_tb_Contacts.fields("Contact_Home_Email") = myContact_Home_Email
		mySet_tb_Contacts.fields("Contact_Home_Web") = myContact_Home_Web
		mySet_tb_Contacts.fields("Contact_Home_Type") = myContact_Home_Type
		mySet_tb_Contacts.fields("Contact_Comments") = myContact_Comments
		mySet_tb_Contacts.fields("Contact_Author_Update") = myContact_Author_Update
		mySet_tb_Contacts.fields("Contact_Date_Update") = myContact_Date_Update
	
	mySet_tb_Contacts.Update
	mySet_tb_Contacts.close
	Set mySet_tb_Contacts = Nothing

	ElseIf myAction = "Update" Then

	'UPDATE
	mySQL_Select_tb_Contacts = "SELECT * FROM tb_contacts WHERE  Contact_ID =" & myContact_ID		
	Set mySet_tb_Contacts = server.createobject("adodb.recordset")
   	mySet_tb_Contacts.open mySQL_Select_tb_Contacts, myConnection, 3,3

		mySet_tb_Contacts.fields("Contact_Title_ID") = myContact_Title_ID
		mySet_tb_Contacts.fields("Contact_Name") = myContact_Name
		mySet_tb_Contacts.fields("Contact_FirstName") = myContact_FirstName
		mySet_tb_Contacts.fields("Contact_Company_Type") = myContact_Company_Type
		mySet_tb_Contacts.fields("Contact_Company_Fonction") = myContact_Company_Fonction 
		mySet_tb_Contacts.fields("Contact_Company") = myContact_Company
		mySet_tb_Contacts.fields("Contact_Company_Activity_ID") = myContact_Company_Activity_ID_Selected 
		mySet_tb_Contacts.fields("Contact_Company_Address") = myContact_Company_Address
		mySet_tb_Contacts.fields("Contact_Company_Zip") = myContact_Company_zip
		mySet_tb_Contacts.fields("Contact_Company_City") = myContact_Company_City 
		mySet_tb_Contacts.fields("Contact_Company_State") = myContact_Company_State 
		mySet_tb_Contacts.fields("Contact_Company_Country_ID") = myContact_Company_Country_ID_Selected 
		mySet_tb_Contacts.fields("Contact_Company_Phone") = myContact_Company_Phone
		mySet_tb_Contacts.fields("Contact_Company_Mobile") = myContact_Company_Mobile 
		mySet_tb_Contacts.fields("Contact_Company_Fax") = myContact_Company_Fax
		mySet_tb_Contacts.fields("Contact_Company_Email") = myContact_Company_Email 
		mySet_tb_Contacts.fields("Contact_Company_Web") = myContact_Company_Web 
		mySet_tb_Contacts.fields("Contact_Home_Type") = myContact_Home_Type
		mySet_tb_Contacts.fields("Contact_Home_Address") = myContact_Home_Address 
		mySet_tb_Contacts.fields("Contact_Home_Zip") = myContact_Home_zip
		mySet_tb_Contacts.fields("Contact_Home_City") = myContact_Home_City 
		mySet_tb_Contacts.fields("Contact_Home_State") = myContact_Home_State 
		mySet_tb_Contacts.fields("Contact_Home_Country_ID") = myContact_Home_Country_ID_Selected 
		mySet_tb_Contacts.fields("Contact_Home_Phone") = myContact_Home_Phone
		mySet_tb_Contacts.fields("Contact_Home_Mobile") = myContact_Home_Mobile 
		mySet_tb_Contacts.fields("Contact_Home_Fax") = myContact_Home_Fax
		mySet_tb_Contacts.fields("Contact_Home_Email") = myContact_Home_Email 
		mySet_tb_Contacts.fields("Contact_Home_Web") = myContact_Home_Web 
		mySet_tb_Contacts.fields("Contact_Comments") = myContact_Comments 
		mySet_tb_Contacts.fields("Contact_Author_Update") = myContact_Author_Update 
		mySet_tb_Contacts.fields("Contact_Date_Update") = myContact_Date_Update

		mySet_tb_Contacts.Update
		' Close Recordset	
		mySet_tb_Contacts.close
		Set mySet_tb_Contacts = Nothing

		end if


	' Close Connection 
	myConnection.close
	set myConnection = nothing	
	
	Response.Redirect("__Contacts_List.asp")
		

end if 

end if 


%>
<html>

<head>
<title><%=mySite_Name%> - Add/Modification - Contact</title>

</head>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">

<%
' TOP
%>

<!-- #include file="_borders/Top.asp" --> 

<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>" bgColor="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> 

<%
' CENTER LEFT
%>


 <TD WIDTH="<%=myLeft_Width%>"> <!-- #include file="_borders/Left.asp" --></td>

<%
' CENTER APPLICATION
%>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form 					'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myAction="Update" and Request.form("Validation")<>myMessage_Go then


	' db connection
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String


	mySQL_Select_tb_Contacts = "SELECT * FROM tb_contacts WHERE Contact_ID = "&myContact_ID
	set mySet_tb_Contacts = myConnection.Execute(mySQL_Select_tb_Contacts)

	if not mySet_tb_Contacts.eof then
		' read values
		myContact_Site_ID = mySet_tb_Contacts("Site_ID")
		myContact_Member_ID = mySet_tb_Contacts("Member_ID")
		myDirectory_ID = mySet_tb_Contacts("Directory_ID")
		myContact_Title_ID = mySet_tb_Contacts("Contact_Title_ID")
		myContact_Name = mySet_tb_Contacts("Contact_Name")
		myContact_FirstName = mySet_tb_Contacts("Contact_FirstName")
		myContact_Company_Type = mySet_tb_Contacts("Contact_Company_Type")
		if myContact_Company_Type=0 then
			 myContact_Company_type=null 
		end if
		myContact_Company = mySet_tb_Contacts("Contact_Company")
		myContact_Company_Activity_ID_selected = mySet_tb_Contacts("Contact_Company_Activity_ID")
		myContact_Company_Address = mySet_tb_Contacts("Contact_Company_Address")
		myContact_Company_Zip = mySet_tb_Contacts("Contact_Company_Zip")
		myContact_Company_City = mySet_tb_Contacts("Contact_Company_City")
		myContact_Company_State = mySet_tb_Contacts("Contact_Company_State")
		myContact_Company_Country_ID_Selected = mySet_tb_Contacts("Contact_Company_Country_ID")
		myContact_Company_Phone = mySet_tb_Contacts("Contact_Company_Phone")
		myContact_Company_Mobile = mySet_tb_Contacts("Contact_Company_Mobile")
		myContact_Company_Fax = mySet_tb_Contacts("Contact_Company_Fax")
		myContact_Company_Email = mySet_tb_Contacts("Contact_Company_Email")
		myContact_Company_Web = mySet_tb_Contacts("Contact_Company_Web")
		myContact_Company_Fonction = mySet_tb_Contacts("Contact_Company_Fonction")
		myContact_Home_Type = mySet_tb_Contacts("Contact_Home_Type")
		if myContact_Home_Type=0 then
			 myContact_Home_type=null 
		end if
		myContact_Home_Address = mySet_tb_Contacts("Contact_Home_Address")
		myContact_Home_Zip = mySet_tb_Contacts("Contact_Home_Zip")
		myContact_Home_City = mySet_tb_Contacts("Contact_Home_City")
		myContact_Home_State = mySet_tb_Contacts("Contact_Home_State")
		myContact_Home_Country_ID_Selected = mySet_tb_Contacts("Contact_Home_Country_ID")
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

	' Close Recordset
	mySet_tb_Contacts.close
	Set mySet_tb_Contacts=Nothing
	' Close Connection
	myConnection.close
	set myConnection = nothing



end if '/ UPDATE

%> 


<td valign="top"  bgcolor="<%=mybgColor%>" Width="<%=myApplication_Width%>">

<form method="POST" action="<%=myPage%>" name="myForm" > 

<table border="0" bgcolor="<%=mybgColor%>" cellpadding="5" cellspacing="1" Width="<%=myApplication_Width%>"> 

<%
' Application Title AND Hidden Fields
%>

<tr>
<td  align=Center colspan="2" bgcolor="<%=myApplicationColor%>">
<font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="4"><b><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR="<%=myApplicationTextColor%>"><%=myApplication_Title%></FONT></b></font>
<INPUT TYPE="hidden" NAME="Directory_ID" VALUE="<%=myDirectory_ID%>">
<INPUT TYPE="hidden" NAME="Contact_Site_ID" VALUE="<%=myContact_Site_ID%>"> 
<INPUT TYPE="hidden" NAME="Contact_Member_ID" VALUE="<%=myContact_Member_ID%>"> 
<INPUT TYPE="hidden" NAME="Contact_ID" VALUE="<%=myContact_ID%>">
<INPUT TYPE="hidden" NAME="Action" VALUE="<%=myAction%>"> 
</td>
</tr>

<%
' Title _ Civility
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Title%><BR></FONT></B>
</td>
<td align="left" valign="top"> 
<SELECT NAME="Contact_Title_ID"> 
<OPTION VALUE="0" <%if myContact_Title_Id=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myContact_Title_Id=1 then%>Selected<%end if%>><%=myMessage_Mister%></OPTION> 
<OPTION VALUE="2" <%if myContact_Title_Id=2 then%>Selected<%end if%>><%=myMessage_Misses%></OPTION> 
<OPTION VALUE="3" <%if myContact_Title_Id=3 then%>Selected<%end if%>><%=myMessage_Miss%></OPTION></SELECT> 
</td>
</tr>

<%
' FirstName
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_FirstName%></FONT></B>
</td>
<td align="left" valign="top">
<INPUT TYPE="text" NAME="Contact_FirstName" Value="<%=myContact_FirstName%>"> 
</td>
</tr> 


<%
' Name
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Name%>*<br><%=myFormGetErrMsg("Contact_Name")%></FONT></B>
</td>
<td align="left" valign="top">
<INPUT TYPE="text" NAME="Contact_Name" Value="<%=myContact_Name%>">
</td>
</tr> 

<%
' Company Type
%>

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"> 
<%=myMessage_Office%> : <INPUT TYPE="checkbox" NAME="Contact_Company_Type" VALUE="checkbox" <%if len(myContact_Company_Type)>0 then %> CHECKED <% end if %>></font></b>
</td>
</tr> 


<%
' Contact Fonction (in the Company)
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Fonction%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Contact_Company_Fonction" Value="<%=myContact_Company_Fonction%>">
</td>
</tr> 


<%
' Company Name
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Company%>
</font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Contact_Company"  Value="<%=myContact_Company%>">
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
<INPUT TYPE="text" NAME="Contact_Company_Address" Value="<%=myContact_Company_Address%>">
</td>
</tr>

<%
' Zip Code
%>
 
<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Zip_Code%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Contact_Company_Zip" Value="<%=myContact_Company_Zip%>">
</td>
</tr> 

<%
' Company City
%>


<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_City%></font></b>
</td>
<td align="left"><INPUT TYPE="text" NAME="Contact_Company_City" Value="<%=myContact_Company_City%>">
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
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Contact_Company_State" Value="<%=myContact_Company_State%>"></font>
</td>
</tr> 

<%
' Company Country
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Country%></font></b>
</td>
<td align="left"> 
<% 
'''''''''''''''''''''''''''''''''
' Company Country '	
''''''''''''''''''''''''''''''''

' db Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Countries = "SELECT * FROM tb_Countries order by Country"
set mySet_Tb_Countries = myConnection.Execute(mySQL_Select_tb_Countries)
 %> 
<P><select name="Contact_Company_Country_ID" size="" tabindex="1"> 
<option value="<%=myContact_Home_Country_ID%>" <%if myContact_Home_Country_ID_selected = 0  then%> SELECTED <%end if %> > <%=myMessage_Select%></Option>

<%
do while not mySet_Tb_Countries.eof
	myContact_Company_Country_ID = mySet_Tb_Countries("Country_ID")
	myContact_Company_Country = mySet_Tb_Countries("Country")
	%> 

	<option value="<%=myContact_Company_Country_ID%>"  <%if myContact_Company_Country_ID_selected = myContact_Company_Country_ID then%> SELECTED <%end if %> ><%=myContact_Company_Country%></option> 

	<%
	mySet_Tb_Countries.MoveNext
loop
%> 

</select></P>

<%
' Close Recordset
mySet_Tb_Countries.close
Set mySet_Tb_Countries = Nothing
' Close Connection
myConnection.Close
set myConnection = Nothing
%>
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
<INPUT TYPE="text" NAME="Contact_Company_Phone" Value="<%=myContact_Company_Phone%>">
</td>
</tr>

<%
' Mobile
%>
 
<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Mobile%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Contact_Company_Mobile" Value="<%=myContact_Company_Mobile%>">
</td>
</tr> 

<%
' Fax
%>
 
<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Fax%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Contact_Company_Fax" Value="<%=myContact_Company_Fax%>">
</td>
</tr> 

<%
' Email
%>

<tr><td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Email%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Contact_Company_Email" Value="<%=myContact_Company_Email%>"></font>
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
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Contact_Company_Web" Value="<%=myContact_Company_Web%>"></font>
</td>
</tr> 


<%
' HOME
%>

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"><%=myMessage_Home%> : <INPUT TYPE="checkbox" NAME="Contact_Home_Type" VALUE="checkbox" <%if len(myContact_Home_Type)>0 then %> CHECKED <% end if %>></font></b>
</td>
</tr> 

<%
' Home Address
%>

<tr><td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Address%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Contact_Home_Address" Value="<%=myContact_Home_Address%>">
</td>
</tr> 


<%
' Zip Code
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Zip_Code%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Contact_Home_Zip" Value="<%=myContact_Home_Zip%>">
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
<INPUT TYPE="text" NAME="Contact_Home_City" Value="<%=myContact_Home_City%>">
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
<INPUT TYPE="text" NAME="Contact_Home_State" Value="<%=myContact_Home_State%>">
</td>
</tr> 

<%
' Home Country
%>


<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Country%></font></b>
</td>
<td align="left"> 
<% 
''''''''''''''''''''''''''''''''''
' Home Country					 '	
''''''''''''''''''''''''''''''''''

' DataBase Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Countries = "SELECT * FROM tb_Countries order by Country"

set mySet_Tb_Countries = myConnection.Execute(mySQL_Select_tb_Countries) 
%> 

<P><select name="Contact_Home_Country_ID" size="" tabindex="1"> 

<option value="<%=myContact_Home_Country_ID%>" <%if myContact_Home_Country_ID_selected = 0  then%> SELECTED <%end if %> > <%=myMessage_Select%></Option>

<%
do while not mySet_Tb_Countries.eof
	myContact_Home_Country_ID = mySet_Tb_Countries("Country_ID")
	myContact_Home_Country = mySet_Tb_Countries("Country")

	%> 
	<option value="<%=myContact_Home_Country_ID%>" <%if myContact_Home_Country_ID_selected =	 myContact_Home_Country_ID then%> SELECTED <%end if %> ><%=myContact_Home_Country%></option> <%
	mySet_Tb_Countries.MoveNext
loop
%> 
</select></P>
<%
' Close 
mySet_Tb_Countries.close
Set mySet_Tb_Countries = Nothing
myConnection.Close
set myConnection = Nothing
%>
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
<INPUT TYPE="text" NAME="Contact_Home_Phone" Value="<%=myContact_Home_Phone%>">
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
<INPUT TYPE="text" NAME="Contact_Home_Mobile" Value="<%=myContact_Home_Mobile%>">
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
<INPUT TYPE="text" NAME="Contact_Home_Fax" Value="<%=myContact_Home_Fax%>">
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
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Contact_Home_Email" Value="<%=myContact_Home_Email%>"></font>
</td>
</tr> 


<%
' home Web
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Web%></font></b>
</td>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Contact_Home_Web" Value="<%=myContact_Home_Web%>"></font>
</td>
</tr>

<%
' Separator
%>

<TR>
<TD VALIGN="top" ALIGN="right" colspan="2" BGCOLOR="<%=myApplicationColor%>">&nbsp;

</TD>
</TR>

<%
' Comments
%>

<tr>
<TD VALIGN="top" ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Comments%></FONT></B>
</TD>
<TD VALIGN="top" ALIGN="left">
<textarea name="Contact_Comments" cols="50" rows="5"><%=myContact_Comments%></textarea>
</TD>
<TR> 

<%
' Validation 
%>


<TD ALIGN="right" VALIGN="top" BGCOLOR="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" Color="<%=myBorderTextColor%>">* = <%=myMessage_Required%></FONT></B>
</TD>
<TD VALIGN="top" ALIGN="left"> 
<INPUT TYPE="submit" VALUE="<%=myMessage_Go%>" NAME="Validation">
</TD>
</TR>

<%
' Author - date
%>


<TR>
<TD VALIGN="top" ALIGN="right" COLSPAN="2" BGCOLOR="<%=myApplicationColor%>"> 
<P ALIGN="center"><B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" Color="<%=myApplicationTextColor%>"><%=myDate_Display(myContact_Date_Update,2)%> -- <%=myContact_Author_Update%></FONT></B></P>
</TD>
</TR>
</table>

</form>

<% 
' ADMINISTRATION
' Everybody can Add, Modify and Delete Contacts in this Version
%>
<% If myAction="Update" then %>
<TABLE BORDER="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="3" CELLSPACING="0"> 
<tr>
<td>
&nbsp;
<a href="__Contact_Modification.asp?Action=New&Directory_ID=<%=myDirectory_ID%>"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Add%>&nbsp;<%=myMessage_Contact%></font></a>
,  
<A HREF="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Contact_Modification.asp?Action=Delete&amp;Contact_ID=<%=myContact_ID%>';"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Delete%></font></A>
</TD>
</TR> 
</TABLE>
<% End If %>

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
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors
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
<html></html>