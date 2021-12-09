<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApp s - http://www.overapps.com
'
' This program "__Site_Member_Modification.asp" is free software; you can 
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
' Cache non géré par PWS
%>

<%
' ----------------------------------------------------------------------------------
' OverApps official Source : 
' Name			: __Site_Member_Modification.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Add/Modify/delete a Site Member
' By 			: Pierre Rouarch
' Company 		: OverApps
' Update		: November,21, 2001
' Contributions : Jean-Luc Lesueur, Christophe Humbert, Nicolas Sanchez
'
' Last Modifications :  
'
' 
'------------------------------------------------------------------------------------
' If you have modified the original OverApps official Source please Fill :
' 
' Modify by		:
' Company		:
' Date			:
'
' 
' -----------------------------------------------------------------------------------

Dim myPage
myPage = "__Site_Member_Modification.asp"

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


' DB 
Dim mySQL_Delete_tb_Sites_Members, mySQL_Insert_tb_Sites_Members, mySQL_Update_tb_Sites_Members

Dim myMember_Login, myMember_Password, myMember_Password_Confirmation, myMember_Site_ID,  myMember_ID,  myMember_Company_Activity,  myMember_Member_ID, myMember_Title_ID, myMember_Title,  myMember_Name, myMember_FirstName, myMember_Pseudo, myMember_Email,  myMember_Company_Type, myMember_Company_Fonction, myMember_Company, myMember_Company_Activity_ID, myMember_Company_Activity_ID_Selected, myMember_Company_Activity_Name, myMember_Company_Address, myMember_Company_Zip, myMember_Company_City, myMember_Company_State, myMember_Company_Country_ID, myMember_Company_Country_ID_Selected, myMember_Company_Country, myMember_Company_Phone, myMember_Company_Mobile, myMember_Company_Fax, myMember_Company_Email, myMember_Company_Web,  myMember_Home_Type, myMember_Home_Address, myMember_Home_Zip, myMember_Home_City, myMember_Home_State, myMember_Home_Country_ID,   myMember_Home_Country_ID_Selected, myMember_Home_Country, myMember_Home_Phone, myMember_Home_Mobile, myMember_Home_Fax, myMember_Home_Email, myMember_Home_Web, myMember_Comments, myMember_type_ID, myMember_Author_Update, myMember_Date_Update

Dim  myMethod_Search, mySearch_Member_Company, mySearch_Member_Company_Activity, mySearch_Member_Company_Activity_ID, mySearch_Member_Name, mySearch_Member_City

Dim   mySQL_Select_tb_Sites_Members_Activities, mySet_tb_Sites_Members_Activities, mySQL_Select_tb_Countries, mySet_tb_Countries


Dim   myAction, myTitle, myBoutton,   myDateMaj, myAuteurmaj, myControl_Value, myControl_Name
					
Dim myError

%>

<%

myError=0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Parameters
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myAction = Request("Action")
myMember_ID = Request("Member_ID")
if len(myMember_ID)=0 or myMember_ID=0 then
	myMember_ID=0
	myAction="New"
end if
myMember_ID=Cint(myMember_ID)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Verify Autorization to go further
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' only Administrators

if myAction="New" or myAction="Delete" then
	if  (myUser_Type_ID<>1) then 
		response.redirect("__Sites_Members_List.asp")
	end if
end if


' Only me and Administrators
if myAction="Update" then
	if (myMember_ID<>myUser_ID) then 
		if (myUser_Type_ID<>1) then
			response.redirect("__Sites_Members_List.asp")
		end if
	end if
end if 


myPage="__Site_Member_Modification.asp"
myPage=myPage & "?Action="&myAction&"&amp;Member_ID="&myMember_ID&"&amp;Site_ID="&mySite_ID


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DELETE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction = "Delete" then

	' CAN'T DELETE AN ADMINISTRATOR YOU MUST CHANGE THE ADMINISTRATOR IN A SAMPLE USER
	if myMember_Type_ID=1 then 	
		Response.Redirect("__Sites_Members_List.asp") 
	end if

	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

	mySQL_Delete_Tb_Sites_Members = "DELETE FROM tb_Sites_Members WHERE Member_ID = " & myMember_ID

	myConnection.Execute(mySQL_Delete_Tb_Sites_Members)
	myConnection.Close
	set myConnection = Nothing
	Response.Redirect("__Sites_Members_List.asp")

end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.form("Validation")=myMessage_Go then

	' Get Fields
	myMember_Site_ID=mySite_ID
	myMember_Login = Replace(Request.Form("Member_Login"),"'"," ")
	myMember_Password = Replace(Request.Form("Member_Password"),"'"," ")
	myMember_Password_Confirmation = Replace(Request.Form("Member_Password_Confirmation"),"'"," ")
	myMember_Title_ID = Request.Form("Member_Title_ID")
	myMember_FirstName = Replace(Request.Form("Member_FirstName"),"'"," ")
	myMember_Name = Replace(Request.Form("Member_Name"),"'"," ")
	myMember_Pseudo=Replace(myMember_Login,"'"," ")
	myMember_Email = Replace(Request.Form("Member_Email"),"'"," ")
	myMember_Company_Type = Request.Form("Member_Company_Type")
	if len(myMember_Company_Type)>0 then 
		myMember_Company_Type=true
	else
		myMember_Company_Type=False
	end if 
	myMember_Company_Fonction = Replace(Request.Form("Member_Company_Fonction"),"'"," ")
	myMember_Company = Replace(Request.Form("Member_Company"),"'"," ")
	myMember_Company_Activity_ID_Selected=0
	myMember_Company_Address = Replace(Request.Form("Member_Company_Address"),"'"," ")
	myMember_Company_Zip = Replace(Request.Form("Member_Company_Zip"),"'"," ")
	myMember_Company_City = Replace(Request.Form("Member_Company_City"),"'"," ")
	myMember_Company_State = Replace(Request.Form("Member_Company_State"),"'"," ")
	myMember_Company_Country_ID_Selected =	Request.Form("Member_Company_Country_ID_Selected")
	myMember_Company_Country_ID_Selected=CInt(myMember_Company_Country_ID_Selected)
	myMember_Company_Phone = Replace(Request.Form("Member_Company_Phone"),"'"," ")
	myMember_Company_Mobile = Replace(Request.Form("Member_Company_Mobile"),"'"," ")
	myMember_Company_Fax = Replace(Request.Form("Member_Company_Fax"),"'"," ")
	myMember_Company_Email = Replace(Request.Form("Member_Company_Email"),"'","")
	myMember_Company_Web = Replace(Request.Form("Member_Company_Web"),"'"," ")
	myMember_Home_Type = Request.Form("Member_Home_Type")
	if len(myMember_Home_Type)>0 then 
		myMember_Home_Type=true
	else
		myMember_Home_Type=False
	end if 
	myMember_Home_Address = Replace(Request.Form("Member_Home_Address"),"'"," ")
	myMember_Home_Zip = Replace(Request.Form("Member_Home_Zip"),"'"," ")
	myMember_Home_City = Replace(Request.Form("Member_Home_City"),"'"," ")
	myMember_Home_State = Replace(Request.Form("Member_Home_State"),"'"," ")
	myMember_Home_Country_ID_Selected = Request.Form("Member_Home_Country_ID_Selected")
	myMember_Home_Country_ID_Selected=CInt(myMember_Home_Country_ID_Selected)
	myMember_Home_Phone = Replace(Request.Form("Member_Home_Phone"),"'"," ")
	myMember_Home_Mobile = Replace(Request.Form("Member_Home_Mobile"),"'"," ")
	myMember_Home_Fax = Replace(Request.Form("Member_Home_Fax"),"'"," ")
	myMember_Home_Email = Replace(Request.Form("Member_Home_Email"),"'"," ")
	myMember_Home_Web = Replace(Request.Form("Member_Home_Web"),"'"," ")
	myMember_Comments = Replace(Request.Form("Member_Comments"),"'"," ")
	myMember_Type_ID = Request.Form("Member_Type_ID")
	myMember_type_ID=Cint(myMember_Type_ID)


	' Test form Validation
	Call myFormSetEntriesInString
	' Test Values :
	if myAction="New" then 
		myFormCheckEntry null, "Member_Login",true,null,null,0,100
 	end if
	myFormCheckEntry null, "Member_Password",true,null,null,0,100
	myFormCheckEntry null, "Member_Password_Confirmation",true,null,null,0,100
	myFormCheckEntry null, "Member_Email",true,null,null,0,100

	if not myform_entry_error then


		'''''''''''''''''''''''''''''''''''''''''''''''''
		' Test LOGIN ONLY FOR NEW
		'''''''''''''''''''''''''''''''''''''''''''''''''
		if myAction = "New" then
			' Check existing Login if yes myError=1
			' DB Connection
			set myConnection = Server.CreateObject("ADODB.Connection")
			myConnection.Open myConnection_String
			' Check Login
			mySQL_Select_tb_sites_members = "SELECT * FROM tb_Sites_Members WHERE Member_Login = '"&myMember_Login&"' "
			myConnection.Execute(mySQL_select_tb_sites_members)
			set mySet_tb_sites_members = myConnection.Execute(mySQL_select_tb_sites_members)
			' if it's OK
			if  not mySet_tb_sites_members.eof then
				myError=1
			end if
			' Close Recordset 
			mySet_tb_sites_members.close
			Set mySet_tb_sites_members = Nothing
		end if  '/NEW

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Check if Member_Password<>Member_Password_Confirmation if Yes myError=2
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		if myError=0 then 
			' Response.Write "Test Password"
			if myMember_Password<>myMember_Password_Confirmation then 
				myError=2
			end if 
		end if 


		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' IF IT'S ALL OK INSERT OR UPDATE
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		if myError=0 then 

			myMember_Author_Update=myUser_Pseudo
			myMember_Date_Update=myDate_Now()
			' DB Connection
			set myConnection = Server.CreateObject("ADODB.Connection")
			myConnection.Open myConnection_String

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' INSERT
			''''''''''''''''''''''''''''''''''''''''''''''''''''''

			if myAction="New" then 

				mySQL_Select_tb_Sites_Members = "SELECT * FROM tb_Sites_Members"
				Set mySet_tb_Sites_Members = server.createobject("adodb.recordset")
				mySet_tb_Sites_Members.open mySQL_Select_tb_Sites_Members, myConnection, 3, 3
				mySet_tb_Sites_Members.AddNew
	            mySet_tb_Sites_Members.fields("Site_ID") = myMember_Site_ID
				mySet_tb_Sites_Members.fields("Member_Login") = myMember_Login
				mySet_tb_Sites_Members.fields("Member_Password") = myMember_Password
				mySet_tb_Sites_Members.fields("Member_Title_ID") = myMember_Title_ID
				mySet_tb_Sites_Members.fields("Member_Name") = myMember_Name
				mySet_tb_Sites_Members.fields("Member_FirstName") = myMember_Firstname
				mySet_tb_Sites_Members.fields("Member_Pseudo") = myMember_Pseudo
				mySet_tb_Sites_Members.fields("Member_Email") = myMember_Email
				mySet_tb_Sites_Members.fields("Member_Company_Type") = myMember_Company_Type
				mySet_tb_Sites_Members.fields("Member_Company") = myMember_Company
				mySet_tb_Sites_Members.fields("Member_Company_Activity_ID") = myMember_Company_Activity_ID_Selected
				mySet_tb_Sites_Members.fields("Member_Company_Address") = myMember_Company_Address
				mySet_tb_Sites_Members.fields("Member_Company_Zip") = myMember_Company_Zip
				mySet_tb_Sites_Members.fields("Member_Company_City") = myMember_Company_City
				mySet_tb_Sites_Members.fields("Member_Company_State") = myMember_Company_State
				mySet_tb_Sites_Members.fields("Member_Company_Country_ID") = myMember_Company_Country_ID_Selected
				mySet_tb_Sites_Members.fields("Member_Company_Phone") = myMember_Company_Phone
				mySet_tb_Sites_Members.fields("Member_Company_Mobile") = myMember_Company_Mobile
				mySet_tb_Sites_Members.fields("Member_Company_Fax") = myMember_Company_Fax
				mySet_tb_Sites_Members.fields("Member_Company_Email") = myMember_Company_Email
				mySet_tb_Sites_Members.fields("Member_Company_Web") = myMember_Company_Web
				mySet_tb_Sites_Members.fields("Member_Home_Type") = myMember_Home_Type 
				mySet_tb_Sites_Members.fields("Member_Home_Address") = myMember_Home_Address
				mySet_tb_Sites_Members.fields("Member_Home_Zip") = myMember_Home_Zip
				mySet_tb_Sites_Members.fields("Member_Home_City") = myMember_Home_City
				mySet_tb_Sites_Members.fields("Member_Home_State") = myMember_Home_State
				mySet_tb_Sites_Members.fields("Member_Home_Country_ID") = myMember_Home_Country_ID_Selected
				mySet_tb_Sites_Members.fields("Member_Home_Phone") = myMember_Home_Phone
				mySet_tb_Sites_Members.fields("Member_Home_Mobile") = myMember_Home_Mobile
				mySet_tb_Sites_Members.fields("Member_Home_Fax") = myMember_Home_Fax
				mySet_tb_Sites_Members.fields("Member_Home_Email") = myMember_Home_Email
				mySet_tb_Sites_Members.fields("Member_Home_Web") = myMember_Home_Web
				mySet_tb_Sites_Members.fields("Member_Comments") = myMember_Comments
				mySet_tb_Sites_Members.fields("Member_type_ID") = myMember_Type_ID
				mySet_tb_Sites_Members.fields("Member_Author_Update") = myMember_Author_Update
				mySet_tb_Sites_Members.fields("Member_Date_Update") =  myMember_Date_Update

				mySet_tb_Sites_Members.Update

				' Close Recordset 
		  		mySet_tb_Sites_Members.close
		  		Set mySet_tb_Sites_Members = Nothing

			end if ' NEW 

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' UPDATE
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			if myAction = "Update" then

				' Update in DB
				mySQL_Select_tb_Sites_Members  = "SELECT * FROM tb_Sites_Members WHERE Member_ID =" & myMember_ID		
				Set mySet_tb_Sites_Members = server.createobject("adodb.recordset")
   				mySet_tb_Sites_Members.open mySQL_Select_tb_Sites_Members, myConnection, 3,3
				mySet_tb_Sites_Members.fields("Member_Password") = myMember_Password
			   	mySet_tb_Sites_Members.fields("Member_Title_ID") = myMember_Title_ID 
				mySet_tb_Sites_Members.fields("Member_Name") = myMember_Name
				mySet_tb_Sites_Members.fields("Member_FirstName") = myMember_FirstName 
				mySet_tb_Sites_Members.fields("Member_Pseudo") = myMember_Pseudo 
				mySet_tb_Sites_Members.fields("Member_Email") = myMember_Email 
				mySet_tb_Sites_Members.fields("Member_Company_Type") = myMember_Company_Type 
				mySet_tb_Sites_Members.fields("Member_Company_Fonction") = myMember_Company_Fonction 
				mySet_tb_Sites_Members.fields("Member_Company") = myMember_Company 
				mySet_tb_Sites_Members.fields("Member_Company_Activity_ID") = myMember_Company_Activity_ID_Selected 
				mySet_tb_Sites_Members.fields("Member_Company_Address") = myMember_Company_Address 
				mySet_tb_Sites_Members.fields("Member_Company_Zip") = myMember_Company_zip 
				mySet_tb_Sites_Members.fields("Member_Company_City") = myMember_Company_City 
				mySet_tb_Sites_Members.fields("Member_Company_State") = myMember_Company_State 
				mySet_tb_Sites_Members.fields("Member_Company_Country_ID") = myMember_Company_Country_ID_Selected 
				mySet_tb_Sites_Members.fields("Member_Company_Phone") = myMember_Company_Phone 
				mySet_tb_Sites_Members.fields("Member_Company_Mobile") = myMember_Company_Mobile 
				mySet_tb_Sites_Members.fields("Member_Company_Fax") = myMember_Company_Fax 
				mySet_tb_Sites_Members.fields("Member_Company_Email") = myMember_Company_Email 
				mySet_tb_Sites_Members.fields("Member_Company_Web") = myMember_Company_Web 
				mySet_tb_Sites_Members.fields("Member_Home_Type") = myMember_Home_Type 
				mySet_tb_Sites_Members.fields("Member_Home_Address") = myMember_Home_Address 
				mySet_tb_Sites_Members.fields("Member_Home_Zip") = myMember_Home_zip 
				mySet_tb_Sites_Members.fields("Member_Home_City") = myMember_Home_City 
				mySet_tb_Sites_Members.fields("Member_Home_State") = myMember_Home_State 
				mySet_tb_Sites_Members.fields("Member_Home_Country_ID") = myMember_Home_Country_ID_Selected 
				mySet_tb_Sites_Members.fields("Member_Home_Phone") = myMember_Home_Phone 
				mySet_tb_Sites_Members.fields("Member_Home_Mobile") = myMember_Home_Mobile 
				mySet_tb_Sites_Members.fields("Member_Home_Fax") = myMember_Home_Fax 
				mySet_tb_Sites_Members.fields("Member_Home_Email") = myMember_Home_Email 
				mySet_tb_Sites_Members.fields("Member_Home_Web") = myMember_Home_Web 
				mySet_tb_Sites_Members.fields("Member_Comments") = myMember_Comments 
				mySet_tb_Sites_Members.fields("Member_Type_ID") = myMember_Type_ID 
				mySet_tb_Sites_Members.fields("Member_Author_Update") = myMember_Author_Update 
				mySet_tb_Sites_Members.fields("Member_Date_Update") = myMember_Date_Update

				mySet_tb_Sites_Members.Update
				' Close Recordset 
		  		mySet_tb_Sites_Members.close
		  		Set mySet_tb_Sites_Members = Nothing
			

			end if '/Update

			' Close Connection
			myConnection.close
			set myConnection = nothing	
			' And Go Back
			Response.Redirect("__Sites_Members_List.asp")
	
		end if  'myError=0

	end if '/not my Form Entry Error

end if '/Validation


%>
<html>

<head>
<title><%=mySite_Name%> - <%=myMessage_Add%>/<%=myMessage_Modify%> <%=myMessage_Member%></title>

</head>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
'TOP
%> <!-- #include file="_borders/Top.asp" -->

<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form															'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Only if Update and not already validate
if myAction="Update" and Request.form("Validation")<>myMessage_Go then

	' DB Connection
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String


	mySQL_Select_tb_Sites_Members = "SELECT * FROM tb_Sites_Members WHERE Member_ID = "&myMember_ID
	set mySet_tb_Sites_Members = myConnection.Execute(mySQL_Select_tb_Sites_Members)

	if not mySet_tb_Sites_Members.eof then

	' Read in DB
	myMember_Site_ID = mySet_tb_Sites_Members("Site_ID")
	myMember_Member_ID = mySet_tb_Sites_Members("Member_ID")
	myMember_ID = mySet_tb_Sites_Members("Member_ID")
	myMember_Login=mySet_tb_Sites_Members("Member_Login")
	myMember_Password=mySet_tb_Sites_Members("Member_Password")
	myMember_Password_Confirmation=mySet_tb_Sites_Members("Member_Password")
	myMember_Title_ID = mySet_tb_Sites_Members("Member_Title_ID")
	myMember_Name = mySet_tb_Sites_Members("Member_Name")
	myMember_FirstName = mySet_tb_Sites_Members("Member_FirstName")
'	myMember_Pseudo = mySet_tb_Sites_Members("Member_Pseudo")
	myMember_Pseudo=myMember_Login
	myMember_Email = mySet_tb_Sites_Members("Member_Email")
	myMember_Company_Type = mySet_tb_Sites_Members("Member_Company_Type")
	if myMember_Company_Type=False then
		 myMember_Company_type=null 
	end if
	myMember_Company = mySet_tb_Sites_Members("Member_Company")
	' for future extension
'	myMember_Company_Activity_ID_selected = mySet_tb_Sites_Members("Member_Company_Activity_ID")
	myMember_Company_Activity_ID_Selected=0
	myMember_Company_Address = mySet_tb_Sites_Members("Member_Company_Address")
	myMember_Company_Zip = mySet_tb_Sites_Members("Member_Company_Zip")
	myMember_Company_City = mySet_tb_Sites_Members("Member_Company_City")
	myMember_Company_State = mySet_tb_Sites_Members("Member_Company_State")
	myMember_Company_Country_ID_Selected = mySet_tb_Sites_Members("Member_Company_Country_ID")
	myMember_Company_Phone = mySet_tb_Sites_Members("Member_Company_Phone")
	myMember_Company_Mobile = mySet_tb_Sites_Members("Member_Company_Mobile")
	myMember_Company_Fax = mySet_tb_Sites_Members("Member_Company_Fax")
	myMember_Company_Email = mySet_tb_Sites_Members("Member_Company_Email")
	myMember_Company_Web = mySet_tb_Sites_Members("Member_Company_Web")
	myMember_Company_Fonction = mySet_tb_Sites_Members("Member_Company_Fonction")
	myMember_Home_Type = mySet_tb_Sites_Members("Member_Home_Type")
	if myMember_Home_Type=False then
		 myMember_Home_type=null 
	end if
	myMember_Home_Address = mySet_tb_Sites_Members("Member_Home_Address")
	myMember_Home_Zip = mySet_tb_Sites_Members("Member_Home_Zip")
	myMember_Home_City = mySet_tb_Sites_Members("Member_Home_City")
	myMember_Home_State = mySet_tb_Sites_Members("Member_Home_State")
	myMember_Home_Country_ID_Selected = mySet_tb_Sites_Members("Member_Home_Country_ID")
	myMember_Home_Phone = mySet_tb_Sites_Members("Member_Home_Phone")
	myMember_Home_Mobile = mySet_tb_Sites_Members("Member_Home_Mobile")
	myMember_Home_Fax = mySet_tb_Sites_Members("Member_Home_Fax")
	myMember_Home_Email = mySet_tb_Sites_Members("Member_Home_Email")
	myMember_Home_Web = mySet_tb_Sites_Members("Member_Home_Web")
	myMember_Comments = mySet_tb_Sites_Members("Member_Comments")
	myMember_type_ID= mySet_tb_Sites_Members("Member_Type_ID")

	myMember_Author_Update = mySet_tb_Sites_Members("Member_Author_Update")
	myMember_Date_Update = mySet_tb_Sites_Members("Member_Date_Update")


	else

		' Close Recordset 
		mySet_tb_Sites_Members.close
		Set mySet_tb_Sites_Members = Nothing
		' Close Connection
		myConnection.close
		set myConnection = nothing
		' And Go Back
		Response.Redirect("__Sites_Members_List.asp")
	end if


	' Close Recordset 
	mySet_tb_Sites_Members.close
	Set mySet_tb_Sites_Members = Nothing
			
	' Close Connection
	myConnection.close
	set myConnection = nothing



end if




%> 

<td valign="top" bgcolor="<%=mybgColor%>" Width="<%=myApplication_Width%>"> 

<form method="POST" action="<%=myPage%>" name="myForm" > 

<table border="0" Width="<%=myApplication_Width%>" BGColor="<%=myBGColor%>" cellpadding="5" cellspacing="1"> 

<%
' Table Title AND general hidden Fields
%>


<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"> 
<b><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" Color="<%=myApplicationTextColor%>"><%=myMessage_Member%>
: <%=mySite_Name%></FONT></b>
<INPUT TYPE="hidden" NAME="Member_Site_ID" VALUE="<%=myMember_Site_ID%>"> 
<INPUT TYPE="hidden" NAME="Member_ID" VALUE="<%=myMember_ID%>"> 
<INPUT TYPE="hidden" NAME="Action" VALUE="<%=myAction%>"> 
</td>
</tr> 

<%
' Login
%>

<% if myAction="New" then %> 
<tr> 
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Login%>* 
<br><%=myFormGetErrMsg("Member_Login")%></FONT></B>
</td>
<td align="left"  valign="top"> 
<INPUT TYPE="text" NAME="Member_Login" Value="<%=myMember_Login%>">&nbsp;
<% if myError=1 then %> 
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><FONT COLOR="#ff0000"><%=myError_Message_Login%></FONT></FONT></B> 
<%end if %>
</td>
</tr>
<% else%>
<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Login%></FONT></B>
</td>
<td align="left" valign="top"> 
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMember_Login%></font></b>
<INPUT TYPE="hidden" NAME="Member_Login" VALUE="<%=myMember_Login%>" >
</td>
</tr>
<%end if%> 
<%
'/NEW OR NOT
%> 

<%
' Password
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Password%>*<br><%=myFormGetErrMsg("Member_Password")%></FONT></B>
</td>
<td align="left" valign="top">
<INPUT TYPE="PASSWORD" NAME="Member_Password" Value="<%=myMember_Password%>"> 
<%
if myUser_type_ID=1 then
	Response.write("("&myMember_Password&")")
end if
%>
</td>
</tr> 

<%
' Password Confirmation
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Password%> 
(<%=myMessage_Confirmation%>)*<br><%=myFormGetErrMsg("Member_Password_Confirmation")%></FONT></B>
</td>
<td align="left" valign="top">
<INPUT TYPE="PASSWORD" NAME="Member_Password_Confirmation" Value="<%=myMember_Password_Confirmation%>">
<% if myError=2 then %> 
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><FONT COLOR="#ff0000"><%=myError_Message_Password_Confirmation%></FONT></FONT></B> 
<%end if %>
</td>
</tr>

<%
' Title (Civility)
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"  Color="<%=myBorderTextColor%>"><%=myMessage_Title%></FONT></B>
</td>
<td align="left" valign="top"> 
<SELECT NAME="Member_Title_ID">
<OPTION VALUE="0" <%if myMember_Title_Id=0 then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMember_Title_Id=1 then%>Selected<%end if%>><%=myMessage_Mister%></OPTION> 
<OPTION VALUE="2" <%if myMember_Title_Id=2 then%>Selected<%end if%>><%=myMessage_Misses%></OPTION> 
<OPTION VALUE="3" <%if myMember_Title_Id=3 then%>Selected<%end if%>><%=myMessage_Miss%></OPTION></SELECT> 
<!--
<OPTION VALUE="4" <%if myMember_Title_Id=4 then%>Selected<%end if%>><%=myMessage_Father%></OPTION> 
<OPTION VALUE="5" <%if myMember_Title_Id=5 then%>Selected<%end if%>><%=myMessage_Brother%></OPTION> 
<OPTION VALUE="6" <%if myMember_Title_Id=6 then%>Selected<%end if%>><%=myMessage_Sister%></OPTION> 
-->
</td>
</tr>

<%
' Firstname
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Firstname%></FONT></B>
</td>
<td align="left" valign="top">
<INPUT TYPE="text" NAME="Member_FirstName" Value="<%=myMember_FirstName%>"> 
</td>
</tr>


<%
' Name
%>


<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Name%> 
</FONT></B>
</td>
<td align="left" valign="top">
<INPUT TYPE="text" NAME="Member_Name" Value="<%=myMember_Name%>">
</td>
</tr>


<%
' Email
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Email%>* 
<br><%=myFormGetErrMsg("Member_Email")%></FONT></B>
</td>
<td align="left" valign="top">
<INPUT TYPE="text" NAME="Member_Email" Value="<%=myMember_Email%>">
</td>
</tr> 

<%
' TITLE - Company
%>

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"> 
<%=myMessage_Office%> : <INPUT TYPE="checkbox" NAME="Member_Company_Type" VALUE="checkbox" <%if len(myMember_Company_Type)>0 then %> CHECKED <% end if %>></font></b>
</td>
</tr>

<%
' Company - Fonction
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Fonction%></font></b>
</td>
<td align="left" >
<INPUT TYPE="text" NAME="Member_Company_Fonction" Value="<%=myMember_Company_Fonction%>">
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
<INPUT TYPE="text" NAME="Member_Company"  Value="<%=myMember_Company%>">
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
<INPUT TYPE="text" NAME="Member_Company_Address" Value="<%=myMember_Company_Address%>">
</td>
</tr> 


<%
' Company Zip Code
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Zip_Code%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Member_Company_Zip" Value="<%=myMember_Company_Zip%>">
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
<INPUT TYPE="text" NAME="Member_Company_City" Value="<%=myMember_Company_City%>">
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
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Member_Company_State" Value="<%=myMember_Company_State%>"></font>
</td>
</tr> 

<%
' Company Country
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>"><b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Country%></font></b>
</td>
<td align="left"> 

<% 
'''''''''''''''''''''''''''''''''
' Get Country
''''''''''''''''''''''''''''''''

' DB Connection 

set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String
mySQL_Select_tb_Countries = "SELECT * FROM tb_Countries order by Country"

set mySet_Tb_Countries = 	myConnection.Execute(mySQL_Select_tb_Countries) %> 

<P><select name="Member_Company_Country_ID_Selected" size="" tabindex="1"> 
<option value="0"  <%if myMember_Company_Country_ID_selected = 0 then%> SELECTED <%end if %> > 
<%=myMessage_Select%></option>

<%do while not mySet_Tb_Countries.eof
	myMember_Company_Country_ID = mySet_Tb_Countries("Country_ID")
	myMember_Company_Country = mySet_Tb_Countries("Country")
%> 

	<option value="<%=myMember_Company_Country_ID%>"  <%if myMember_Company_Country_ID_selected = myMember_Company_Country_ID then%> SELECTED <%end if %> > 

<%=myMember_Company_Country%></option> <%
	mySet_Tb_Countries.MoveNext
	loop
	%> 
</select>
</P>
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
<td align="left"><INPUT TYPE="text" NAME="Member_Company_Phone" Value="<%=myMember_Company_Phone%>">
</td>
</tr> 

<%
' Company Mobile 
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Mobile%>
</font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Member_Company_Mobile" Value="<%=myMember_Company_Mobile%>">
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
<INPUT TYPE="text" NAME="Member_Company_Fax" Value="<%=myMember_Company_Fax%>">
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
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Member_Company_Email" Value="<%=myMember_Company_Email%>"> 
</font>
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
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Member_Company_Web" Value="<%=myMember_Company_Web%>"></font>
</td>
</tr> 


<%
' HOME TITLE
%>

<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"> 
<%=myMessage_Home%> : <INPUT TYPE="checkbox" NAME="Member_Home_Type" VALUE="checkbox" <%if len(myMember_Home_Type)>0 then %> CHECKED <% end if %>></font></b>
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
<INPUT TYPE="text" NAME="Member_Home_Address" Value="<%=myMember_Home_Address%>">
</td>
</tr> 


<%
' Home Zip Code
%>

<tr><td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Zip_Code%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Member_Home_Zip" Value="<%=myMember_Home_Zip%>">
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
<INPUT TYPE="text" NAME="Member_Home_City" Value="<%=myMember_Home_City%>">
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
<INPUT TYPE="text" NAME="Member_Home_State" Value="<%=myMember_Home_State%>">
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
' Get Country					 '	
''''''''''''''''''''''''''''''''''


' DB Connection 
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Countries = "SELECT * FROM tb_Countries order by Country"

set mySet_Tb_Countries = myConnection.Execute(mySQL_Select_tb_Countries) %> 
<P><select name="Member_Home_Country_ID_Selected" size="" tabindex="1"> 
<option value="0"  <%if myMember_Home_Country_ID_selected = 0 then%> SELECTED <%end if %> > 
<%=myMessage_Select%></option>


<%do while not mySet_Tb_Countries.eof
	myMember_Home_Country_ID = mySet_Tb_Countries("Country_ID")
	myMember_Home_Country = mySet_Tb_Countries("Country")
%> 

	<option value="<%=myMember_Home_Country_ID%>" <%if myMember_Home_Country_ID_selected = myMember_Home_Country_ID then%> SELECTED <%end if %> ><%=myMember_Home_Country%></option>

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
' Home Phone
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Phone%></font></b>
</td>
<td align="left">
<INPUT TYPE="text" NAME="Member_Home_Phone" Value="<%=myMember_Home_Phone%>">
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
<INPUT TYPE="text" NAME="Member_Home_Mobile" Value="<%=myMember_Home_Mobile%>">
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
<INPUT TYPE="text" NAME="Member_Home_Fax" Value="<%=myMember_Home_Fax%>">
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
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Member_Home_Email" Value="<%=myMember_Home_Email%>""></font>
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
<font face="Arial, Helvetica, sans-serif" size="2"><INPUT TYPE="text" NAME="Member_Home_Web" Value="<%=myMember_Home_Web%>"></font>
</td>
</tr> 

<%
' Separator
%>

<TR>
<TD VALIGN="top" ALIGN="right" COLSPAN="2" BGCOLOR="<%=myApplicationColor%>">&nbsp;

</TD>
</TR>

<%
' Comments
%>


<tr>
<TD VALIGN="top" ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Comments%> 
</FONT></B>
</TD>
<TD VALIGN="top" ALIGN="left">
              <textarea name="Member_Comments" cols="50" rows="5"><%=myMember_Comments%></textarea>
</TD>
</tr>
 


<%
' USER Type 
%>

<% if myUser_Type_ID=1 then %>
<tr>
<TD VALIGN="top" ALIGN="right" BGCOLOR="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Member_Type%>
</FONT></B>
</TD>
<TD VALIGN="top"  ALIGN="left">
              <input type="radio" name="Member_Type_ID" value="1" <%if myMember_Type_ID=1 then %> CHECKED <% end if %>>
              <%=myMessage_Administrator%><br>
<!--
              <input type="radio" name="Member_Type_ID" value="2" <%if myMember_Type_ID=2 then %> CHECKED <% end if %>>
              <%=myMessage_Moderator%> <br>
-->

             <input type="radio" name="Member_Type_ID" value="3" <%if myMember_Type_ID=3 OR len(myMember_Type_ID) = 0 then %> CHECKED <% end if %>>
              <%=myMessage_Intranet_Member%><br>

<!--

             <input type="radio" name="Member_Type_ID" value="4" <%if myMember_Type_ID=4 then %> CHECKED <% end if %>>


              <%=myMessage_Extranet_Member%><br>
              <input type="radio" name="Member_Type_ID" value="5" <%if myMember_Type_ID=5 then %> CHECKED <% end if %>>
              <%=myMessage_Web_Member%> <br>
              <input type="radio" name="Member_Type_ID" value="6" <%if myMember_Type_ID=6 then %> CHECKED <% end if %>>
              <%=myMessage_Identified_Email%><br>

-->
</TD>
</tr> 
<%end if%> 

<%
' Validation
%>

<TR> 
<TD ALIGN="right" VALIGN="top" BGCOLOR="<%=myBorderColor%>">&nbsp;

</TD>
<TD VALIGN="top" ALIGN="left">
<INPUT TYPE="submit" VALUE="<%=myMessage_Go%>" NAME="Validation">
</TD>
</TR>


<%
' Date - Author
%>

<TR>
<TD VALIGN="top" ALIGN="right" COLSPAN="2" BGCOLOR="<%=myApplicationColor%>">
<P ALIGN="center"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myApplicationTextColor%>">&nbsp;
<% if len(myMember_Date_Update) > 0 then %> 
	<% = myDate_Display(myMember_Date_Update,2) %> -- <% = myMember_Author_Update %>
<% end if %>
</FONT></P>
</TD>
</TR>
 



</table>

</form>



<TABLE BORDER="0" WIDTH="90%" CELLPADDING="3" CELLSPACING="0"> 

<%
' ADMINISTRATION
%> 

<% if (myMember_Site_ID=mySite_ID AND myUser_Type_ID=1) then %> 
<table border="0" width="<%=myApplication_Width%>" cellpadding="5" cellspacing="0"> 
<tr>
<td>
<a href="__Site_Member_Modification.asp?Action=New"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Add%>&nbsp; 
<%=myMessage_Member%></font></a> </td></tr> 
<% If myAction="Update" then %>
<TR><TD> 
<% if myMember_Type_ID<>1 AND myUser_Type_ID=1 then %> 
<A HREF="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Site_Member_Modification.asp?Action=Delete&amp;Member_ID=<%=myMember_ID%>';"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_delete%></font></A>
<% end if%>
</TD>
</TR>
<% End If ' Update %> 
</table>
<%end if ' me or Administrator %>
 
</TABLE>
</td>
<%
' / CENTER APPLICATION
%>

</TR>
</TABLE>
<%
' / CENTER 
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
<html><script language="JavaScript"></script></html>