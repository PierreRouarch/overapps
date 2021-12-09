<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   Pierre Rouarch
'
' This program "__Administration_Site_General.asp" is free software; 
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
' Does n't Work with PWS ???
%>
<!-- #include file="_INCLUDE/Global_Parameters.asp" -->
<!-- #include file="_INCLUDE/Form_validation.asp" -->
<!--#include file="_INCLUDE/Environment_Tools.asp"-->

<%
' ------------------------------------------------------------
' Name		: __Administration_Site_General.asp
' Path    	: /
' Description 	: Site Global Parameter
' By		: Pierre Rouarch, 
' Company	: 
' Date		: December, 10, 2001
' Versions : 1.18.0
'
' Contributions : Jean-Luc Lesueur, Christophe Humber, Dania Tcherkezoff
'
' Modify by	:
' Company	:
' Date		:
' ------------------------------------------------------------

Dim myPage
myPage = "__Administration_Site_General.asp"



%>

<!-- #include file="_INCLUDE/DB_Environment.asp" -->


<%

Dim myAuthor_Update, myDate_Update

Dim  mySQL_Select_tb_Sites_Activities, mySet_tb_Sites_Activities
Dim mySQL_Select_tb_Countries, mySet_tb_Countries
Dim myCountry_ID, myCountry

Dim mySQL_Select_tb_Applications 					
Dim mySet_tb_Applications

Dim mySQL_Delete_tb_Sites_Applications 					

Dim i,myTemp
Dim myMax_Applications

Dim myApplication_ID()
Dim myApplication_Name()



Dim mySite_Application_Title()
Dim mySite_Application_Public_Type_ID()
Dim mySite_Application_Opened()

Dim myApplication_Field_Name




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if Request.form("Validation")=myMessage_Modify then

	' Get Inputs 
	mySite_URL=Request.Form("Site_URL")
	mySite_Name=Request.Form("Site_Name")
	mySite_Presentation=Request.Form("Site_Presentation")
	
    'mySite_Public_Type_ID = Request.Form("Site_Public_Type_ID")
	mySite_Public_Type_ID=3

	mySite_Public_Type_ID = Cint(mySite_Public_Type_ID)
	if len(mySite_Public_Type_ID)=0 or mySite_Public_type_ID<1 or mySite_Public_type_ID>7 then 
		mySite_Public_type_ID=1
	end if 
	
	mySite_Extranet_Code = Request.Form("Site_Extranet_Code")

	' ADDRESS
	mySite_Company = Request.Form("Site_Company")
	mySite_Address = Request.Form("Site_Address")
	mySite_Zip = Request.Form("Site_Zip")
	mySite_City = Request.Form("Site_City")
	mySite_State = Request.Form("Site_State")
	mySite_Country_ID = Request.Form("Site_Country_ID")
	mySite_Country_ID=CInt(mySite_Country_ID)
	mySite_Phone = Request.Form("Site_Phone")
	mySite_Fax = Request.Form("Site_Fax")
	mySite_Web = Request.Form("Site_Web")
	mySite_Email = Request.Form("Site_Email")
	
		
	'Time Format
	myDate_Format = Request.Form("Date_Format")
	myHour_Format = Request.Form("Hour_Format")


	'Agenda Start and end time
	
	myAgenda_start = Cint(Request.Form("Agenda_Start"))
	myAgenda_End   = Cint(Request.Form("Agenda_End"))
	
	'End must be greater than start
	If  myAgenda_Start > myAgenda_End Then 
	 
	 myTemp = myAgenda_End
	 myAgenda_End = myAgenda_Start
	 myAgenda_Start = myTemp 
	 
	
	end if
	
	myAgenda_Week_Start = Cint(Request.Form("Agenda_Week_Start"))
		
	' For Future Extension Purpose
	mySite_Public_Directory = Request.Form("Site_Public_Directory")
	if len(mySite_Public_Directory)>0 then
		mySite_Public_Directory = True
	else
		mySite_Public_Directory = False
	end if ' site directory


	' Language
	mySite_Language_ID = Request.Form("Site_Language_ID")
	mySite_Language_ID=CInt(mySite_Language_ID)

	' Style : 
	mySite_Style_ID = Request.Form("Site_Style_ID")
	mySite_Style_ID=CInt(mySite_Style_ID)

	' Applications
	myMax_Applications=Request.Form("Max_Applications")


	myMax_Applications=Cint(myMax_Applications)

	ReDim myApplication_ID(myMax_Applications)
	ReDim mySite_Application_Title(myMax_Applications)
	ReDim mySite_Application_Public_Type_ID(myMax_Applications)
	ReDim mySite_Application_Opened(myMax_Applications)



	i=0
	do while i<=myMax_Applications


		
	myApplication_Field_Name="Application_ID_"&i
	

	myApplication_ID(i)=request.form(myApplication_Field_Name)

	myApplication_Field_Name="Site_Application_Opened_"&i
	mySite_Application_Opened(i)=request.form(myApplication_Field_Name)
	if len(mySite_Application_Opened(i))>0 then
			mySite_Application_Opened(i)=1
	else
			mySite_Application_Opened(i)=0
	end if

	myApplication_Field_Name="Site_Application_Title_"&i
	mySite_Application_Title(i)=request.form(myApplication_Field_Name)

	myApplication_Field_Name="Site_Application_Public_"&i
	mySite_Application_Public_Type_ID(i)=request.form(myApplication_Field_Name)
	mySite_Application_Public_Type_ID(i)=Cint(mySite_Application_Public_Type_ID(i))


	if mySite_Application_Public_Type_ID(i)>mySite_Public_Type_ID then 
		mySite_Application_Public_Type_ID(i)=mySite_Public_Type_ID
	end if
	i=i+1
	Loop

	' Test Inputs
	Call myFormSetEntriesInString
	' Only One Field Required
	myFormCheckEntry null, "Site_Name",true,null,null,0,100
	myFormCheckEntry null, "Site_Presentation",false,null,null,0,255

	if not myform_entry_error then

		myAuthor_Update = myUser_Pseudo
		myDate_Update = myDate_Now()

		'DB Connection
		set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String

		'Update in DB
		mySQL_Select_tb_Sites = "SELECT * FROM tb_Sites WHERE SITE_ID="&mySite_ID
		Set mySet_tb_Sites= Server.CreateObject("ADODB.Recordset")
    	mySet_tb_Sites.open mySQL_Select_tb_Sites, myConnection, 3, 3


		
		mySet_tb_Sites.fields("Site_URL")=mySite_URL
		mySet_tb_Sites.fields("Site_Name")=mySite_Name
		mySet_tb_Sites.fields("Site_Presentation")=mySite_Presentation
		mySet_tb_Sites.fields("Site_Public_Type_ID")=mySite_Public_Type_ID
 		mySet_tb_Sites.fields("Site_Extranet_Code")=mySite_Extranet_Code
 		mySet_tb_Sites.fields("Site_Company")= mySite_Company
 		mySet_tb_Sites.fields("Site_Address")=mySite_Address
 		mySet_tb_Sites.fields("Site_Zip")=mySite_zip
 		mySet_tb_Sites.fields("Site_City")=mySite_City
		mySet_tb_Sites.fields("Site_State")=mySite_State
 		mySet_tb_Sites.fields("Site_Country_ID")=mySite_Country_ID
 		mySet_tb_Sites.fields("Site_Phone")=mySite_Phone
 		mySet_tb_Sites.fields("Site_Fax")=mySite_Fax
 		mySet_tb_Sites.fields("Site_Email")=mySite_Email
 		mySet_tb_Sites.fields("Site_Web")=mySite_Web
		mySet_tb_Sites.fields("Site_Language_ID")=mySite_Language_ID
 		mySet_tb_Sites.fields("Site_Style_ID")=mySite_Style_ID
		
  		mySet_tb_Sites.fields("Site_Public_Directory")=mySite_Public_Directory
 		mySet_tb_Sites.fields("Site_Author_Update")=myAuthor_Update
 		mySet_tb_Sites.fields("Site_Date_Update")=myDate_Update
		
		mySet_tb_Sites.fields("Site_Date_Format") = myDate_Format
		mySet_tb_Sites.fields("Site_Hour_Format") = myHour_Format
		
		mySet_tb_Sites.fields("Site_Agenda_Start") = myAgenda_Start
		mySet_tb_Sites.fields("Site_Agenda_End")   = myAgenda_End
		mySet_tb_Sites.fields("Site_Agenda_Week_Start") = myAgenda_Week_Start
	
		mySet_tb_Sites.Update
		'Close Recordset
		mySet_tb_Sites.close
		Set mySet_tb_Sites = Nothing




		' UPDATE Sites_APPLICATIONS

		' Delete all Sites_Applications 
		mySQL_Delete_tb_Sites_Applications = "DELETE FROM tb_Sites_Applications WHERE Site_ID ="&mySite_ID
		 myConnection.Execute(mySQL_Delete_tb_Sites_Applications)


		' Re insert Sites Applications
		i=0
		do while i<myMax_Applications
			

		mySQL_Select_tb_Sites_Applications = "SELECT * FROM tb_Sites_Applications" 
		Set mySet_tb_Sites_Applications= Server.CreateObject("ADODB.Recordset")
    	mySet_tb_Sites_Applications.open mySQL_Select_tb_Sites_Applications, myConnection, 3, 3
		mySet_tb_Sites_Applications.AddNew

			mySet_tb_Sites_Applications.fields("Site_ID")=mySite_ID

			mySet_tb_Sites_Applications.fields("Application_ID")=myApplication_ID(i)
	
			mySet_tb_Sites_Applications.fields("Site_Application_Opened")= mySite_Application_Opened(i)
	
			mySet_tb_Sites_Applications.fields("Site_Application_Title")= mySite_Application_Title(i)
			mySet_tb_Sites_Applications.fields("Site_Application_Public_Type_ID")=			 mySite_Application_Public_Type_ID(i)
 			mySet_tb_Sites_Applications.fields("Site_Application_Author_Update")= myAuthor_Update
 			mySet_tb_Sites_Applications.fields("Site_Application_Date_Update")=myDate_Update
	
			mySet_tb_Sites_Applications.Update

		'Close Recordset
		mySet_tb_Sites_Applications.close
		Set mySet_tb_Sites_Applications = Nothing

		i=i+1
		Loop

		' Close Connection
		myConnection.close
		set myConnection = nothing	
	


		Response.Redirect("__Administration_Site.asp")
		
	end if ' Not Entry Error

end if ' Validation


%>
<html>

<head>
<title><%=mySite_Name%> - Administration - <%=myMessage_General_Parameters%></title>

</head>

<BODY BackGround="<%=myBGImage%>" bgColor="<%=myBGColor%>"  Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">

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

<TD WIDTH="<%=myLeft_Width%>"> 
<!-- #include file="_borders/Left.asp" --> 
</td>

<%
' CENTER APPLICATION
%> 

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form	- Information coming from  DB_Environment.asp		'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>


<td WIDTH="<%=myApplication_Width%>" BGCOLOR="<%=myBGCOLOR%>" valign="top"> 
<form method="POST" action="<%=myPage%>" name="myForm"> 

        <table border="0" Width="<%=myApplication_Width%>" bgcolor="<%=myBGColor%>" cellpadding="5" cellspacing="1">
          <%
' TITLE
%>
          <tr> 
            <td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"> 
              <font color="<%=myApplicationTextColor%>" face="Arial, Helvetica, sans-serif" size="4"><b><%=myMessage_Administration%> 
              : <%=myMessage_General_Parameters%> </b></font></td>
          </tr>
          <%
' URL ADDRESS
%>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_URL_Address%></FONT></B> 
            </td>
            <td align="left" width="72%"> 
              <INPUT TYPE="text" SIZE="60" NAME="Site_URL" Value="<%=mySite_URL%>">
              <INPUT TYPE="hidden" NAME="Site_ID" VALUE="<%=mySite_ID%>">
            </td>
          </tr>
          <%
' NAME
%>
          <TR> 
            <TD ALIGN="right" BGCOLOR="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Name%>*<br>
              <%=myFormGetErrMsg("Site_Name")%></Font></B> </TD>
            <TD ALIGN="left" width="72%" > <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> 
              <INPUT TYPE="text" SIZE="60" NAME="Site_Name" Value="<%=mySite_Name%>">
              </FONT> </TD>
          </TR>
          <%
' Presentation
%>
          <TR> 
            <TD ALIGN="right" BGCOLOR="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Presentation%><br>
              <%=myFormGetErrMsg("Site_Presentation")%></FONT></B> </TD>
            <TD ALIGN="left" width="72%"> <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> 
              <TEXTAREA NAME="Site_Presentation" ROWS="4" COLS="60" wrap="PHYSICAL">
<%=mySite_Presentation%> 
</TEXTAREA>
              </FONT> </TD>
          </TR>
          <%
' Company
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Company%></font></b> 
            </td>
            <td align="left" width="72%" ><font face="Arial, Helvetica, sans-serif" size="2"> 
              <INPUT TYPE="text" NAME="Site_Company" VALUE="<%=mySite_Company%>" SIZE="60">
              </font> </td>
          </tr>
          <%
' Address
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Address%></font></b> 
            </td>
            <td align="left" width="72%" > 
              <INPUT TYPE="text" NAME="Site_Address" VALUE="<%=mySite_Address%>" SIZE="60">
            </td>
          </tr>
          <%
' Zip Code
%>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"> <b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Zip_Code%></font></b> 
            </td>
            <td align="left" width="72%" ><font face="Arial, Helvetica, sans-serif" size="2"> 
              <INPUT TYPE="text" NAME="Site_Zip" value="<%=mySite_Zip%>" SIZE="20">
              </font> </td>
          </tr>
          <%
' City
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_City%></FONT></B> 
            </td>
            <td align="left" width="72%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
              <INPUT TYPE="text" NAME="Site_City" VALUE="<%=mySite_City%>" SIZE="60">
              </font> </td>
          </tr>
          <%
' State
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_State%></FONT></B> 
            </td>
            <td align="left" width="72%" > <font face="Arial, Helvetica, sans-serif" size="2"> 
              <INPUT TYPE="text" NAME="Site_State" VALUE="<%=mySite_State%>" SIZE="60">
              </font> </td>
          </tr>
          <%
' Country
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Country%></font></b> 
            </td>
            <td align="left" width="72%"> 
              <% 
''''''''''''''''''''''''''''''''''
' Get Country			 '	
''''''''''''''''''''''''''''''''''
' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String
mySQL_Select_tb_Countries = "SELECT * FROM tb_Countries order by Country"
set mySet_Tb_Countries = 	myConnection.Execute(mySQL_Select_tb_Countries) %>
              <P> 
                <select name="Site_Country_ID" size="" tabindex="1">
                  <option value="0"  <%if mySite_Country_ID = 0 then%> SELECTED <%end if %> > 
                  <%=myMessage_Select%></option>
                  <%do while not mySet_Tb_Countries.eof
	myCountry_ID = mySet_Tb_Countries("Country_ID")
	myCountry = mySet_Tb_Countries("Country")
%>
                  <option value="<%=myCountry_ID%>"  <%if mySite_Country_ID = myCountry_ID then%> SELECTED <%end if%>><%=myCountry%></option>
                  <%
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
' Phone
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Phone%></FONT></B> 
            </td>
            <td align="left" width="72%" ><font face="Arial, Helvetica, sans-serif" size="2"> 
              <INPUT TYPE="text" NAME="Site_Phone" VALUE="<%=mySite_Phone%>">
              </font> </td>
          </tr>
          <%
' Fax
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Fax%></FONT></B> 
            </td>
            <td align="left" width="72%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
              <INPUT TYPE="text" NAME="Site_Fax" VALUE="<%=mySite_Fax%>">
              </font> </td>
          </tr>
          <%
' Other Web
%>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Other_Web_Site%></FONT></B> 
            </td>
            <td align="left" width="72%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
              <INPUT TYPE="text" NAME="Site_Web" VALUE="<%=mySite_Web%>" SIZE="60">
              </font> </td>
          </tr>
          <%
' Email
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Email%><BR>
              </FONT></B> </td>
            <td align="left" width="72%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
              <INPUT TYPE="text" NAME="Site_Email" VALUE="<%=mySite_Email%>" SIZE="60">
              </font> </td>
          </tr>
          <%
' Language
%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Language%><BR>
              </FONT></B> </td>
            <td align="left" width="72%"> 
              <select name="Site_Language_ID">
                <option value="5" <%if mySite_Language_ID = 5 then%> SELECTED <%end if%>>Deutsch</option>                                                         
				<option value="1" <%if mySite_Language_ID = 1 then%> SELECTED <%end if%>>English</option>
                <option value="3" <%if mySite_Language_ID = 3 then%> SELECTED <%end if%>>Español</option>
     			<option value="2" <%if mySite_Language_ID = 2 then%> SELECTED <%end if%>>Français</option>
				<option value="6" <%if mySite_Language_ID = 6 then%> SELECTED <%end if%>>Italiano</option>
                <option value="4" <%if mySite_Language_ID = 4 then%> SELECTED <%end if%>>Portuguese</option>
              </select>
            </td>
          </tr>
<% 'Date and Hour Format %>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><b><%=myMessage_Date_Format%></b></font></td>
            <td align="left" width="72%"> 
              <input type="radio" name="Date_Format" value="1"
<%
If myDate_Format = 1 Then response.Write " checked "

 
   


%>			  
			  >
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Date_US %><br>
              <input type="radio" name="Date_Format" value="2"
<%
If myDate_Format = 2 Then response.Write " checked "
%>			  
					  
			  >
              <%= myMessage_Date_Europe  %></font></td>
          </tr>
          <tr>
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><b><%=myMessage_Hour_Format%></b></font></td>
            <td align="left" width="72%"> 
              <input type="radio" name="Hour_Format" value="1"
<%
If myHour_Format = 1 Then response.Write " checked "
%>					  
			  >
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%= myMessage_Hour_US   %><br>
              <input type="radio" name="Hour_Format" value="2"
<%
If myHour_Format = 2 Then response.Write " checked "
%>					  
			  >
              <%=myMessage_Hour_Europ  %></font></td>
          </tr>
<%
'AGENDA START AND END TIME
%>

<tr> 
<td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%= myMessage_Agenda_Start %><BR>
</FONT></B> </td>
<td align="left" width="72%"> 
<select name=Agenda_Start>
<%
for i = 0 to 23
%>
<option value=<%= i %> <% If myAgenda_Start=i then response.Write " selected " %>>
<%
if myHour_Format = 1 and i < 10 Then response.write "0" & i & ":00 A.M "
if myHour_Format = 1 and i < 12 and i > 9 Then response.write i & ":00 A.M " 
if myHour_Format = 1 and i = 12 Then response.write i & ":00 P.M " 
if myHour_Format = 1 and i > 12 and i < 22 Then response.write "0" &  i - 12 & ":00 P.M "
if myHour_Format = 1 and i > 12 and i > 21 Then response.write i - 12 & ":00 P.M "
if myHour_Format <> 1 and i < 10 Then response.write "0"  & i & ":00"
if myHour_Format <> 1 and i > 9 Then response.write  i & ":00"
%>

<%
next 
%>
</select>              
            </td>
          </tr>

<tr> 
<td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%= myMessage_Agenda_End %><BR>
</FONT></B> </td>
<td align="left" width="72%"> 
<select name=Agenda_End>
<%
for i = 0 to 23
%>
<option value=<%= i %> <% If myAgenda_End=i then response.Write " selected " %>>
<%
if myHour_Format = 1 and i < 10 Then response.write "0" & i & ":00 A.M "
if myHour_Format = 1 and i < 12 and i > 9 Then response.write i & ":00 A.M " 
if myHour_Format = 1 and i = 12 Then response.write i & ":00 P.M " 
if myHour_Format = 1 and i > 12 and i < 22 Then response.write "0" & i - 12 & ":00 P.M "
if myHour_Format = 1 and i > 12 and i > 21 Then response.write i - 12 & ":00 P.M "
if myHour_Format <> 1 and i < 10 Then response.write "0"  & i & ":00"
if myHour_Format <> 1 and i > 9 Then response.write  i & ":00"
%>

<%
next 
%>
</select>              
            </td>
          </tr>


<%
'Agenda  Day Start
%>

<tr> 
<td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%= myMessage_Agenda_Week_start  %><BR>
</FONT></B> </td>
<td align="left" width="72%"> 
<select name=Agenda_Week_Start>
<option value=0 <%If myagenda_Week_Start=0 Then response.write"selected" %>><%=MyMessage_Sunday%></option>
<option value=1 <%If myagenda_Week_Start=1 Then response.write"selected" %>><%=MyMessage_Monday%></option>

</select>              
            </td>
          </tr>


<%
' Public
%>
          <!--

<tr> 
<td align="right"  bgcolor="<%=myBorderColor%>">
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Site_Public%><BR></FONT></B> 
</td>
<td align="left" > 
 


<% 
' Default Public For the Site
%>


Future Extension

	<select name="Site_Public_Type_ID">
        <option value="1" <%if mySite_Public_Type_ID = 1 then%> SELECTED <%end if%>> <%=myMessage_Administrator%></option>
        <option value="2" <%if mySite_Public_Type_ID = 2 then%> SELECTED <%end if%>> <%=myMessage_Moderator%></option>
	    <option value="3" <%if mySite_Public_Type_ID = 3 then%> SELECTED <%end if%>> <%=myMessage_Intranet_Member%></option>

		<option value="4" <%if mySite_Public_Type_ID = 4 then%> SELECTED <%end if%>>
 <%=myMessage_Extranet_Member%></option>
        <option value="5" <%if mySite_Public_Type_ID = 5 then%> SELECTED <%end if%>>
 <%=myMessage_Web_Member%></option>
        <option value="6" <%if mySite_Public_Type_ID = 6 then%> SELECTED <%end if%>> <%=myMessage_Identified_Email%></option>
		<option value="7" <%if mySite_Public_Type_ID = 7 then%> SELECTED <%end if%>> <%=myMessage_Public%></option>

 </select>

</td></tr>


-->
          <%
' Style
%>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"><b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Style%></font></b> 
            </td>
            <td align="left" width="72%"> 
              <% 
'''''''''''''''''''''''''''''''''
' Get the Style 				'	
'''''''''''''''''''''''''''''''''

' Database Connection 
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Styles = "SELECT * FROM tb_Styles Where  Site_ID="&mySite_ID &" order by Style_Name"

set mySet_tb_Styles = 	myConnection.Execute(mySQL_Select_tb_Styles) %>
              <P> 
                <select name="Site_Style_ID" size="" tabindex="1">
                  <option value="0" <%if mySite_Style_ID=0 then%> SELECTED <%end if%>><%=myMessage_Select%></option>
                  <%do while not mySet_tb_Styles.eof
	myStyle_ID = mySet_tb_Styles("Style_ID")
	myStyle_Name = mySet_tb_Styles("Style_Name")
%>
                  <option value="<%=myStyle_ID%>" <%if mySite_Style_ID = myStyle_ID then%> SELECTED <%end if %> > 
                  <%=myStyle_Name%></option>
                  <%
	mySet_tb_Styles.MoveNext
	loop
%>
                </select>
                &nbsp;&nbsp;<font face=Arial size=2 color="<%= myBGTextColor %>"><a href=__Styles_list.asp><%= myStyles_Administration %></a></font></P>
              <%
' Close
mySet_tb_Styles.close
Set mySet_tb_Styles = Nothing
myConnection.Close
set myConnection = Nothing
%>
            </td>
          </tr>
          <%
''''''''''''''''
' APPLICATIONS '
''''''''''''''''
%>
          <%
' Title
%>
          <tr> 
            <td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"><b><font face="Arial, Helvetica, sans-serif" size="3" color="<%=myApplicationTextColor%>"> 
              <%=myMessage_Applications%></font></b> </td>
          </tr>
          <%
' Applications Colums
%>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%">&nbsp; </td>
            <td width="72%"> 
              <Table >
                <TR> 
                  <TD><font face=Arial size=2 color="<%= myBGTextColor %>"> <b><%=myMessage_Application%></b> </font></td>
                  <td align=center><font face=Arial size=2 color="<%= myBGTextColor %>"> <b><%=myMessage_Opened%></b> </font></td>
                  <td align=center><font face=Arial size=2 color="<%= myBGTextColor %>"> <b><%=myMessage_Title%></b> </font></td>
                  <td align=center><font face=Arial size=2 color="<%= myBGTextColor %>"> <b><%=myMessage_Public%></b> </font></td>
                </tr>
                <%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  APPLICATIONS														 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Applications ="  Select * from tb_Applications"  

set mySet_tb_Applications = myConnection.Execute(mySQL_Select_tb_Applications)

i=0

do while not mySet_tb_Applications.eof
			
ReDim myApplication_ID(i+1)
ReDim myApplication_Name(i+1)
ReDim myApplication_Title(i+1)
ReDim myApplication_Public_Type_ID(i+1)


ReDim mySite_Application_Title(i+1)
ReDim mySite_Application_Public_Type_ID(i+1)
ReDim mySite_Application_Opened(i+1)

	myApplication_ID(i)  = mySet_tb_Applications("Application_ID")
	myApplication_Name(i) = mySet_tb_Applications("Application_Name")
	myApplication_Title(i)	= mySet_tb_Applications("Application_Title")
	myApplication_Public_Type_ID(i) = mySet_tb_Applications("Application_Public_Type_ID")
				
	mySQL_Select_tb_Sites_Applications = "Select * from tb_Sites_Applications WHERE  Site_ID="&mySite_ID&" AND Application_ID="&myApplication_ID(i)
	set mySet_tb_Sites_Applications = myConnection.Execute(mySQL_Select_tb_Sites_Applications)
	if not mySet_tb_Sites_Applications.eof then
		mySite_Application_Title(i) = mySet_tb_Sites_Applications("Site_Application_Title")
		mySite_Application_Opened(i) = mySet_tb_Sites_Applications("Site_Application_Opened")
		mySite_Application_Public_Type_ID(i) = mySet_tb_Sites_Applications("Site_Application_Public_Type_ID")
	end if

	if len(mySite_Application_Title(i))=0 then 
			mySite_Application_Title(i)=myApplication_Title(i)
	end if

	if len(mySite_Application_Opened(i))=0 then 
			mySite_Application_Opened(i)=0
	end if

	if len(mySite_Application_Public_Type_ID(i))=0 then 
			mySite_Application_Public_Type_ID(i)=myApplication_Public_Type_ID(i)
	end if

%>
                <TR> 
                  <TD> <font face=Arial size=2 color="<%= myBGTextColor %>"><%=myApplication_Name(i)%>&nbsp;</font> 
                    <%
	myApplication_Field_Name="Application_ID_"&i
	%>
                    <input type="hidden" name="<%=myApplication_Field_Name%>" value="<%=myApplication_ID(i)%>">
                  </TD>
                  <TD align=center> 
                    <%
	myApplication_Field_Name="Site_Application_Opened_"&i
	%>
                    <input type="checkbox" name="<%=myApplication_Field_Name%>"
	<%if not mySet_tb_Sites_Applications.eof and mySite_Application_Opened(i)=1 then%> value="on" checked<%end if%>>
                  </TD>
                  <TD> 
                    <%
	myApplication_Field_Name="Site_Application_Title_"&i
	%>
                    <input type="text" name="<%=myApplication_Field_Name%>" value="<%=mySite_Application_Title(i)%>">
                  </TD>
                  <TD> 
                    <%
	myApplication_Field_Name="Site_Application_Public_"&i
	%>
                    <select name="<%=myApplication_Field_Name%>">
                      <option value="1" <%if mySite_Application_Public_Type_ID(i) = 1 then%> SELECTED <%end if%>> 
                      <%=myMessage_Administrator%></option>
                      <!--
		<% if mySite_Public_type_ID>=2 then %>
        <option value="2" <%if mySite_Application_Public_Type_ID(i) = 2 then%> SELECTED <%end if%>> <%=myMessage_Moderator%></option>
		<% end if %>
-->
                      <% if mySite_Public_type_ID>=3 then %>
                      <option value="3" <%if mySite_Application_Public_Type_ID(i) = 3 then%> SELECTED <%end if%>> 
                      <%=myMessage_Intranet_Member%></option>
                      <% end if %>
                      <!--

		<% if mySite_Public_type_ID>=4 then %>
		<option value="4" <%if mySite_Application_Public_Type_ID(i) = 4 then%> SELECTED <%end if%>> <%=myMessage_Extranet_Member%></option>
		<% end if %>

		<% if mySite_Public_type_ID>=5 then %>
        <option value="5" <%if mySite_Application_Public_Type_ID(i) = 5 then%> SELECTED <%end if%>> <%=myMessage_Web_Member%></option>
		<% end if %>

		<% if mySite_Public_type_ID>=6 then %>
        <option value="6" <%if mySite_Application_Public_Type_ID(i) = 6 then%> SELECTED <%end if%>> <%=myMessage_Identified_Email%></option>
		<% end if %>

		<% if mySite_Public_type_ID>=7 then %>
		<option value="7" <%if mySite_Application_Public_Type_ID(i) = 7 then%> SELECTED <%end if%>> <%=myMessage_Public%></option>
		<% end if %>

-->
                    </select>
                  </TD>
                </TR>
                <%
	mySet_tb_Applications.movenext
	i=i+1
	loop


' Close
mySet_tb_Applications.close
Set mySet_tb_Applications=nothing
myConnection.close
set myConnection = nothing


myMax_Applications=i
%>
              </table>
              <input type="hidden" name="Max_Applications" value="<%=myMax_Applications%>">
            </td>
          </tr>
          <%
' VALIDATION
%>
          <TR> 
            <TD ALIGN="right" VALIGN="top" BGCOLOR="<%=myBorderColor%>" width="28%">&nbsp; 
            </TD>
            <TD VALIGN="top" ALIGN="left" width="72%"> 
              <INPUT TYPE="submit" VALUE="<%=myMessage_Modify%>" NAME="Validation">
            </TD>
          </TR>
          <%
' End 
%>
          <TR> 
            <TD VALIGN="top" ALIGN="right" COLSPAN="2" BGCOLOR="<%=myApplicationColor%>"> 
              <P ALIGN="center">&nbsp;</P>
            </TD>
          </TR>
        </table>
</form>
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
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</body>
</html>

