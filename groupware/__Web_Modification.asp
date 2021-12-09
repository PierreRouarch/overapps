<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Web_modification.asp" is free software; 
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
' Doesn't Work with PWS ???
%>


<%
'------------------------------------------------------------
' Name 			: __Web_modification.asp
' Path    		: /
' Version 		: 1.13.0
' Description 	: Add/Modify/Delete Web
' By 			: Pierre Rouarch
' Company		: OverApps
' Date			: October 10, 2001
' Version       : 1.15.0
' Contributions : Jean-Luc Lesueur, Christophe Humbert
'
' Modify by 	:
' Company		:
' Date			:
' ------------------------------------------------------------

Dim myPage, myNewSite_ID
myPage = "__Web_Modification.asp"

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

Dim myWebDirectory_ID, myCategory_ID, myMember_ID, myWeb_ID, myWeb_URL, myWeb_Name,   myWeb_Description_Short, myWeb_Description_Long, myWeb_Public, myWeb_Top, myWeb_Author_Update, myWeb_Date_Update

Dim myAction, myTitle, myBoutton, myReset
					
Dim mySQL_Select_tb_Webs, mySQL_Delete_tb_Webs, mySQL_Insert_tb_Webs, mySQL_Update_tb_Webs, myQuery, mySet_tb_Webs


' Get Parameters
' What do We Do now ?
myAction = Request("Action")
if len(myAction)=0 then 
	myAction="Update"
end if

myWeb_ID  = Request("Web_ID") 
if len(myWeb_ID)=0 then 
	myAction="New"
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Delete
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction= "Delete" then
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

	myWeb_ID  = Request.QueryString("Web_ID") 
	mySQL_Delete_tb_Webs = "DELETE FROM tb_Webs WHERE Site_ID="&mySite_ID&" AND Web_ID="&myWeb_ID

	myConnection.Execute(mySQL_Delete_tb_Webs)
	myConnection.Close
	set myConnection = Nothing
	' And go back to the list
	Response.Redirect("__Webs_List.asp")
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if Request.form("Validation")=myMessage_Go then

	' Get Entries
	myWeb_ID = Request.Form("Web_ID")
	myWeb_Name = Replace(Request.Form("Web_Name"),"'"," ")
	myWeb_URL = Replace(Request.Form("Web_URL"),"'"," ")
	myWeb_Description_Short   = Replace(Request.Form("Web_Description_Short"),"'"," ")
	myWeb_Description_Long   = Replace(Request.Form("Web_Description_Long"),"'"," ")

	' Test Entriee
	Call myFormSetEntriesInString

	myFormCheckEntry null, "Web_Name",true,null,null,0,255
	myFormCheckEntry null, "Web_URL",true,null,null,0,255
	myFormCheckEntry null, "Web_Description_Short",false,null,null,0,255
	
	if not myform_entry_error then
	
		' Intialization in this version 
		myWebDirectory_ID = 1
		myCategory_ID = 1
		myWeb_Top=False
		myWeb_Public=False
		myWeb_Author_Update = myUser_Pseudo
		myWeb_Date_Update = myDate_Now()

		' db Connection
		set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String

		If myAction = "New" Then

			mySQL_Select_tb_Webs = "SELECT * FROM tb_webs"			
			Set mySet_tb_Webs = server.createobject("adodb.recordset")
		   	mySet_tb_Webs.open mySQL_Select_tb_Webs, myConnection, 3,3
		   	mySet_tb_Webs.AddNew
	
			mySet_tb_Webs.fields("WebDirectory_ID") = myWebDirectory_ID
			mySet_tb_Webs.fields("Category_ID") = myCategory_ID
			mySet_tb_Webs.fields("Site_ID") = mySite_ID
			mySet_tb_Webs.fields("Member_ID") = myUser_ID
			mySet_tb_Webs.fields("Web_URL") = myWeb_URL
			mySet_tb_Webs.fields("Web_Name") = myWeb_Name
			mySet_tb_Webs.fields("Web_Description_Short") = myWeb_Description_Short
			mySet_tb_Webs.fields("Web_Description_Long") = myWeb_Description_Long
			mySet_tb_Webs.fields("Web_Top") = myWeb_Top
			mySet_tb_Webs.fields("Web_Public") = myWeb_Public
			mySet_tb_Webs.fields("Web_Author_Update") = myWeb_Author_Update
			mySet_tb_Webs.fields("Web_Date_Update") = myWeb_Date_Update
		
			mySet_tb_Webs.Update
			mySet_tb_Webs.close
			Set mySet_tb_Webs = Nothing
	
		ElseIf myAction = "Update" Then

			mySQL_Select_tb_Webs = "SELECT * FROM tb_webs WHERE Site_ID="&mySite_ID&" AND Web_ID="&myWeb_ID			
			Set mySet_tb_Webs = server.createobject("adodb.recordset")
		   	mySet_tb_Webs.open mySQL_Select_tb_Webs, myConnection, 3,3

			mySet_tb_Webs.fields("Web_Name") = myWeb_Name 
			mySet_tb_Webs.fields("Web_URL") = myWeb_URL 
			mySet_tb_Webs.fields("Web_Description_Short") = myWeb_Description_Short 
			mySet_tb_Webs.fields("Web_Description_Long") = myWeb_Description_Long
			mySet_tb_Webs.fields("Web_Author_Update") = myWeb_Author_Update 
			mySet_tb_Webs.fields("Web_Date_Update") = myWeb_Date_Update 
	
			mySet_tb_Webs.Update
			mySet_tb_Webs.close
			Set mySet_tb_Webs = Nothing
	
		end if

		' Close Connection 
		myConnection.close
		set myConnection = nothing	
		' And go Back to List
		Response.Redirect("__Webs_List.asp")
		
	end if 

end if 

%>

<html>

<head>
<title><%=mySite_Name%> Web - Add/Modify/Delete</title>

</head>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">

<%
' TOP
%>

<!-- #include file="_borders/top.asp" --> 

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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form 													'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

if myAction = "Update" and Request.form("Validation")<>myMessage_Go then
	' Get information 
	myQuery = "SELECT * FROM tb_webs WHERE Site_ID="&mySite_ID&" AND Web_ID="&myWeb_ID
	set mySet_tb_Webs = myConnection.Execute(myQuery)

	' If Nothing Go Back
	if mySet_tb_Webs.eof then
		' Close Recordset
		mySet_tb_Webs.close
		Set mySet_tb_Webs=Nothing
		' Close Connection
		myConnection.close
		set myConnection = nothing
		Response.Redirect("__Webs_List.asp")
	end if

	myWebDirectory_ID = mySet_tb_Webs("WebDirectory_ID")
	myCategory_ID = mySet_tb_Webs("WebDirectory_ID")

	mySite_ID = mySet_tb_Webs("Site_ID")
	myWebDirectory_ID = mySet_tb_Webs("WebDirectory_ID")
	myMember_ID = mySet_tb_Webs("Member_ID")
 
	myWeb_ID = mySet_tb_Webs("Web_ID")
	myWeb_Name  = mySet_tb_Webs("Web_Name")
	myWeb_URL = mySet_tb_Webs("Web_URL")
	myWeb_Description_Short  = mySet_tb_Webs("Web_Description_Short")
	myWeb_Description_Long  = mySet_tb_Webs("Web_Description_Long")
	myWeb_Top  = mySet_tb_Webs("Web_Top")
	myWeb_Public  = mySet_tb_Webs("Web_Public")

	myWeb_Author_Update   = mySet_tb_Webs("Web_Author_Update")				
	myWeb_Date_Update     = mySet_tb_Webs("Web_Date_Update")

	' Close Recordset
	mySet_tb_Webs.close
	Set mySet_tb_Webs=Nothing

end if


' Close Connection
myConnection.close
set myConnection = nothing

%>



<td valign="top" Width="<%=myApplication_Width%>" bgcolor="<%=myBGColor%>"> 
<form method="POST" action="<%=myPage%>" name="myForm" >
<table border="0" Width="<%=myApplication_Width%>" cellpadding="5" cellspacing="1"> 

<%
' Application Title AND Hidden Fields
%>

<tr> 
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"> 
<b><FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR="<%=myApplicationTextColor%>"><%=myApplication_Title%></FONT></b>

<INPUT TYPE="hidden" NAME="Web_ID" VALUE="<%=myWeb_ID%>"> 
<INPUT TYPE="hidden" NAME="WebDirectory_ID" VALUE="<%=myWebDirectory_ID%>">
<INPUT TYPE="hidden" NAME="Action" VALUE="<%=myAction%>"> 
</td>
</tr>


<%
' Web Address (URL)
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Address%>&nbsp(URL)*<BR> 
<%=myFormGetErrMsg("Web_URL")%></B></FONT>
</td>
<td align="left"  valign="top"> 
<input type="text" size="30" name="Web_URL" value="<%=myWeb_URL%>">&nbsp;( 
http://www.xyz.com)
</td>
</tr>

<%
' Web Name
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Name%>*<BR> <%=myFormGetErrMsg("Web_Name")%></B></font>
</td>
<td align="left" valign="top"> 
<input type="text" size="30" name="Web_Name" value="<%=myWeb_Name %>">
</td>
</tr>

<%
' Web Presentation
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Presentation%><BR> <%=myFormGetErrMsg("Web_Description_Short")%></B></font>
</td>
<td align="left" valign="top"> 
<input type="text" size="50" name="Web_Description_Short" value="<%=myWeb_Description_Short%>"> 
</td>
</tr>

<%
' Web description
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_More%></b></font>
</td>
<td valign="top" align="left"> 
<TEXTAREA NAME="Web_Description_Long" COLS="50" ROWS="5"><%=myWeb_Description_Long%></TEXTAREA> 
</td>
</tr>

<%
' Validation
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">&nbsp;

</td>
<td valign="top" align="left"> 
<input type="submit" value="<%=myMessage_Go%>" name="Validation">
</td>
</tr>

<%
' Date -- Author
%>

<tr ALIGN="CENTER"> 
<td valign="top" colspan="2" bgcolor="<%=myApplicationColor%>">
<FONT FACE="Arial, Helvetica, sans-serif"  Size="1" COLOR="<%=myApplicationTextColor%>"> 
&nbsp; <% If myAction="Update" then %> <%=myDate_Display(myWeb_Date_Update,2)%> -- <%=myWeb_Author_Update%> <% end if%> </font>
</td>
</tr>
</table>
</form>

<%
' ADMINISTRATION - EveryBody Can Add, Delete or Modify in this Version
%> 

<% 
If myAction="Update" then 
	%>
	<TABLE BORDER="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="3" CELLSPACING="0"> 
	<TR>
	<TD WIDTH="1%">&nbsp;</TD>
	<TD>
	<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><A HREF="__Web_Modification.asp?WebDirectory_ID=<%=myWebDirectory_ID%>&Action=New"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> <%=myMessage_Add%>&nbsp;<%=myMessage_Web%></font></A></FONT>
	</TD>
	</TR>

	<TR>
	<TD WIDTH="1%">&nbsp;
	
	</TD>
	<TD>
	<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><A HREF="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Web_modification.asp?Action=Delete&amp;Web_ID=<%=myWeb_ID%>';"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Delete%></font></A></FONT> 
	</TD>
	</TR>
	</TABLE>
	<% 
End If
%>

</td>
</TR>
</TABLE>

<%
' /ADMINISTRATION
%>


<!-- #include file="_borders/down.asp" --> 

<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.														'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0">
<TR ALIGN="RIGHT"><TD>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> 
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</FONT>
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


<html></html>
<html><script language="JavaScript"></script></html>