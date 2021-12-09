<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   -+ OverApps +- http://www.overapps.com
'
' This program "__New_modification.asp" is free software; 
' you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License.
'
' This program "__New_modification.asp" is distributed in the hope 
' that it will be useful,
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
' Doen't Work with PWS
%>


<%
' ------------------------------------------------------------
' 
' Name			: __New_modification.asp
' Päth		    : /
' Version		: 1.15.0
' Description 	: Add/Modify/Delete Article
' by 			: Pierre Rouarch	
' Company		: Overapps
' Date			: December,10, 2001 
'
' Contributions : 	Christophe Humbert, Jean-Luc Lesueur, Dania Tcherkezoff
'
'
' Modify by		: 
' Company		:
' Date			:
' ------------------------------------------------------------

Dim myPage
myPage = "__New_modification.asp"

Dim myPage_Application
myPage_Application="News"
	
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

Dim myMember_ID

Dim myNew_Site_ID, myNew_Member_ID, myNew_NewsWire_ID, myNew_ID, myNew_Title,  myNew_Description_Short, myNew_Description_Long,  myNew_Date, myNew_Date_Update, myNew_Author_Update  

Dim   myNewsWire_ID, myNewsWire_Name


Dim myAction, myTitle
					
Dim mySQL_Select_tb_News, mySQL_Delete_tb_News, mySQL_Insert_tb_News, mySQL_Update_tb_News, mySet_tb_News

Dim mySQL_Select_tb_NewsWires, mySet_tb_NewsWires


' Get Parameters
myAction = Request("Action")

myNew_NewsWire_ID = Request("NewsWire_ID")
if len(myNew_NewsWire_ID)=0 then
		myNew_NewsWire_ID=1
end if
	
myNew_ID = Request("New_ID")
if len(myNew_ID)=0 then
	myAction = "New"
else 
	if myAction<>"Delete" then 
		myAction="Update"
	end if	
end if

myNew_Date = Request("New_Date")
if len(myNew_Date)=0 then
	myNew_Date = myDate_Now()
end if


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Delete
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


if myAction = "Delete" then

	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

	myNew_ID  = Request.QueryString("New_ID") 
	mySQL_Delete_tb_News = "DELETE FROM tb_News WHERE Site_ID = "&Session("Site_ID")&" AND New_ID = " & myNew_ID

	myConnection.Execute(mySQL_Delete_tb_News)
	myConnection.Close
	set myConnection = Nothing
	' And go Back
		Response.Redirect("__News_List.asp")

end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


if Request.form("Validation")=myMessage_Go then


	' Get Entries
	myNew_ID	      = Request.Form("New_ID")
	myNew_Title	= Replace(Request.form("New_Title"),"'"," ")
	myNew_Description_Short   = Replace(Request.Form("New_Description_Short"),"'"," ")
	myNew_Description_Long   = Replace(Request.Form("New_Description_Long"),"'"," ")


	' Test Entries
	Call myFormSetEntriesInString

	myFormCheckEntry null, "New_Title",true,null,null,0,255
	myFormCheckEntry null, "New_Description_Short",true,null,null,0,255

	if not myform_entry_error then

		' connexion à la base 
		set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String

		myNew_Author_Update   = myUser_Pseudo
		myNew_Date_Update     = myDate_Now()




		' INSERT
		if myAction = "New" then


			mySQL_Select_tb_News = "SELECT * FROM tb_News"
			Set mySet_tb_News = server.createobject("adodb.recordset")
			mySet_tb_News.open mySQL_Select_tb_News, myConnection, 3, 3
			mySet_tb_News.AddNew

			mySet_tb_News.fields("Site_ID")=mySite_ID
			mySet_tb_News.fields("Member_ID")=myUser_ID
			mySet_tb_News.fields("NewsWire_ID")=myNew_NewsWire_ID
			mySet_tb_News.fields("New_Title")=myNew_Title
			mySet_tb_News.fields("New_Description_Short")=myNew_Description_Short
			mySet_tb_News.fields("New_Description_Long")=myNew_Description_Long
			mySet_tb_News.fields("New_Date")=myNew_Date
			mySet_tb_News.fields("New_Date_Update")=myNew_Date_Update
			mySet_tb_News.fields("New_Author_Update")=myNew_Author_Update
	
			mySet_tb_News.Update
			' Close Recordset 
			mySet_tb_News.close
			Set mySet_tb_News = Nothing

			
		' UPDATE
		elseif myAction = "Update" then

		mySQL_Select_tb_News = "SELECT * FROM tb_News WHERE New_ID="&myNew_ID
			Set mySet_tb_News = server.createobject("adodb.recordset")
			mySet_tb_News.open mySQL_Select_tb_News, myConnection, 3, 3


			mySet_tb_News.fields("New_Title")=myNew_Title
			mySet_tb_News.fields("New_Description_Short")=myNew_Description_Short
			mySet_tb_News.fields("New_Description_Long")=myNew_Description_Long
			mySet_tb_News.fields("New_Date_Update")=myNew_Date_Update
			mySet_tb_News.fields("New_Author_Update")=myNew_Author_Update
	
			mySet_tb_News.Update
			' Close Recordset 
			mySet_tb_News.close
			Set mySet_tb_News = Nothing


		end if


		' Close Connection
		myConnection.close
		set myConnection = nothing	
		' And Go Back
		Response.Redirect("__News_List.asp")
		

	
	end if 

end if 


%>
<html>

<head>
<title><%=mySite_Name%> - Add/Modify/Delete Article </title>

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

<TD WIDTH="<%=myLeft_Width%>">
<!-- #include file="_borders/Left.asp" -->
</td>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form															'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myAction = "Update" and Request.form("Validation")<>myMessage_Go then

	' DB Connexion 
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

	mySQL_Select_tb_News = "SELECT   *,New_Description_Long,New_Author_Update,New_Date_Update FROM tb_News INNER JOIN tb_NewsWires_Sites ON tb_News.NewsWire_ID=tb_Newswires_Sites.NewsWire_ID WHERE tb_NewsWires_Sites.Site_ID="&mySite_ID&" AND NEW_ID="&myNew_ID 


	set mySet_tb_News = myConnection.Execute(mySQL_Select_tb_News)

	' If Nothing Go Back
	if mySet_tb_News.eof then
		' Close Reordset
		mySet_tb_News.close
		Set mySet_tb_News=nothing
		' Close Connection
		myConnection.close
		set myConnection = nothing
		Response.Redirect("__News_List.asp")
	end if

	' Read informations
	myNew_Site_ID = mySet_tb_News("Site_ID")
	myNew_Member_ID = mySet_tb_News("Member_ID")
 	myNewsWire_ID = mySet_tb_News("NewsWire_ID")
	myNew_ID = mySet_tb_News("New_ID")
	myNew_Date=mySet_tb_News("New_Date")
	myNew_Title  = mySet_tb_News("New_Title")
	myNew_Description_Short  = mySet_tb_News("New_Description_Short")
	myNew_Description_Long  = mySet_tb_News("New_Description_Long")
	myNew_Author_Update   = mySet_tb_News("New_Author_Update")				
	myNew_Date_Update     = mySet_tb_News("New_Date_Update")

	

	' Close Reordset
	mySet_tb_News.close
	Set mySet_tb_News=nothing
	' Close Connection
	myConnection.close
	set myConnection = nothing

end if



%> 

<td valign="top" BGCOLOR="<%=myBGColor%>">
<form method="POST" action="<%=myPage%>" name="myForm" >
 
<table WIDTH="<%=myApplication_Width%>" border="0" cellpadding="5" cellspacing="1">

<% 
' Application Title and hidden fields
%>


<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="4" COLOR="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></FONT>
<INPUT TYPE="hidden" NAME="Action" VALUE="<%=myAction%>"> 
<INPUT TYPE="hidden" NAME="NewsWire_ID" VALUE="<%=myNewsWire_ID%>"> 
<INPUT TYPE="hidden" NAME="New_ID" VALUE="<%=myNew_ID%>"> 
<INPUT TYPE="hidden" NAME="New_Date" VALUE="<%=myNew_Date%>"> 
</TD>
</TR> 







<%
' Date
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" COLOR="<%=myBorderTextColor%>"><b><%=myMessage_Date%> 
</b></font>
</td>
<td align="left"  valign="top">
<font size="2" face="Arial, Helvetica, sans-serif"><b><%=myDate_Display(myNew_Date,2)%></b></font>
</td>
</tr>


<%
' New Title
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Title%>*<BR><%=myFormGetErrMsg("New_Title")%></b></FONT>
</td>
<td align="left"   valign="top"> 
<input type="text" size="65" name="New_Title" value="<%=myNew_Title%>"> 
</td>
</tr>

<%
' Presentation
%>


<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Presentation%>*<BR><%=myFormGetErrMsg("New_Description_Short")%></b></FONT>
</td>
<td align="left"  valign="top">
<input type="text" size="65" name="New_Description_Short" value="<%=myNew_Description_Short%>">
</td>
</tr>



<%
' Article
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Article%> 
</b></font>
</td>
<td valign="top" align="left">
<TEXTAREA NAME="New_Description_Long" COLS="60" ROWS="10"><%=myNew_Description_Long%></TEXTAREA>
</td>
</tr>

<%
' Validation
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">&nbsp;

</td>
<td valign="top"  align="left"> 
<input type="submit" value="<%=myMessage_Go%>" name="Validation">
</td>
</tr>

<%
' date --Author
%>

<tr> 
<td align="Center" valign="top" colspan="2" bgcolor="<%=myApplicationColor%>">
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myApplicationTextColor%>"> 
<%=myDate_Display(myNew_Date_Update,2)%> -- <%=myNew_Author_Update%></font>
</td>
</tr>

</table>
</form>



<%
' ADMINISTRATION - Everybody Can Add/Delete/modify 
%> 

<% If myAction="Update" then %> 

<TABLE BORDER="0" WIDTH="<%=myApplication_Width%>" CELLPADDING="3" CELLSPACING="0"> 
<TR>
<TD>
&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><A HREF="__New_Modification.asp?NewsWire_ID=<%=myNew_NewsWire_ID%>"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Add%>&nbsp;<%=myMessage_Article%></font></A> 
</FONT>
</TD>
</TR>
<TR>
<TD>
&nbsp;<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><A HREF="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__New_modification.asp?Action=Delete&amp;New_ID=<%=myNew_ID%>';"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myMessage_Delete%></font></A></FONT>
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
' license's compliances.							'
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
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 

</body>
</html>


<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>