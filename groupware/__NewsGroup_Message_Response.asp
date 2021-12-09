<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   _ Over Apps _ http://www.overapps.com
'
' This program "__NewsGroup_messages_Response.asp" is free software; 
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
' Does n't Work with PWS ??????
%>

<%
' ------------------------------------------------------------
' Name			: __NewsGroup_messages_Response.asp
' Path		   	: /
' Version 		: 1.15.0
' Description 	: Message Response
' By 			: Pierre Rouarch	
' Company		: OverApps
' Date 			: Ddecember 10, 2001 
'Contributor : Dania Tcherkezoff
' ------------------------------------------------------------

Dim myPage
myPage = "__NewsGroups_Messages_Response.asp"



Dim myPage_Application
myPage_Application="Newsgroups"
	
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

Function myAnsi2Html(ByVal mystring)
 
  mystring = trim(mystring)
  if isNull(mystring) then mystring=""
  if Len(mystring)>0 then
   mystring=server.HTMLEncode(mystring)
   mystring = Replace(mystring,chr(10),"<br>")
  end if  
  myAnsi2Html = mystring
 
 End Function

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

Dim mySQL_Select_tb_Newsgroups_Messages, mySQL_Delete_tb_Newsgroups_Messages, mySQL_Update_tb_Newsgroups_Messages, mySQL_Insert_tb_NewsGroups_Messages, mySet_tb_Newsgroups_Messages 



Dim  myKey

Dim mySQL_Select_tb_Newsgroups, mySet_tb_Newsgroups 

Dim myNewsGroup_ID, myNewsGroup_Moderator_ID, myNewsGroup_Moderator_Site_ID, myNewsGroup_Name

Dim myNewsGroup_Parent_Message_ID,  myNewsGroup_Parent_Message_Author,  myNewsGroup_Parent_Message_Date,  myNewsGroup_Parent_Message_Title, myNewsGroup_Parent_Message,  myNewsGroup_Parent_Message_Thread,  myNewsGroup_Parent_Message_Thread_Date

 
Dim myNewsGroup_Message_Title, myNewsGroup_Message, myNewsGroup_Message_Thread, myNewsGroup_Message_Thread_Date


Dim myURL

Dim mySubstring, myResult

myNewsGroup_ID=request("Newsgroup_ID")
if len(myNewsGroup_ID)=0 then 
	myNewsgroup_ID=1
end if
myKey	= Request("KEY")
myNewsGroup_Parent_Message_ID=Request("KEY")



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Delete Message
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Request.QueryString("Action")="Delete" then
 if myUser_Type_Id = 1 Then
	set myConnection = CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

 	mySQL_Delete_tb_Newsgroups_Messages = "DELETE FROM tb_NewsGroups_Messages WHERE NewsGroup_Message_ID=" & Request.QueryString("KEY") 
	myConnection.Execute(mySQL_Delete_tb_Newsgroups_Messages)
	myConnection.Close
  end if	
 Response.redirect "__Newsgroup_Messages_List.asp?NewsGroup_ID=" & myNewsGroup_ID
 
End IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clean Message
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Request.QueryString("Action")="Blank" then
 if myUser_Type_ID = 1 Then
	set myConnection = CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

	mySQL_Update_tb_Newsgroups_Messages="UPDATE tb_newsgroups_messages SET NewsGroup_Message_Title='"&myMessage_Blank_by_The_Moderator&"', NewsGroup_Message='"&myMessage_Blank_by_The_Moderator&"' WHERE tb_newsgroups_messages.NewsGroup_Message_ID=" & Request.QueryString("KEY")


	myConnection.Execute(mySQL_Update_tb_Newsgroups_Messages)
	myConnection.Close

	myURL="__Newsgroup_Messages_List.asp?NewsGroup_ID="&myNewsGroup_ID
  end if	
 Response.redirect(myURL)
End IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






if Request.form("Validation")=myMessage_Go then

	' Get Entries
	myNewsGroup_Id=request.form("Newsgroup_ID")
	myNewsgroup_Message_Title=Replace(Request.Form("NewsGroup_Message_Title"),"'","''")
	myNewsgroup_Message= Replace(Request.Form("NewsGroup_Message"),"'","''")
	myNewsGroup_Message_Thread=request.form("Newsgroup_Message_Thread")
	myNewsGroup_Message_Thread_Date=request.form("Newsgroup_Message_Thread_Date")

	' Form Validation
	Call myFormSetEntriesInString

	' Test Entries :

	myFormCheckEntry null, "NewsGroup_Message_Title",true,null,null,0,100
	myFormCheckEntry null, "NewsGroup_Message",False,null,null,0,10000

	if not myform_entry_error then

		If Request.QueryString("Action")="Response" then
	
			set myConnection = Server.CreateObject("ADODB.Connection")
			myConnection.Open myConnection_String

			mySQL_Select_tb_Newsgroups_Messages = "SELECT * FROM tb_Newsgroups_Messages"
			Set mySet_tb_Newsgroups_Messages = server.createobject("adodb.recordset")
			mySet_tb_Newsgroups_Messages.open mySQL_Select_tb_Newsgroups_Messages, myConnection, 3, 3
			mySet_tb_Newsgroups_Messages.AddNew
	
			mySet_tb_Newsgroups_Messages.fields("Site_ID")=mySite_ID
			mySet_tb_Newsgroups_Messages.fields("Member_ID")=myUser_ID
			mySet_tb_Newsgroups_Messages.fields("NewsGroup_ID")=myNewsGroup_ID
			mySet_tb_Newsgroups_Messages.fields("NewsGroup_Message_Date")=myDate_Now()
			mySet_tb_Newsgroups_Messages.fields("NewsGroup_Message_Author")=myUser_Pseudo
			mySet_tb_Newsgroups_Messages.fields("NewsGroup_Message_Title")=myNewsGroup_Message_Title
			mySet_tb_Newsgroups_Messages.fields("NewsGroup_Message")=myNewsGroup_Message

			mySet_tb_Newsgroups_Messages.fields("NewsGroup_Message_Thread") = myNewsGroup_Message_Thread
			mySet_tb_Newsgroups_Messages.fields("NewsGroup_Message_Thread_Date") = myNewsGroup_Message_Thread_Date


	
			mySet_tb_Newsgroups_Messages.Update
			' Close Recordset 
			mySet_tb_Newsgroups_Messages.close
			Set mySet_tb_Newsgroups_Messages = Nothing
			' Close Connection
			myConnection.close
			set myConnection = Nothing
		Response.redirect "__Newsgroup_Messages_List.asp?NewsGroup_ID="&myNewsGroup_ID

		end if   ' Response

	end if

end if
		


%>
<html>

<head>
<title><%=mySite_Name%> - NewsGroup Message Response</title>
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


 <TD WIDTH="<%=myLeft_Width%>"> <!-- #include file="_borders/Left.asp" --></td>

<%
' CENTER APPLICATION
%>

<td  bgcolor="<%=mybgcolor%>"> <% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




set myConnection = CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Newsgroups_Messages="SELECT tb_newsgroups.newsgroup_ID, tb_Newsgroups.Member_ID as NewsGroup_Moderator_ID, tb_Newsgroups.Site_ID as NewsGroup_Moderator_Site_ID, tb_NewsGroups.NewsGroup_Name, NewsGroup_Message_ID, NewsGroup_Name, NewsGroup_Message_Date, NewsGroup_Message_Author, NewsGroup_Message_Title, NewsGroup_Message, NewsGroup_Message_Thread, NewsGroup_Message_Thread_Date  FROM tb_newsgroups INNER JOIN tb_newsgroups_messages ON tb_newsgroups.NewsGroup_ID=tb_newsgroups_messages.NewsGroup_ID WHERE tb_newsgroups_messages.NewsGroup_Message_ID="&myNewsgroup_Parent_Message_ID



set mySet_tb_Newsgroups_Messages= myConnection.Execute(mySQL_Select_tb_Newsgroups_Messages)
mySet_tb_Newsgroups_Messages.movefirst
myNewsGroup_Name=mySet_tb_Newsgroups_Messages("NewsGroup_Name")
myNewsGroup_Moderator_ID=mySet_tb_Newsgroups_Messages("NewsGroup_Moderator_ID")
myNewsGroup_Moderator_Site_ID=mySet_tb_Newsgroups_Messages("NewsGroup_Moderator_Site_ID")
myNewsGroup_Parent_Message_ID=mySet_tb_Newsgroups_Messages("NewsGroup_Message_ID")
myNewsGroup_Parent_Message_Date=myDate_Display(mySet_tb_Newsgroups_Messages("NewsGroup_Message_Date"),2)
myNewsGroup_Parent_Message_Author=mySet_tb_Newsgroups_Messages("NewsGroup_Message_Author")
myNewsGroup_Parent_Message_title= myAnsi2HTML(mySet_tb_Newsgroups_Messages("NewsGroup_Message_Title"))
myNewsGroup_Parent_Message=myAnsi2HTML( mySet_tb_Newsgroups_Messages("NewsGroup_Message"))
myNewsGroup_Parent_Message_Thread= mySet_tb_Newsgroups_Messages("NewsGroup_Message_Thread")
myNewsGroup_Parent_Message_Thread_Date= mySet_tb_Newsgroups_Messages("NewsGroup_Message_Thread_Date")

mySet_tb_Newsgroups_Messages.Close
Set mySet_tb_Newsgroups_Messages=Nothing
myConnection.Close
set myConnection = Nothing


%>
 <TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" BGCOLOR="<%=myApplicationColor%>" cellspacing="1" cellpadding="5" >
 <TR><TD><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b>
<%=myApplication_Title%></b></font></TD></TR> </TABLE>
<table WIDTH="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" border="0" cellspacing="1" cellpadding="5"> 

<tr>
 <td bgcolor="<%= myBorderColor %>" align=right width=40%>
 <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>">
    <strong><%=myMessage_Title%></strong>&nbsp;
 </font>
</td>
<td bbgcolor="<%= myBGColor %>">
  <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBGTextColor%>">
    <b> <%=myNewsGroup_Parent_Message_Title%></b>
  </font></td> 
</tr>

<tr>
 <td bgcolor="<%= myBorderColor %>" align=right width=40%>
 <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>">
    <strong><%=myMessage_Date%></strong>&nbsp;
 </font>
</td>
<td bbgcolor="<%= myBGColor %>">
  <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBGTextColor%>">
     <%=myNewsGroup_Parent_Message_Date%>
  </font></td> 
</tr>

<tr>
 <td bgcolor="<%= myBorderColor %>" align=right width=40%>
 <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>">
    <strong><%=myMessage_Author%></strong>&nbsp;
 </font>
</td>
<td bbgcolor="<%= myBGColor %>">
  <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBGTextColor%>">
     <%=myNewsGroup_Parent_Message_Author%>
  </font></td> 
</tr>


<tr>
 <td bgcolor="<%= myBorderColor %>" align=right width=40%>
 <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>">
    <strong><%=myMessage_Message%></strong>&nbsp;
 </font>
</td>
<td bbgcolor="<%= myBGColor %>">
  <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBGTextColor%>">
    <%=replace(myNewsGroup_Parent_Message& " ", chr(10), "<br>")%>
  </font></td> 
</tr>


<%	IF myNewsGroup_Moderator_ID=myUser_ID OR (myNewsGroup_Moderator_Site_ID=mySite_ID and myUser_type_ID= 1)  then %> 
<tr>
 <td bgcolor="<%= myBGColor %>" align=left width=40% colspan=2>
 <FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>">
<a href="__NewsGroup_message_response.asp?NewsGroup_ID=<%=myNewsGroup_ID%>&amp;Key=<%=myNewsGroup_Parent_Message_ID%>&amp;Action=Blank" onclick="return(confirm('<%=myMessage_Clean%> ?'))">
<%=myMessage_Clean%></a>, 
<a href="__NewsGroup_message_response.asp?NewsGroup_ID=<%=myNewsGroup_ID%>&amp;Key=<%=myNewsGroup_Parent_Message_ID%>&amp;Action=Delete" onclick="return(confirm('<%=myMessage_Delete%> ?'))">
<%=myMessage_Delete%></a> 
</font>
</td></tr>
<%end if%>
</table>

 
 


<table width="<%=myApplication_Width%>" bgColor="<%=myBGColor%>" border="0" cellspacing="1" cellpadding="5" > 
<tr>
<td colspan="2" BGCOLOR="<%=myApplicationColor%>">
<font face="Arial, Helvetica, sans-serif" size="3" color="<%=myApplicationTextColor%>"><b><%=myMessage_Response%></b></font>
</td>
</tr>

<tr> <td bgcolor="<%= myBorderColor %>" align=right> 
<form method="POST" action="__NewsGroup_message_response.asp?NewsGroup_ID=<%=myNewsGroup_ID%>&amp;Action=Response" target="_top" > 
<input type="hidden" name="NewsGroup_ID" value="<%=myNewsGroup_ID%>">
<input type="hidden" name="KEY" value="<%=myKey%>">
<input type="hidden" name="NewsGroup_Parent_Message_Thread_Date" value="<%=myNewsGroup_Parent_Message_Thread_Date%>"> 
<input type="hidden" name="NewsGroup_Message_Thread" value="<%=myNewsGroup_Parent_Message_Thread%>/<%= myNewsGroup_Parent_Message_ID %>"> 
<input type="hidden" name="NewsGroup_Message_Thread_Date" value="<%=myNewsGroup_Parent_Message_Thread_Date%>">

<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>"><strong><%=myMessage_Title%>* </strong>  
<%=myFormGetErrMsg("NewsGroup_Message_Title")%> </FONT> 
<td bgcolor="<%= myBGColor %>"> <input type="text" name="NewsGroup_Message_Title"  size="53" maxlength="255" value="Re: <%= myNewsGroup_Parent_Message_Title%>"> 
</td></tr>

<tr> <td bgcolor="<%= myBorderColor %>" align=right><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>"><strong><%=myMessage_Message%>&nbsp;</strong></font></td>
<td bbgcolor="<%= myBGColor %>"><textarea name="NewsGroup_Message"  rows="10" cols="40"><%=myNewsGroup_Message%></textarea></td> 
</tr>

<tr> <td bgcolor="<%= myBorderColor %>" align=right><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>"><strong><input type="SUBMIT" Name="Validation" value="<%=myMessage_Go%>"></strong>&nbsp;
<br>* :  <%= myMessage_Required %>&nbsp;
</font></td>
<td bbgcolor="<%= myBGColor %>">&nbsp;</td> 
</tr>

<tr>
<td colspan="2" BGCOLOR="<%=myApplicationColor%>">
&nbsp;
</td>
</tr>

<tr>
<td colspan="2" BGCOLOR="<%=myBGColor%>">
&nbsp;
</td>
</tr>





</table>







</table></td></TR> </TABLE>

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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</FONT>
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