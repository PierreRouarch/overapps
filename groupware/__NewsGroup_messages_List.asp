<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   + OverApps + http://www.overapps.com
'
' This program "__NewsGroup_Messages_List.asp?" is free software; you can 
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
'
'-----------------------------------------------------------------------------
%>
<% 	Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Does n't Work with PWS ?????
%>

<%
' ------------------------------------------------------------
' Nom 			 : __Newsgroup_messages_List.asp
' Path 		      : /
' Description : List and new Message
' by			    : Pierre Rouarch	
' Company	  : OverApps
' Date			   : December 10, 2001
' Version        : 1.15.0
' Contributor  : Franck Couvy , Dania Tcherkezoff 
' Modify by		:
' Company	   :	
' Date			   :
' ------------------------------------------------------------


Dim myPage, myURL
myPage = "__NewsGroup_Messages_List.asp?NewsGroup_ID=1"

' This page is NOT an open page
if len(session("Site_ID"))=0 then
	mySite_ID=1
	session("Site_ID")=mySite_ID
end if

if len(session("User_ID"))=0 or session("User_ID")=0 then
	myURL="__Identification_site.asp?Reloc="&myPage
	response.Redirect(myURL)
end if 



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



Dim mySQL_Select_tb_Newsgroups_Messages, mySQL_Insert_tb_Newsgroups_Messages, mySQL_Delete_tb_Newsgroups_Messages, mySet_tb_Newsgroups_Messages, mySQL_Select_tb_Newsgroups, mySet_tb_Newsgroups

Dim myNewsGroup_ID, myNewsGroup_Name, myNewsGroup_Message_ID, myNewsGroup_Message_Title, myNewsGroup_Message, myNewsGroup_Message_Author, myNewsGroup_Message_Date, myNewsGroup_Message_Thread

Dim  myOrder,  myOrder2

Dim mySubstring, myResult
'--------------------------- by couvy 30/11/2001-2002 
'  new vars
dim myNewsGroup_Message_Text, mystring, myKey, myGifLink
myKey = Request.QueryString("key")
'---------------------------

myNewsGroup_ID=request("Newsgroup_ID")
if len(myNewsGroup_ID)=0 then 
	'Force to NewsGroup 1 in the Mono Newsgroup Version 
	myNewsGroup_ID=1
end if 

' For Presentation 

Function Tabulation(mystring)
	mysubstring=mystring
	myresult=""
	while instr(mysubstring,"/") > 0 
		myresult=myresult & "&nbsp;&nbsp;&nbsp;&nbsp;" 
		mysubstring=right(mysubstring,Len(mysubstring)-instr(mysubstring,"/"))
	wend
	Tabulation=myresult
end function

' Formating ansi string to html output		
Function myAnsi2Html(ByVal mystring)
  mystring = trim(mystring)
  if isNull(mystring) then mystring=""
  if Len(mystring)>0 then
   mystring=server.HTMLEncode(mystring)
   mystring = Replace(mystring,chr(10),"<br>")
  end if  
  myAnsi2Html = mystring
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FORM VALIDATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	' Get Entries


	myNewsgroup_Message_Title=Request.Form("NewsGroup_Message_Title")
	myNewsgroup_Message=Request.Form("NewsGroup_Message")


if Request.form("Validation")=myMessage_Go then

	' Form Validation
	Call myFormSetEntriesInString

	' Test  Entries :

	myFormCheckEntry null, "NewsGroup_Message_Title",true,null,null,0,100
	myFormCheckEntry null, "NewsGroup_Message",false,null,null,0,10000

	if not myform_entry_error then

		' Write a new Message in DB 
	
			myNewsGroup_Message_Thread="0"
			' DB Connection
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
			mySet_tb_Newsgroups_Messages.fields("NewsGroup_Message_Thread_Date") = myDate_Now()

			mySet_tb_Newsgroups_Messages.Update
			' Close Recordset 
			mySet_tb_Newsgroups_Messages.close
			Set mySet_tb_Newsgroups_Messages = Nothing

			myNewsGroup_Message_Title=""
			myNewsGroup_Message=""
			myConnection.Close
			set myConnection = Nothing

	end if ' No Entry Error

end if	' Validation 

%>

<html>

<head>
<title><%=mySite_Name%></title>
</head>

<BODY BackGround="<%=myBGImage%>" bgColor="<%=myBGColor%>" Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%>
<!-- #include file="_borders/Top.asp" -->
<TABLE WIDTH="<%=myGlobal_Width%>" bgcolor="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0" > 
<%
' CENTER 
%> 

<TR VALIGN="TOP">

<%
' CENTER LEFT
%>

<TD WIDTH="<%=myLeft_Width%>"> <!-- #include file="_borders/Left.asp" --></td>

<%
' CENTER APPLICATION
%>

<td bgcolor="<%=mybgcolor%>"> <table width="<%=myApplication_Width%>" bgcolor="<%=mybgcolor%>" border="0" cellspacing="0" cellpadding="0" > 
<tr> <td> <% 	
' db Access
	set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String

%>

<TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" BGCOLOR="<%=myApplicationColor%>"><TR><TD><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b> </font></TD></TR> </TABLE>

<%	

myOrder2=  " [NewsGroup_Message_Thread]" & "&'/'&" & "[NewsGroup_Message_ID] "



mySQL_Select_tb_Newsgroups_Messages="SELECT  * ,tb_newsgroups_messages.newsgroup_Message FROM tb_newsgroups_messages WHERE NewsGroup_ID ="&myNewsGroup_ID&" ORDER BY NewsGroup_Message_thread_Date DESC"', " &myOrder2

'response.write mySQL_Select_tb_Newsgroups_Messages
'response.end


%> </font> <table BORDER="0" BGCOLOR="<%=myBGColor%>" CELLPADDING="0" CELLSPACING="0" width="<%=myApplication_Width%>"> 
<tr> 
<th width=20% bgcolor="<%=myBorderColor%>"><font face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><%=myMessage_Date%></font></th>
<th width=20% bgcolor="<%=myBorderColor%>"> <font face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><%=myMessage_Member%></font></th>
<th width=60% bgcolor="<%=myBorderColor%>"><font face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><%=myMessage_Title%></font></th></tr> 
<%	
' response.write(myQuery)
Set mySet_tb_Newsgroups_Messages = server.createobject("adodb.recordset")
set mySet_tb_Newsgroups_Messages= myConnection.Execute(mySQL_Select_tb_Newsgroups_Messages)


	do while not mySet_tb_Newsgroups_Messages.eof
		myNewsGroup_Message_ID     = mySet_tb_Newsgroups_Messages("NewsGroup_Message_ID")
		myNewsGroup_Message_Date   = myDate_Display(mySet_tb_Newsgroups_Messages("NewsGroup_Message_Date"),2) 
		myNewsGroup_Message_Author = mySet_tb_Newsgroups_Messages("NewsGroup_Message_Author")
		myNewsGroup_Message_Thread = mySet_tb_Newsgroups_Messages("NewsGroup_Message_Thread")
		myNewsGroup_Message_Title  = mySet_tb_Newsgroups_Messages("NewsGroup_Message_Title")
        
       
%> 
<tr> 
<td align="left" style='font-family: Helvetica,Arial;font-weight:bold; font-size:.8em'> 
&nbsp;<%=myNewsGroup_Message_Date%> 
</td>

<td align="left" style='font-family: Helvetica,Arial;font-weight:bold; font-size:1em'> 
&nbsp;&nbsp;<%=myNewsGroup_Message_Author%> 
</td>

<%
'--------------------------- by couvy 30/11/2001-2002 
' link for displaying the text of the message
' adding a new param  to the application: key
%>

<%
'--------------------------- by couvy 30/11/2001-2002 
'display the text of the messsage
'create the gif and link to display the message
myNewsGroup_Message_Text = ""
myGifLink = "<a href='__NewsGroup_Messages_list.asp?key=" & myNewsGroup_Message_ID & "'><img src='Images/plus.gif' width=13 height=12 border=0></a>"

if myKey<>"" and (cstr(myNewsGroup_Message_ID) = myKey or instr(myNewsGroup_Message_Thread,"/"& myKey)>0) then 
	
	myNewsGroup_Message_Text   = myAnsi2HTML( mySet_tb_Newsgroups_Messages("NewsGroup_Message")) 
	myGifLink = "<a href='__NewsGroup_Messages_list.asp'> <img src='Images/leaf.gif' width=13 height=12 border=0></A>"
end if
%>

<td align="left" style='font-family: Helvetica,Arial;font-weight:bold; font-size:1em'> 

<%=myGifLink %>&nbsp;
<%=tabulation(myNewsGroup_Message_Thread)%> 
<a href="__NewsGroup_Message_Response.asp?key=<%=myNewsGroup_Message_ID%>&amp;NewsGroup_ID=<%=myNewsGroup_ID%>"><%=MyNewsGroup_Message_Title%>
</a>
</td>

<tr ><td ></td><td></td>

<td align="left" style='font-family: Helvetica,Arial; font-size:0.8em'> 
<%=myNewsGroup_Message_Text%>
</td>
</tr>
</td>
</tr> 


<%	
	mySet_tb_Newsgroups_Messages.movenext
	loop 


' CLose Recordset
mySet_tb_Newsgroups_Messages.close
Set mySet_tb_Newsgroups_Messages=nothing
' Close Connection
myConnection.Close
set myConnection = Nothing

%> 



</table></td></tr>

<tr><td bgcolor="<%= myBGColor%>">
<br><br><br><br>

</td></tr>

 </table><font face="Arial, Helvetica, sans-serif"> 
</font> <font face="Arial, Helvetica, sans-serif">
<table width="<%=myApplication_Width%>" bgColor="<%=myBGColor%>" border="0" cellspacing="1" cellpadding="5" > 
<tr>
<td colspan="2" BGCOLOR="<%=myApplicationColor%>">
<font face="Arial, Helvetica, sans-serif" size="3" color="<%=myApplicationTextColor%>"><b><%= myMessage_New %>&nbsp;<%= myMessage_Message %></b></font>
</td>
</tr>

<tr> <td bgcolor="<%= myBorderColor %>" align=right> 
<form method="POST" action="__NewsGroup_messages_List.asp" target="_top" name="theForm" > 
<input type="hidden" name="NewsGroup_ID" value="<%=myNewsGroup_ID%>">
<input type="hidden" name="NewsGroup_Message_Thread" value="0">

<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>"><strong><%=myMessage_Title%>* </strong>  
<%=myFormGetErrMsg("NewsGroup_Message_Title")%> </FONT> 
<td bgcolor="<%= myBGColor %>"> <input type="text" name="NewsGroup_Message_Title"  size="53" maxlength="255"> 
</td></tr>

<tr> <td bgcolor="<%= myBorderColor %>" align=right><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>"><strong><%=myMessage_Message%>&nbsp;</strong></font></td>
<td bbgcolor="<%= myBGColor %>"><textarea name="NewsGroup_Message"  rows="10" cols="40"><%=myNewsGroup_Message%></textarea></td> 
</tr>

<tr> <td bgcolor="<%= myBorderColor %>" align=right><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%= myBorderTextColor %>"><strong><input type="SUBMIT" Name="Validation" value="<%=myMessage_Go%>">&nbsp;</strong>
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
</TABLE>
 
 </td></TR> </TABLE>


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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</FONT></A> & contributors</FONT>
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