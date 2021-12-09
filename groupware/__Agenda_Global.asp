<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Styles_Modification.asp" is free software; 
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
' at the bottom of the page with an active link from the name "OverApps"  
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
' Doesn't Work With PWS ????
%>
<%
' ------------------------------------------------------------
' Name 		       	: __Agenda_Global.asp
' Path   	      	: /
' Version 	    	: 1.15.0
' Description   	: To choose memeber to be included in global agenda
' Company	      	: OverApps
' Date			    : January 20, 2002
' Author            : Dania Tcherkezoff
' Modify by		    :
' Company		      :
' Date			      :
' ------------------------------------------------------------

Dim myPage
myPage = "__Agenda_Global.asp"
Dim myPage_Application
myPage_Application="Agenda"

Dim myDate_Agenda

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' INCLUDES 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<!-- #include file="_INCLUDE/Global_Parameters.asp" -->

<!-- #include file="_INCLUDE/Form_validation.asp" -->

<!-- #include file="_INCLUDE/DB_Environment.asp" -->

<!-- #include file="_INCLUDE/Environment_Tools.asp" -->

<!-- #include file="_INCLUDE/Files_Upload_Class.asp" -->

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




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET PARAMETERS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

myDate_Agenda = Request.QueryString("date_agenda")


'Open Connection, will serve in any case
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


'''''''''''''''''''''''''''''''''''''''''''''''''''
'UPDATING 
'''''''''''''''''''''''''''''''''''''''''''''''''''
Dim myAgenda_Date,i

if Request.Form("submit") = myMessage_Go Then


 for i =1 to 200
 
 session("Agenda_Global_" & i) = False
  
 next 

mySQL_Select_tb_Sites_Members ="  Select * from tb_Sites_Members  WHERE  Site_ID="&mySite_ID &" ORDER BY Member_Pseudo"
set mySet_tb_Sites_Members = myConnection.Execute(mySQL_Select_tb_Sites_Members)

  do while not mySet_tb_Sites_Members.eof
	If Request.Form(mySet_tb_Sites_Members("Member_Login") ) = "on" Then Session("Agenda_Global_" & mySet_tb_Sites_Members("Member_ID"))    = True 
    mySet_tb_Sites_Members.movenext
  loop

  'Close Conection
   myConnection.close
   set myConnection = Nothing 


  Response.Redirect "__Agenda_Global_Day.asp?Date_agenda=" & myDate_agenda

end if

%>
<html>

<head>
<title><%=mySite_Name%>  </title>
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

<TD WIDTH="<%=myLeft_Width%>"><!-- #include file="_borders/Left.asp" --></td>


<%
' CENTER APPLICATION
%> 


<%

%> 



<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 

<%
' APPLICATION TITLE
%>

<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myMessage_Global_Agenda%></b></font></TD></TR> 
</table>


<%
' FORM BOXS
%>


<table  cellpadding="3" cellspacing="1">
<form action="__Agenda_Global.asp?date_agenda=<%= myDate_Agenda %>" method=post >
<tr>
<td  bgcolor="<%= myBorderColor %>" align=right width=50%><font face="Arial, Helvetica, sans-serif" size=2 color="<% =myBorderTextColor %>"><b><%= myMessage_Global_Agenda_Members %></b></td>
<td >
<font face="Arial, Helvetica, sans-serif" size=2 color="<% =myBGTextColor%>">
<table border="0" cellpadding="5" cellspacing="0"> 
<%
Dim myCounter,myParticipant_ID,myParticipant_Pseudo

mySQL_Select_tb_Sites_Members ="  Select * from tb_Sites_Members  WHERE  Site_ID="&mySite_ID &" ORDER BY Member_Pseudo"
set mySet_tb_Sites_Members = myConnection.Execute(mySQL_Select_tb_Sites_Members)

' each 3 members CRLF
myCounter = 1
do while not mySet_tb_Sites_Members.eof
			
	myParticipant_ID     = mySet_tb_Sites_Members("Member_ID")
	myParticipant_Pseudo = mySet_tb_Sites_Members("Member_Pseudo")
	if myParticipant_Id <> myUSer_ID Then
	
		
	if myCounter = 1 then
			Response.Write "<tr bgcolor=#ffffff>"
	end if
		
	
%> 

	<td valign="top"  bgcolor="<%=myBGColor%>">
	<% If session("Agenda_Global_" & myParticipant_ID) Then %>
	  <input type=checkbox  name="<%=myParticipant_Pseudo%>" checked>
	<%else%>
	 	  <input type=checkbox name="<%=myParticipant_Pseudo%>" >
	<%end if%>	  
		
<Font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myParticipant_Pseudo%></font>


	</td>
	<%
	if myCounter = 3 then
		Response.Write "</tr>"
		myCounter = 1
	else
		myCounter= myCounter + 1
	end if
	end if		
	mySet_tb_Sites_Members.movenext
loop

if myCounter = 3 then
		Response.Write "<td>&nbsp; </td><td>&nbsp; </td></tr>"
elseif myCounter = 4 then
		Response.Write "<td>&nbsp; </td></tr>"
end if


	
%> 
</table>

</font>
</td>
</tr>


<tr>
<td bgcolor="<%= myBorderColor %>" align=right width=50%><font face="Arial, Helvetica, sans-serif" size=2 color="<% =myBorderTextColor %>"><input type=submit name=submit value="<%= myMessage_Go %>">&nbsp;</td>
<td>&nbsp;</td>
</tr>


</table>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>">&nbsp;</font></TD></TR> 
<TR><TD bgcolor="<%=myBGColor%>">
&nbsp;</td></tr>
</table>
</td>
</tr>
</table>
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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors</FONT>
</TD>
</TR>
</TABLE>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' End Copyright		                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 

</body>
</html>
