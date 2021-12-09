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
' Name 		       	: __Folders_Modification.asp
' Path   	      	: /
' Version 	    	: 1.15.0
' Description   	: Folder Information Modification
' By		        	: Dania TCHERKEZOFF
' Company	      	: OverApps
' Date			      : December 10, 2001 		
'
' Modify by		    :
' Company		      :
' Date			      :
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Folders_Modification.asp"

Dim myPage_Application
myPage_Application="Files"

Dim  mySQL_Select_tb_Meetings_Members, mySet_tb_Meetings_Members, mySQL_Insert_tb_Meetings_members,  mySQL_Delete_tb_Meetings_Members 

Dim URL, myCounter, myParticipant_ID, myMember_Name, myParticipant_Pseudo, myParticipant_Pseudo_Value

Dim myMembers_Public_Type_ID, myAction, myError, myFolder_Public, mySQL_Select_Folders_Access, mySet_tb_Folders_Access, mySet_tb_Sites_Member,mySQL_Select_tb_Sites_Member,mySQL_Select_tb_Folders_Access

Dim mySQL_Select_tb_Folders, mySet_tb_Folders, myFolder_ID,myFolder_Name, myFolder_Long_Description, myFolder_Short_Description, myFolder_Creator, MyFolder_Files_Number
	
Dim mySet_tb_Files, mySQL_Select_tb_Files, myFile_System_Object,myFolder_Modificator,myFolder_Modification_Date
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET APPLICATION TITLE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myApplication_Title = Get_Application_Title(myPage_Application)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CHECK IF THE USER CAN ENTER IN THIS APPLICATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

myApplication_Public_Type_ID = Get_Application_Public_Type_ID(myPage_Application)
if myApplication_Public_type_ID<myUser_type_ID then
	Response.redirect("__Quit.asp")
end if


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET PARAMETER
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myAction = Request.QueryString("myAction")

if len(myAction ) = 0 Then myAction = Request.Form("myAction") 
if len(myAction) = 0 Then myAction = "Modify"



myError = Request.QueryString("myError")

myFolder_ID = Request.QueryString("ID")

if len(myFolder_ID) = 0 Then myFolder_ID = Request.QueryString("Folder_ID")
if len(myFolder_ID) = 0 Then myFolder_ID = Request.Form("Folder_ID")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DELETE A FOLDER
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If myAction = "Delete" Then 

set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

'Selection of the folder to delete
 mySQL_Select_tb_Folders = "SELECT * FROM tb_folders WHERE Folder_ID=" & myFolder_ID 
 Set mySet_tb_Folders = server.createobject("adodb.recordset")
 set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders)
 if not mySet_tb_Folders.eof Then
  mySet_tb_Folders.MoveFirst
  if mySet_tb_Folders.fields("Folder_Creator_ID") = myUser_ID OR mySet_tb_Folders.fields("Folder_Responsible_ID") = myUser_ID OR myUser_Type_ID = 1 Then 
   myFolder_Name = mySet_tb_Folders.fields("Folder_Name")
   myConnection.Execute("Delete From tb_Folders where Folder_ID = " & myFolder_ID) 
   myConnection.Execute("Delete From tb_Folders_Access where Folder_ID = " & myFolder_ID)
   'DELETE ALL FILES IN FOLDER
   Set mySet_tb_Files = Server.CreateObject("ADODB.RecordSet")
   mySQL_Select_tb_Files = "Select * from tb_Files where File_Folder_ID = " & myFolder_ID
   Set mySet_tb_Files = MyConnection.Execute(mySQL_Select_tb_Files)
   set myFile_System_Object = Server.CreateObject("scripting.FileSystemObject")
    
	Do while not mySet_tb_Files.eof
	    myFile_System_Object.DeleteFile myShared_Files_Path &  myFolder_Name  & "\" & mySet_tb_Files("File_Name")
	   mySet_tb_Files.MoveNext
	loop
   myFile_System_Object.DeleteFolder myShared_Files_Path &  myFolder_Name
   
   myConnection.Execute("Delete from tb_Files where File_Folder_ID = " & myFolder_ID)
  end if    
 end if
 Response.Redirect "__Folders_List.asp"
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UPDATE A FOLDERS											
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction = "Update" Then

'Open Connection
 set myConnection = Server.CreateObject("ADODB.Connection")
 myConnection.Open myConnection_String

'Selection of the style to be updated
 mySQL_Select_tb_Folders = "SELECT * FROM tb_folders WHERE Folder_ID=" & myFolder_ID 
 Set mySet_tb_Folders = server.createobject("adodb.recordset")
 mySet_tb_Folders.open mySQL_Select_tb_Folders, myConnection, 3,3


If Request.form("myFolder_Public") = "on" Then 
 myFolder_Public = 1
else 
 myFolder_Public = 0 
end if


mySet_tb_Folders.fields("Folder_Public") = myFolder_Public

mySet_tb_Folders.fields("Folder_Short_Description") = Replace(Request.Form("myFolder_Short_Description"),"'"," ")
mySet_tb_Folders.fields("Folder_Long_Description") = Replace(Request.Form("myFolder_Long_Description"),"'"," ")

mySet_tb_Folders.fields("Folder_Public") = myFolder_Public

mySet_tb_Folders.fields("Folder_Modificator_ID") = myUser_ID
mySet_tb_Folders.fields("Folder_Modification_Date") = myDate_Now()

mySet_tb_Folders.update

mySet_tb_folders.close

set mySet_tb_Folders = nothing

'ERASE OLD FOLDER ACCES
myConnection.Execute("Delete * from tb_Folders_Access where Folder_ID = " & myFolder_ID)


 'SET FOLDERS ACCES IF FOLDER IS NOT PUBLIC
 if myFolder_Public = 0 Then
  
  'CREATE A RECORD SET FOR FOLDER ACCESS
  mySQL_Select_tb_Folders_Access = "Select * From tb_Folders_Access"
  set mySet_tb_Folders_Access  = Server.CreateObject("ADODB.RecordSet")
  mySet_tb_Folders_Access.open mySQL_Select_tb_Folders_Access, myConnection, 3,3
        mySet_tb_Folders_Access.MoveFirst
  'GET MEMBERS LIST
  
  mySQL_Select_tb_Sites_Member  = "Select * from tb_Sites_Members"
  set mySet_tb_Sites_Member = Server.CreateObject("ADODB.Recordset")
  set mySet_tb_Sites_Member = myConnection.Execute (MySQL_Select_tb_Sites_Member)
  mySet_tb_Sites_Member.MoveFirst
  
  Do while Not  mySet_tb_Sites_Member.eof
  
   IF Request.Form(mySet_tb_Sites_Member.fields("Member_Login")) = "on" Then

	 mySet_tb_Folders_Access.AddNew
     mySet_tb_Folders_Access.fields("Folder_ID") = myFolder_ID
	 mySet_tb_Folders_Access.fields("Member_ID") = mySet_tb_Sites_Member.fields("Member_ID")
     mySet_tb_Folders_Access.Update
	end if  
   mySet_tb_Sites_Member.MoveNext
  loop 

end if


response.redirect "__Folders_List.asp"


end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ADD A FOLDER
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If myAction = "Add" Then

myFolder_Name= Replace(Request.Form("myFolder_Name"),"'"," ")

 if len(myFolder_Name) = 0 Then  myError = "A" 

if not myValueIsGood(20,myFolder_Name) Then myError = "C"


 If Request.form("myFolder_Public") = "on" Then 
  myFolder_Public = 1
 else 
  myFolder_Public = 0 
 end if
 myFolder_Short_Description = Replace(Request.Form("myFolder_Short_Description"),"'"," ")
 myFolder_Long_Description =  Replace(Request.Form("myFolder_Long_Description"),"'"," ")
 'CONNECTION AND DB UPDATE
 'Open Connection
 set myConnection = Server.CreateObject("ADODB.Connection")
 myConnection.Open myConnection_String
 'Record Set Creation
 mySQL_Select_tb_Folders = "SELECT * FROM tb_Folders "
 Set mySet_tb_Folders = server.createobject("adodb.recordset")
 'CHECK IF THE FOLDER ALLREADY EXIST
 set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders & " Where Folder_Name= '" & myFolder_Name &"'")
 if mySet_tb_Folders.eof and len(myError) =0 Then
 mySet_tb_Folders.close
 mySet_tb_Folders.open mySQL_Select_tb_Folders, myConnection, 3,3
 mySet_tb_Folders.AddNew
 mySet_tb_Folders.fields("Folder_Name") = myFolder_Name
 mySet_tb_Folders.fields("Folder_Short_Description") = myFolder_Short_Description
 mySet_tb_Folders.fields("Folder_Long_Description") = myFolder_Long_Description
 mySet_tb_Folders.fields("Folder_Public") = myFolder_Public
 mySet_tb_Folders.fields("Folder_Creator_ID") = myUser_ID
 mySet_tb_Folders.fields("Folder_Modificator_ID") = myUser_ID
 mySet_tb_Folders.fields("Folder_Modification_Date") = myDate_Now()
 mySet_tb_Folders.fields("Site_ID") = mySite_ID
 mySet_tb_Folders.update
 set myFile_System_Object = Server.CreateObject("scripting.FileSystemObject")
 myFile_System_Object.CreateFolder myShared_Files_Path & myFolder_Name
 'SET FOLDERS ACCESS IF FOLDER IS NOT PUBLIC
 if myFolder_Public = 0 Then
  ' Get NEW Folder ID
  mySQL_Select_tb_Folders = "SELECT Max(Folder_ID) AS Last_Folder_ID From tb_Folders"
  ' Execute
  set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders)
  mySet_tb_Folders.MoveFirst
  myFolder_ID = mySet_tb_Folders.fields("Last_Folder_ID")
  mySet_tb_Folders.close
  set mySet_tb_Folders = NOTHING
  'CREATE A RECORD SET FOR FOLDER ACCESS
  set mySet_tb_Folders_Access = Server.CreateObject("ADODB.recordset")
  mySQL_Select_tb_Folders_Access = "Select * From tb_Folders_Access"
  mySet_tb_Folders_Access.open mySQL_Select_tb_Folders_Access, myConnection, 3,3
  'GET MEMBERS LIST
  mySQL_Select_tb_Sites_Member  = "Select * from tb_Sites_Members"
  set mySet_tb_Sites_Member = Server.CreateObject("ADODB.Recordset")
  set mySet_tb_Sites_Member = myConnection.Execute (MySQL_Select_tb_Sites_Member)
  mySet_tb_Sites_Member.MoveFirst
   Do while Not  mySet_tb_Sites_Member.eof
     IF Request.Form(mySet_tb_Sites_Member.fields("Member_Login")) = "on" Then
	  mySet_tb_Folders_Access.AddNew
      mySet_tb_Folders_Access.fields("Folder_ID") = myFolder_ID
	  mySet_tb_Folders_Access.fields("Member_ID") = mySet_tb_Sites_Member.fields("Member_ID")
      mySet_tb_Folders_Access.Update
	 end if  
     mySet_tb_Sites_Member.MoveNext
   loop 
  end if
 ' mySet_tb_Folders_Access.close
 else
  myError = myError  & "B"
 end if
 If len(myError) = 0 Then
  response.redirect "__Folders_List.asp"
 else 
  myAction="New"
 end if
end if 



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MODIFY A FOLDER
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction = "Modify" Then


' GET FOLDER INFORMATION

' dbConnection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

		
' Get Folders Informations
mySQL_Select_tb_Folders = "SELECT *,Folder_Modification_Date from tb_Folders, tb_Sites_Members  where Folder_Modificator_ID=Member_ID  AND Folder_ID = " & myFolder_ID  

' Execute
set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders)

if not mySet_tb_Folders.eof Then
 mySet_tb_Folders.MoveFirst
else Response.Redirect "__Folders_List.asp"
end if

myFolder_Name = mySet_tb_Folders.fields("Folder_Name") 
myFolder_Short_Description = mySet_tb_Folders.fields("Folder_Short_Description")
myFolder_Long_Description  = mySet_tb_Folders.fields("Folder_Long_Description")
myFolder_Public                  = mySet_tb_Folders.fields("Folder_Public")
myFolder_Modificator         = mySet_tb_Folders.fields("Member_Login")
myFolder_Modification_Date = mySet_tb_Folders.fields("Folder_Modification_Date")
'Get Modificator Login





end if
%>
<html>

<head>
<title><%=mySite_Name%> </title>
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

<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 

<%
' APPLICATION TITLE
%>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></TD></TR> 
</table>

<%
' CENTER APPLICATION

%> 


<%
' FORM BOXS
%>

<form action="__Folders_Modification.asp" method=post>
<%
if myAction = "New" Then 
%>
<input type=hidden value=Add name=myAction>
<%
end if
%>
<%
if myAction = "Modify" Then 
%>
<input type=hidden value=Update name=myAction>
<input type=hidden name=Folder_ID value="<%= myFolder_ID %>">
<%
end if
%>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="1" >
          
					<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=my_File_Message_Folder_Name%> &nbsp; </b>
			  

<%
IF (InStr(myError,"A") > 0) Then
%>			 
<br><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=myMessage_Required%></font> 
<%
end if
 
IF (InStr(myError,"B") > 0) and len(myError) = 1 Then
%>			 
<br><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>">* <%=myFile_Message_Folder_Name_Invalid%></font> 
<%
end if


IF (InStr(myError,"C") > 0) Then 
%>			 
<br><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>">* <%=myError_Message_Not_a_Valid_Directory_name  %></font> 
<%
end if
%>			
			
			
	
	
	
			  
			   </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"><%If myAction <> "New" Then%> 
			  <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"> 
              <b>&nbsp;&nbsp;<%=myFolder_Name%> &nbsp; </b>
			  <%else%>
			  &nbsp;&nbsp;<input type=text name=myFolder_Name>			  
			  <%end if%>
	 </font>
             </td>
          </tr>

					<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myMessage_Presentation%> &nbsp; </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
             &nbsp; <input name=myFolder_Short_Description type=text value="<%= myFolder_Short_Description %>">
             </td>
          </tr>
					<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myMessage_More%>&nbsp;</b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              &nbsp;&nbsp;<textarea rows=3 cols=35 name=myFolder_Long_Description><%= myFolder_Long_Description %></textarea>
             </td>
          </tr>
					
					<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=my_File_Message_Public%>&nbsp;&nbsp;</b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
            &nbsp;<input name=myFolder_Public type=Checkbox<%If myFolder_Public = 1 Then%> checked <%end if%>><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">(<%=my_File_Message_Check_Public%>)</font>
             </td>
          </tr>
					
					<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=my_File_Message_Check_Members%>&nbsp;&nbsp;</b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <table border="0" cellpadding="5" cellspacing="0"> 
<%
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Sites_Members ="  Select * from tb_Sites_Members  WHERE  Site_ID="&mySite_ID &" ORDER BY Member_Pseudo"
set mySet_tb_Sites_Members = myConnection.Execute(mySQL_Select_tb_Sites_Members)

if myAction = "Modify" And myFolder_Public = 0 Then

 mySQL_Select_tb_Folders_Access = " SELECT tb_Sites_members.Member_ID FROM tb_Sites_members, tb_Folders_Access WHERE (((tb_Sites_members.Member_ID)=[tb_Folders_Access].[Member_ID]) AND ((tb_Folders_Access.Folder_ID)=" & myFolder_ID & ")) ORDER BY tb_Sites_members.Member_Login"
 set mySet_tb_Folders_Access = myConnection.Execute(mySQL_Select_tb_Folders_Access)
 if not mySet_tb_Folders_Access.eof Then mySet_tb_Folders_Access.MoveFirst
end if


' each 3 members CRLF
myCounter = 1
do while not mySet_tb_Sites_Members.eof
			
	myParticipant_ID     = mySet_tb_Sites_Members("Member_ID")
	myParticipant_Pseudo = mySet_tb_Sites_Members("Member_Pseudo")
		
	if myCounter = 1 then
			Response.Write "<tr bgcolor=#ffffff>"
	end if
		
	
%> 

	<td valign="top"  bgcolor="<%=myBGColor%>">
	<small>
	<% If myAction = "New" OR myFolder_Public = 1 Then %>
	<input type="checkbox" name="<%=myParticipant_Pseudo%>"	 value="on" checked>
	<% end if%> 

	<% If myAction = "Modify" AND myFolder_Public = 0 Then %>
	
	   <%If Not mySet_tb_Folders_Access.eof Then 
	   	       If myParticipant_ID = mySet_tb_Folders_Access.fields("Member_ID") Then %>
	                <input type="checkbox" name="<%=myParticipant_Pseudo%>"	 value="on" checked>
	                <% mySet_tb_Folders_Access.MoveNext %>	
 
               <%else%> 
	                <input type="checkbox" name="<%=myParticipant_Pseudo%>"	 value="on" >
			   <%end if
	        else %>
         	 <input type="checkbox" name="<%=myParticipant_Pseudo%>"	 value="on" >
	  <%end if
	
	end if%> 
	
	
	
	
	<% 
	if myUser_type_ID <= 3 then
		%>
		<A HREF="__Site_Member_Information.asp?Member_ID=<%=mySet_tb_Sites_Members("Member_ID")%>"><Font face="Arial, Helvetica, sans-serif" size="2"><%=myParticipant_Pseudo%></font></a>
		<%
	else
		%>
		<font face="Arial, Helvetica, sans-serif" size="2"><%=myParticipant_Pseudo%></font>
		<%
	end if
	%>

	</small>
	</td>
	<%
	if myCounter = 3 then
		Response.Write "</tr>"
		myCounter = 1
	else
		myCounter= myCounter + 1
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
             </td>
          </tr>

					
<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b>&nbsp;</b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
             &nbsp; <input name=submit type=submit value="<%=myMessage_Go%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"></font>
             </td>
  </tr>

					



</table>	
</form>

<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>" align="CENTER"><FONT FACE="Arial, Helvetica, sans-serif"  Size="1" COLOR="<%=myApplicationTextColor%>"> 
&nbsp; <% If myAction="Modify" then %> <%=myDate_Display(myFolder_Modification_Date,2)%> -- <%=myFolder_Modificator%> <% end if%> </font></TD></TR> 
<TR><TD bgcolor="<%=myBGColor%>">
&nbsp;
<% if myAction <> "New" Then %>
<a href="__Folders_Modification.asp?myAction=New"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Add%></font></a>
<%end if
if myAction = "Modify" Then 
IF  (mySet_tb_Folders.fields("Folder_Creator_ID") = myUser_ID OR mySet_tb_Folders.fields("Folder_Responsible_ID") = myUser_ID OR myUser_Type_ID = 1) Then %>
, <a href="Javascript:if(confirm('<%=myFile_Message_Delete_Folder%>'))document.location='__Folders_Modification.asp?myAction=Delete&Folder_ID=<%= myFolder_ID %>';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Delete%></font></a>
<%
end if
end if%>



</tr>
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
' End Copyright									'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 

</body>
</html>
