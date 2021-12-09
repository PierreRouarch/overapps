
<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Files_Modification.asp" is free software; 
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
' at the bottom of the page with an active link from the name "OverApps & Contributors"  
' the Address http://www.overapps.com where the user (netsurfer) could find the 
' appropriate information about this license. 
'
'-----------------------------------------------------------------------------
%>


<% 	
   ' Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Doesn't Work With PWS ????
Server.ScriptTimeOut = 3600 ' For Big Big  big Files
%>
<%
' ------------------------------------------------------------
' Name 			: __Files_Modification.asp
' Path  	    :   /
' Version 	    : 1.15.0
' Description   : Files Modification
' By		    : Dania TCHERKEZOFF
' Company	    : OverApps
' Date			: April, 17, 2002
'
' Modify by		: Pierre Rouarch			
' Company		:
' Date			: November, 15 2002 
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Files_Modification.asp"

Dim myPage_Application
myPage_Application="Files"

Dim  mySQL_Select_tb_Meetings_Members, mySet_tb_Meetings_Members, mySQL_Insert_tb_Meetings_members,  mySQL_Delete_tb_Meetings_Members 

Dim URL, myCounter, myParticipant_ID, myMember_Name, myParticipant_Pseudo, myParticipant_Pseudo_Value, myUpload,mySet_tb_Files, mySQL_Select_tb_Files, myFolder_Creation_Date,myFolder_Responsible

Dim myMembers_Public_Type_ID, myAction, myError, myFolder_Public, mySQL_Select_Folders_Access, mySet_tb_Folders_Access, mySet_tb_Sites_Member,mySQL_Select_tb_Sites_Member,mySQL_Select_tb_Folders_Access

Dim mySQL_Select_tb_Folders, mySet_tb_Folders, myFolder_ID,myFolder_Name, myFolder_Long_Description, myFolder_Short_Description, myFolder_Creator, MyFolder_Files_Number, myFile_System_Object, myFile_Creator_ID, myFile_Responsible_ID,myFolder_Modification_Date

Dim  myFile_Modification_Date, myFile_Modificator

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' INCLUDES 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<!-- #include file="_INCLUDE/Files_Upload_Class.asp" -->

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

 if myAction <> "Delete" Then Set myUpload = New UplFile

if len(myAction ) = 0 Then myAction = myUpload.Form_Field("myAction") 
if len(myAction) = 0 Then myAction = "Modify"

myError = Request.QueryString("myError")
response.write myError

myFolder_ID = Request.QueryString("ID")

if len(myFolder_ID) = 0 Then myFolder_ID = Request.QueryString("Folder_ID")
if len(myFolder_ID) = 0 Then myFolder_ID  = myUpload.Form_Field("Folder_ID")

myFile_ID = Request.QueryString("File_ID")

if len(myFile_ID=0) and myAction <> "Delete" Then  myFile_ID = myUpload.Form_Field("File_ID")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHECK IF USER  CAN USE  THE CURRENT FOLDER
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' dbConnection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

		
' Get Folders Informations
mySQL_Select_tb_Folders = " Select * from tb_Folders where Folder_ID = " & myFolder_ID 
set mySet_tb_Folders = Server.CreateObject("ADODB.RecordSet")
set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders)

'Get Folders Access Informations
mySQL_Select_tb_Folders_Access = " Select * from tb_Folders_Access where Folder_ID = " & myFolder_ID & " AND Member_ID = " & myUser_ID 
set mySet_tb_Folders_Access = Server.CreateObject("ADODB.RecordSet")
set mySet_tb_Folders_Access = myConnection.Execute(mySQL_Select_tb_Folders_Access)

If myUser_ID <> mySet_tb_Folders("Folder_Creator_ID") AND myUser_ID <> mySet_tb_Folders("Folder_Responsible_ID") and myUser_Type_ID = 1 AND  mySet_tb_Folders_Access.eof Then 
 Response.Redirect("__Folders_List.asp")
end if

mySet_tb_Folders_Access.close
set mySet_tb_Folders_Access = Nothing


 myFolder_Name = mySet_tb_Folders("Folder_Name")
 myFolder_Creator = mySet_tb_Folders("Folder_Creator_ID")
 myFolder_Short_Description = mySet_tb_Folders("Folder_Short_Description")
 myFolder_Long_Description  = mySet_tb_Folders("Folder_Long_Description")
 myFolder_Modification_Date      = mySet_tb_Folders("Folder_Modification_Date")
 myfolder_Responsible          = mySet_tb_Folders("Folder_Responsible_ID")
 
 mySet_tb_Folders.close
 set mySEt_tb_Folders = Nothing                   

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'DELETE A FILE
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction = "Delete" Then
 
 ' dbConnection
 set myConnection = Server.CreateObject("ADODB.Connection")
 myConnection.Open myConnection_String
 
 'Test if User can Delete this Message
 mySQL_Select_tb_Files = "Select * from tb_Files where File_ID =" & myFile_ID
 set mySet_tb_Files = Server.CreateObject("ADODB.RecordSet")
 set mySet_tb_Files = myConnection.Execute( mySQL_Select_tb_Files )
 
 'Redirection if File doesn't exist
 If mySet_tb_Files.eof Then
 
  mySet_tb_Files.close
  set mySet_tb_Files = NOTHING
 
  myconnection.close
  Set myConnection = Nothing
  
  Response.redirect "__Folder_List.asp"
  end if
  
  myFile_Name= mySet_tb_Files.fields("File_Name")
  
  if myUser_ID = mySet_tb_Files.fields("File_Creator_ID") or myUser_ID = mySet_tb_Files.fields("File_Responsible_ID") or myUser_Type_ID = 1 Then
   'ERASE FROM DB
   myConnection.Execute("Delete from tb_Files where File_ID =" & myFile_ID)
   'ERASE FROM  DRIVE
   set myFile_System_Object = Server.CreateObject("scripting.FileSystemObject")
   myFile_System_Object.DeleteFile myShared_Files_Path & myFolder_Name & "\" & myFile_Name
 
   mySet_tb_Files.close
   set mySet_tb_Files = NOTHING
 
   myconnection.close
   Set myConnection = Nothing
   Response.Redirect "__Files_List.asp?folder_id="&myFolder_ID
 
  end if
 
  Response.Redirect "__Folder_List.asp"
 
end if
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ADD A FILE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim myFile_Extension, mySet_tb_Files_Extensions, mySQL_Select_tb_Files_Extensions

if myAction = "Add" Then

 If myUpload.Nb_Files = 1 Then
 
  'TEST IF FILE IS NOT TOO BIG AND Extension  AUTHORISED AND IF FILE DOESN'T EXIST INF THIS FOLDER
 
  myFile_Extension = Right(myUpload.File_Name(1), Len(myUpload.File_Name(1)) - InStrRev(myUpload.File_Name(1),"."))
 
  mySQL_Select_tb_Files_Extensions = "Select * from tb_Files_Extensions where File_Extension_Autorised = 1 AND File_Extension = '" & myFile_Extension & "'"
  set mySet_tb_Files_Extensions = Server.CreateObject("ADODB.RecordSet")
  set mySet_tb_Files_Extensions = myConnection.Execute(mySQL_Select_tb_Files_Extensions)
  if  mySet_tb_Files_Extensions.eof  Then myError = "A"
  
  mySQL_select_tb_Files = "Select * from tb_Files where File_Name = '" & myUpload.File_Name(1) &"' AND File_Folder_ID=" & myFolder_ID
  set mySet_tb_Files = Server.CreateObject("ADODB.RecordSet")  
  set mySet_tb_Files = myConnection.Execute(mySQL_Select_tb_Files)
  if not mySet_tb_Files.eof Then myError = myError & "C"
  mySet_tb_Files.close
  set mySet_tb_Files = Nothing
  
  
IF myUpload.File_Size(1) > myMaximum_File_Size Then myError = myError & "B"

  if len(myError) = 0 Then
   mySQL_Select_tb_Files = " Select * from tb_Files"
   set mySet_tb_Files = Server.CreateObject("ADODB.RecordSet")
   mySet_tb_Files.open mySQL_Select_tb_Files, myConnection, 3,3
   mySet_tb_Files.AddNew
   mySet_tb_Files.fields("File_Name") = myUpload.File_Name(1)
   mySet_tb_Files.fields("File_Size") = myUpload.File_Size(1)
   mySet_tb_Files.fields("File_Type")  = myUpload.File_Type(1)         
   mySet_tb_Files.fields("File_Creator_ID") = myUser_ID
   mySet_tb_Files.fields("File_Modification_Date") = myDate_Now()
   mySet_tb_Files.fields("File_Short_Description") = Replace( myUpload.Form_Field("myFile_Short_Description") ,"'"," ")
   mySet_tb_Files.fields("File_Long_Description") = Replace( myUpload.Form_Field("myFile_Long_Description") ,"'"," ")
   mySet_tb_Files.fields("File_Folder_ID") = myFolder_ID
   mySet_tb_Files.fields("File_Modificator_ID") = myUSer_ID
   mySet_tb_Files.fields("Site_ID") = mySite_ID
   mySet_tb_Files.Update
   myUpload.Save_File(1)
   set myFile_System_Object=server.createobject("scripting.FileSystemObject")
   myFile_System_Object.MoveFile myShared_Files_Path  & myUpload.File_Name(1) , myShared_Files_Path & myFolder_Name & "\" &  myUpload.File_Name(1) 
   set myFile_System_Object=Nothing
   mySet_tb_Files.close
   set mySet_tb_Files = Nothing
   myConnection.close
   set myconnection =Nothing
   if len(myError)=0 Then Response.Redirect "__Files_List.asp?Folder_ID=" & myFolder_ID
  end if     
 end if
 myAction = "New"
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UPDATE FILE INFORMATION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction = "Update" Then 

 if len(myFile_ID) = 0 Then
   myConnection.close
   set myConnection = Nothing
   Response.Redirect "__Folders_List.asp"
 end if

 mySQL_Select_tb_Files = "Select * from tb_Files where File_ID =" & myFile_ID
 set mySet_tb_Files = Server.Createobject("ADODB.RecordSet")
 mySet_tb_Files.open mySQL_Select_tb_Files, myConnection, 3,3
 if not mySet_tb_Files.eof Then
  mySet_tb_Files.fields("File_Modificator_ID") = myUSer_ID
  mySet_tb_Files.fields("File_Short_Description") = Replace( myUpload.Form_Field("myFile_Short_Description") ,"'"," ")
  mySet_tb_Files.fields("File_Long_Description") = Replace( myUpload.Form_Field("myFile_Long_Description") ,"'"," ")
  mySet_tb_Files.fields("File_Modification_Date") = myDate_Now()
  mySet_tb_Files.update
 end if   
  mySet_tb_Files.close
  set mySet_tb_Files = Nothing
  myConnection.close
  set myConnection = Nothing
 
 Response.Redirect "__Files_List.asp?Folder_ID="&myFolder_ID
end if 


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET INFORMATION IF ACTION IS MODIFY
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim myFile_ID, myFile_Short_Description, myFile_Long_Description, myFile_Name

if myAction = "Modify" Then
 myFile_ID = Request.QueryString("File_ID") 
 if len(myFile_ID) = 0 Then Response.Redirect "__Folders_List.asp"
 mySQL_Select_tb_Files = "Select * from tb_Files,tb_Sites_Members  where  File_Modificator_ID=Member_ID  AND File_ID="  & myFile_ID
 set mySet_tb_Files = Server.CreateObject("ADODB.RecordSet")
 set mySet_tb_Files = myConnection.Execute(mySQL_Select_tb_Files)
 
 if mySet_tb_Files.eof Then
   mySet_tb_Files.close
   mySet_tb_Files = Nothing
   myConnection.close
   myConnection=Nothing
   Response.Redirect "__Folders_List.asp"
 end if
 
 myFile_Short_Description = mySet_tb_Files.fields("File_Short_Description")
 myFile_Long_Description  = mySet_tb_Files.fields("File_Long_Description") 
 myFile_Name = mySet_tb_Files.fields("File_Name")
 myFile_ID      = mySet_tb_Files.fields("File_ID")
 myFile_Creator_ID = mySet_tb_Files.fields("File_Creator_ID")
 myFile_Responsible_ID = mySet_tb_Files.fields("File_Responsible_ID")
 myFile_Modification_Date = mySet_tb_Files.fields("File_Modification_Date")
 myFile_Modificator      = mySet_tb_Files.fields("Member_Login")
  mySet_tb_Files.close
  set mySet_tb_Files = Nothing
  myConnection.close
  set myConnection=Nothing

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
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></TD></TR> 
</table>


<%
' FORM BOXS
%>

<form action="__Files_Modification.asp" method=post enctype="multipart/form-data">
<%
if myAction = "New" Then 
%>
<input type=hidden value=Add name=myAction>
<input type=hidden name=Folder_ID value="<%= myFolder_ID %>"> 
<%
end if
%>
<%
if myAction = "Modify" Then 
%>
<input type=hidden value=Update name=myAction>
<input type=hidden name=File_ID value="<%= myFile_ID %>">
<input type=hidden name=Folder_ID value="<%= myFolder_ID %>"> 
<%
end if
%>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="1" >
          
		<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=my_File_Message%> &nbsp; </b> <br>
<%= myFile_Message_Maximum_Size2 %> (			  
<%'
'DISPLAY MAXIMUM SIZE IN Ko OR Mo
 If myMaximum_File_Size <  1048576 Then %>
 
  <%= (int(( myMaximum_File_Size / 1024)*100)) / 100 %> Ko )
   
<%
else
%>

 <%= (int(( myMaximum_File_Size / (1024*1024))*100)) / 100 %> Mo )
 
<%
end if
%>
			  
			  
			  
			  
<%
IF (InStr(myError,"B") > 0) Then
%>			 
<br><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><b> * <%=myFile_Message_File_Too_Big%></b></font> 
<%
end if	
	

IF (InStr(myError,"A") > 0) Then
%>			 
<br><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><b> * <%=myFile_Message_Extension_Not_Allowed%></b></font> 
<%
end if			  
%>

<%
IF (InStr(myError,"C") > 0) Then
%>			 
<br><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><b> * <%=myFile_Message_Files_Exists%></b></font> 
<%
end if			  
%>					  
			  </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
             &nbsp; 
<%
If myAction = "Modify" Then
%>
<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><b><%=myFile_Name%></b></font>
<%
else
%>
 <input name=myFile type=File size=50>
 
<%
end if
%> 
             </td>
          </tr>
		  
<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myMessage_Presentation%> &nbsp; </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
             &nbsp; <input name=myFile_Short_Description type=text value="<%= myFile_Short_Description %>" size=50>
             </td>
          </tr>
		  
		<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myMessage_More%> &nbsp; </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
             &nbsp; <textarea rows=3 cols=38 name=myFile_Long_Description><%= myFile_Long_Description %></textarea>
             </td>
          </tr>  
		  
		<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><input type=submit value="<%=myMessage_go%>"> &nbsp; </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2">&nbsp; 
              
             </td>
          </tr> 
</table>
<br>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>" align="CENTER"><FONT FACE="Arial, Helvetica, sans-serif"  Size="1" COLOR="<%=myApplicationTextColor%>"> 
&nbsp;  <%=myDate_Display(myFile_Modification_Date,2)%> -- <%=myFile_Modificator%>  </font></TD></TR> 
<TR><TD bgcolor="<%=myBGColor%>">
&nbsp;<a href="__Files_Modification.asp?myAction=New&Folder_ID=<%= myFolder_ID %>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Add%></font></a>
<%
If myUser_ID = myFile_Creator_ID OR myUser_ID = myFile_Responsible_ID OR myUSer_Type_ID = 1 and myAction<>"New" then %>
,&nbsp;<a href="__Files_Modification.asp?myAction=Delete&Folder_ID=<%= myFolder_ID %>&File_ID=<%=myFile_ID%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Delete%></font></a>
<%
end if
%>

</table>
</form>
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


<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>