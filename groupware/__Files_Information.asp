<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Styles_list.asp" is free software; 
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
' Doesn't Work With PWS ????
%>

<%
' ------------------------------------------------------------
' Name 			: __Files_Information.asp
' Path   		: /
' Version 		: 1.16.0
' Description 	: File Informations
' By			: Dania Tcherkezoff
' Company		: OverApps
' Date			: December ,10 2001
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Files_Information.asp"

Dim myPage_Application
myPage_Application="Files"
	
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

Dim mySearch, myMaxRspByPage, myUser_Can_Modify

Dim mySet_tb_Files2,mySet_tb_Files, mySQL_Select_tb_Files, myFolder_Responsible_ID

Dim i, j, mySet_tb_Folders_Access, mySQL_tb_Folders_Access, myFile_ID

Dim mySQL_Select_tb_Folders, mySet_tb_Folders, myFolder_ID,myFolder_Name, myFolder_Long_Description, myFolder_Short_Description, myFolder_Creator, MyFolder_Files_Number

'''''''''''''''''''''''''''''''''''''''''''''
' GET PARAMETERS
'''''''''''''''''''''''''''''''''''''''''''''
myFolder_ID = Request.QueryString("Folder_ID")
myFile_ID     = Request.queryString("File_ID")


' dbConnection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

		
' Get Folders Informations
mySQL_Select_tb_Folders = "SELECT  Max(tb_Files.File_Modification_Date) AS Last_File_Date, tb_Folders.Folder_Responsible_ID, tb_Folders.Folder_Creator_ID, tb_Folders.Folder_Name, tb_Folders.Folder_Short_Description, Count(tb_Files.File_ID) AS NBFile, tb_Folders.Folder_Public, tb_Folders.Folder_ID, tb_Sites_members.Member_Login as creator FROM (tb_Folders LEFT JOIN tb_Files ON tb_Folders.Folder_ID = tb_Files.File_Folder_ID) INNER JOIN tb_Sites_members ON tb_Folders.Folder_Modificator_ID = tb_Sites_members.Member_ID WHERE (((tb_Folders.Folder_ID)= "& myFolder_ID & " )) GROUP BY tb_Folders.Folder_Responsible_ID, tb_Folders.Folder_Creator_ID, tb_Folders.Folder_Name, tb_Folders.Folder_Short_Description, tb_Folders.Folder_Public, tb_Folders.Folder_ID, tb_Sites_members.Member_Login ORDER BY tb_Folders.Folder_Name;"



' Execute
set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders)

if mySet_tb_Folders.eof Then
	 Response.Redirect "__Folders_List.asp"
else
   mySet_tb_Folders.MoveFirst
end if	 


'TEST ID USER HAS ACCESS
 mySQL_tb_Folders_Access = "Select * from tb_Folders_Access where Folder_ID = " & mySet_tb_Folders.fields("Folder_ID") & " AND Member_ID = "& myUser_ID &";"  
 set mySet_tb_Folders_Access = myConnection.Execute(mySQL_tb_Folders_Access)


 
 
 
if   mySet_tb_Folders.fields("Folder_Creator_ID") <> myUser_ID AND mySet_tb_Folders.fields("Folder_Responsible_ID") <> myUser_ID AND myUser_Type_ID > 1 Then
	 if mySet_tb_Folders_Access.eof AND mySet_tb_Folders.fields("Folder_Public") = 0 Then Response.Redirect "__Folders_List.asp" 
else 	 
	 myUser_Can_Modify = 1
end if

mySet_tb_Folders_Access.close
Set mySet_tb_Folders_Access=Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET FOLDER INFORMATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''

myFolder_Name = mySet_tb_Folders.fields("Folder_Name")
myFolder_Creator = mySet_tb_Folders.fields("Member_Login")
myFolder_Short_Description = mySet_tb_Folders.fields("Folder_Short_Description")
myFolder_Long_Description = mySet_tb_Folders.fields("Folder_Long_Description")
myFolder_Creator = mySet_tb_Folders.fields("Member_Login")
myFolder_Files_Number = mySet_tb_Folders.fields("NBFile")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET FILE INFORMATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

set mySet_tb_Files = Server.CreateObject("ADODB.RecordSet")
mySQL_Select_tb_Files = "Select * from tb_files,tb_Sites_Members where File_ID = " & myFile_ID & " AND tb_Files.File_Modificator_ID =tb_Sites_Members.Member_ID" 
set mySet_tb_Files = myConnection.Execute(mySQL_Select_tb_Files)


set mySet_tb_Files2 = Server.CreateObject("ADODB.RecordSet")
mySQL_Select_tb_Files = "Select * from tb_files,tb_Sites_Members where File_ID = " & myFile_ID & " AND tb_Files.File_Creator_ID =tb_Sites_Members.Member_ID" 
set mySet_tb_Files2 = myConnection.Execute(mySQL_Select_tb_Files)

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
<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 

<%
' APPLICATION TITLE
%>

<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=my_File_Message_Folder%></b></font></TD></TR> 
</table>


<BR> 

<%
' FOLDER INFORMATION
%>
<table width="<%=myApplication_Width%>" border=0>

<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=my_File_Message%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<a href="__Files_Download.asp?file_ID=<%=mySet_tb_Files.fields("File_ID")%>&Folder_ID=<%= myFolder_ID %>" ><%=mySet_tb_Files.fields("File_Name")%></a></font>
</td>
</tr>

<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myMessage_Presentation%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=mySet_tb_Files.fields("File_Short_Description")%></font>
</td>
</tr>                

<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myMessage_More%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=mySet_tb_Files.fields("File_Long_Description")%></font>
</td>
</tr>    

<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=my_File_Message_Type%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=mySet_tb_Files.fields("File_Type")%></font>
</td>
</tr>   

<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myFile_Message_File_Creator%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=mySet_tb_Files2.fields("Member_Login")%></font>
</td>
</tr>  


<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=my_File_Message_Size%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;
<%'
'DISPLAY  SIZE IN Ko OR Mo
 If mySet_tb_Files.fields("File_Size")  <  1048576 Then %>
 
  <%= (int(( mySet_tb_Files.fields("File_Size") / 1024)*100)) / 100 %> Ko 
   
<%
else
%>

 <%= (int(( mySet_tb_Files.fields("File_Size") / (1024*1024))*100)) / 100 %> Mo 
 
<%
end if
%>
						 
						 
						 </font>
</td>
</tr>    






</table>

<br>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>" align="CENTER"><FONT FACE="Arial, Helvetica, sans-serif"  Size="1" COLOR="<%=myApplicationTextColor%>"> 
&nbsp;  <%=myDate_Display(mySet_tb_Files.fields("File_Modification_Date"),2)%> -- <%=mySet_tb_Files.fields("Member_Login")%>  </font></TD></TR> 
<TR><TD bgcolor="<%=myBGColor%>">
&nbsp;<a href="__Files_Modification.asp?myAction=New&Folder_ID=<%= myFolder_ID %>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Add%></font></a>
<%
If myUser_ID = mySet_tb_Files.fields("File_Creator_ID") OR myUser_ID = mySet_tb_Files.fields("File_Responsible_ID") OR myUSer_Type_ID = 1 Then
%>
,&nbsp;<a href="__Files_Modification.asp?Folder_ID=<%= myFolder_ID %>&File_ID=<%=mySet_tb_Files.fields("File_ID")%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Modification%></font></a>
,&nbsp;<a href="__Files_Modification.asp?myAction=Delete&Folder_ID=<%= myFolder_ID %>&File_ID=<%=mySet_tb_Files.fields("File_ID")%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Delete%></font></a>
<%
end if
%>
</table>
</td>
</tr>
</table>
<%
' Close Recordset
mySet_tb_Folders.close
Set mySet_tb_Folders=Nothing
' Close Connection 	
myConnection.Close
set myConnection = Nothing
%>
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
'				    End Copyright												'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 
</body>
</html>

<html></html>
<html><script language="JavaScript"></script></html>