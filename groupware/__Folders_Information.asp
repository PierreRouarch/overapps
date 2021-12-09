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
' Name 			: __Folders_Information.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Folders Lists
' By			: Dania Tcherkezoff
' Company		: OverApps
' Date			: December, 10 2001
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Folders_Information.asp"

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


Dim i, j, mySet_tb_Folders_Access, mySQL_tb_Folders_Access,myFolder_Modificator, myFolder_Modification_Date

Dim mySQL_Select_tb_Folders,mySQL_Select_tb_Folders2, mySet_tb_Folders,mySet_tb_Folders2, myFolder_ID,myFolder_Name, myFolder_Long_Description, myFolder_Short_Description, myFolder_Creator, MyFolder_Files_Number

'''''''''''''''''''''''''''''''''''''''''''''
' GET PARAMETERS
'''''''''''''''''''''''''''''''''''''''''''''
myFolder_ID = Request.QueryString("ID")


' dbConnection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

		
' Get Folders Number of Files 
  mySQL_Select_tb_Folders2 = "SELECT Folder_Name,Folder_ID,Max(tb_Files.File_Modification_Date) as Last_File_Date , Count(tb_Files.File_ID) AS NBFile FROM tb_Folders, tb_Files WHERE tb_Folders.Folder_ID = tb_Files.File_Folder_ID  GROUP BY tb_Folders.Folder_ID,Folder_Name  ORDER BY tb_Folders.Folder_Name;"
' Execute
set mySet_tb_Folders2 = myConnection.Execute(mySQL_Select_tb_Folders2)

'GET FOLDERS INFORMATIONS

mySQL_Select_tb_Folders = "SELECT *,Folder_Modification_Date from tb_Folders, tb_Sites_Members  where Folder_Modificator_ID=Member_ID  AND Folder_ID = " & myFolder_ID

set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders)


if mySet_tb_Folders.eof Then
	 Response.Redirect "__Folders_List.asp"
else
   mySet_tb_Folders.MoveFirst
end if	 

'TEST IF USER HAS ACCESS
 mySQL_tb_Folders_Access = "Select * from tb_Folders_Access where Folder_ID = " & mySet_tb_Folders.fields("Folder_ID") & " AND Member_ID = "& myUser_ID &";"  
 set mySet_tb_Folders_Access = myConnection.Execute(mySQL_tb_Folders_Access)

if   mySet_tb_Folders.fields("Folder_Creator_ID") <> myUser_ID AND mySet_tb_Folders.fields("Folder_Responsible_ID") <> myUser_ID AND myUser_Type_ID > 1 Then
	 if mySet_tb_Folders_Access.eof AND mySet_tb_Folders.fields("Folder_Public") = 0Then Response.Redirect "__Folders_List.asp" 
else 	 
	 myUser_Can_Modify = 1
end if

mySet_tb_Folders_Access.close
Set mySet_tb_Folders_Access=Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET FOLDER INFORMATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''

myFolder_Name = mySet_tb_Folders.fields("Folder_Name")
myFolder_Short_Description = mySet_tb_Folders.fields("Folder_Short_Description")
myFolder_Long_Description = mySet_tb_Folders.fields("Folder_Long_Description")
myFolder_Modificator = mySet_tb_Folders.fields("Member_Login")
myFolder_Modification_Date = mySet_tb_Folders.fields("Folder_Modification_Date")

myFolder_Files_Number = mySet_tb_Folders2.fields("NBFile")


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
              <b><%=my_File_Message_Folder_Name%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=myFolder_Name%></font>
</td>
</tr>

<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myMessage_Presentation%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=myFolder_Short_Description%></font>
</td>
</tr>                

<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myMessage_More%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=myFolder_Long_Description%></font>
</td>
</tr>    
</tr>     

<tr> 
            <td align="right" bgcolor="<%=myBorderColor%>"   height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=my_File_Message_File_Number%></b> &nbsp;</font></td>
              <td bgcolor="<%=myBGColor%>" align=left  height="10">
						 <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=myFolder_Files_Number%></font>
</td>
</tr>      






</table>

<br>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>" align=center><FONT FACE="Arial, Helvetica, sans-serif"  Size="1" COLOR="<%=myApplicationTextColor%>"> 
&nbsp;  <%=myDate_Display(myFolder_Modification_Date,2)%> -- <%=myFolder_Modificator%>  </font></TD></TR> 
<TR><TD bgcolor="<%=myBGColor%>">
&nbsp;<a href="__Folders_Modification.asp?myAction=New"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=myMessage_Add%></font></a>
<%
If myUser_Can_Modify = 1 Then 
%>
,&nbsp;<a href="__Folders_Modification.asp?ID=<%=mySet_tb_Folders.fields("Folder_ID")%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=myMessage_Modification%></font></a>
,&nbsp;<a href="Javascript:if(confirm('<%=myFile_Message_Delete_Folder%>'))document.location='__Folders_Modification.asp?myAction=Delete&Folder_ID=<%=mySet_tb_Folders.fields("Folder_ID")%>';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">&nbsp;<%=myMessage_Delete%></font></a>
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
