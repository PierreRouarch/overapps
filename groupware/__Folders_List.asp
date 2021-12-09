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
' Name 			: __Folders_list.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	:  List all folder user has access
' By		       : Dania Tcherkezoff
' Company	 : OverApps
' Date			  : December, 10 2001 		
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Folders_List.asp"

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

Dim mySearch, myMaxRspByPage

Dim i, j, mySet_tb_Folders_Access, mySQL_tb_Folders_Access

Dim mySQL_Select_tb_Folders, mySet_tb_Folders

Dim myFolder_ID, myNumPage, myOrder,myNbrPage,indice, myRS


'MAX FOLDER PER PAGE
myMaxRspByPage=10


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET PARAMETERS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

mySearch=Replace(Request.Form("search"),"'","''")
if len(mysearch)=0 Then mySearch = Replace(Request.QueryString("search"),"'","''")

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
' SEARCH BOX
%>


<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0"> 
<tr ALIGN="CENTER">

<td> 
<form method="post" action="__Folders_List.asp" id=form1 name=form1> 
<br> &nbsp; <input type="text" name="search" size="30" VALUE="<%=mySearch%>"> &nbsp; <INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"> 
</form>
</td>
</tr>
</table>


<BR> 

<%
' LISTING

' dbConnection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

myNumPage=Request.QueryString("Page")
if Len(myNumPage)=0 then 
	myNumPage=1
end if

myOrder = Request.QueryString("order")


' Get Folders Informations

If len(mySearch) = 0 Then
 'No search criteria
 if len(myOrder) = 0 Then
  mySQL_Select_tb_Folders = "SELECT Max(tb_Files.File_Modification_Date) as Last_File_Date ,Folder_Responsible_ID,Folder_Creator_ID,tb_Folders.Folder_Name, tb_Folders.Folder_Short_Description, Count(tb_Files.File_ID) AS NBFile, tb_Folders.Folder_Public, tb_Folders.Folder_ID FROM tb_Folders LEFT JOIN tb_Files ON tb_Folders.Folder_ID = tb_Files.File_Folder_ID   Where tb_Folders.Site_Id = "&mySite_ID&"  GROUP BY tb_Folders.Folder_Name, tb_Folders.Folder_Public, tb_Folders.Folder_ID, tb_Folders.Folder_Short_Description,Folder_Creator_ID,Folder_Responsible_ID ORDER BY tb_Folders.Folder_Name;"
 else 
  mySQL_Select_tb_Folders = "SELECT Max(tb_Files.File_Modification_Date) as Last_File_Date ,Folder_Responsible_ID,Folder_Creator_ID,tb_Folders.Folder_Name, tb_Folders.Folder_Short_Description, Count(tb_Files.File_ID) AS NBFile, tb_Folders.Folder_Public, tb_Folders.Folder_ID FROM tb_Folders LEFT JOIN tb_Files ON tb_Folders.Folder_ID = tb_Files.File_Folder_ID Where tb_Folders.Site_Id = "& mySite_ID  &"GROUP BY tb_Folders.Folder_Name, tb_Folders.Folder_Public, tb_Folders.Folder_ID, tb_Folders.Folder_Short_Description,Folder_Creator_ID,Folder_Responsible_ID ORDER BY " & myOrder
 end if

Else
 'With search criteria
  if len(myOrder) = 0 Then
  mySQL_Select_tb_Folders = "SELECT Max(tb_Files.File_Modification_Date) as Last_File_Date ,Folder_Responsible_ID,Folder_Creator_ID,tb_Folders.Folder_Name, tb_Folders.Folder_Short_Description, Count(tb_Files.File_ID) AS NBFile, tb_Folders.Folder_Public, tb_Folders.Folder_ID FROM tb_Folders LEFT JOIN tb_Files ON tb_Folders.Folder_ID = tb_Files.File_Folder_ID Where (tb_Folders.Folder_Name like '%"& mySearch &"%'  OR  tb_Folders.Folder_Short_Description like '%"& mySearch &"%' ) AND tb_Folders.Site_Id = "&mySite_ID&"  GROUP BY tb_Folders.Folder_Name, tb_Folders.Folder_Public, tb_Folders.Folder_ID, tb_Folders.Folder_Short_Description,Folder_Creator_ID,Folder_Responsible_ID ORDER BY tb_Folders.Folder_Name;"
 else 
  mySQL_Select_tb_Folders = "SELECT Max(tb_Files.File_Modification_Date) as Last_File_Date ,Folder_Responsible_ID,Folder_Creator_ID,tb_Folders.Folder_Name, tb_Folders.Folder_Short_Description, Count(tb_Files.File_ID) AS NBFile, tb_Folders.Folder_Public, tb_Folders.Folder_ID FROM tb_Folders LEFT JOIN tb_Files ON tb_Folders.Folder_ID = tb_Files.File_Folder_ID  Where (tb_Folders.Folder_Name like '%"& mySearch &"%'  OR  tb_Folders.Folder_Short_Description like '%"& mySearch &"%') AND tb_Folders.Site_Id = "&mySite_ID&" GROUP BY tb_Folders.Folder_Name, tb_Folders.Folder_Public, tb_Folders.Folder_ID, tb_Folders.Folder_Short_Description,Folder_Creator_ID,Folder_Responsible_ID ORDER BY " & myOrder
 end if
 

End if 




' Execute
set mySet_tb_Folders = Server.CreateObject("ADODB.RecordSet")
set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders)

%>



<%
' Go to the current record
i=0
myRs=(myNumPage-1)*myMaxRspByPage
j=0
if not mySet_tb_Folders.bof then mySet_tb_Folders.MoveFirst
do while not mySet_tb_Folders.eof 
i=i+1
mySet_tb_Folders.movenext
loop 
if not mySet_tb_Folders.bof then 
mySet_tb_Folders.MoveFirst
mySet_tb_Folders.Move(myRs) 
end if
%> 

<%
' ROW TITLES
%>
<table width="<%=myApplication_Width%>" border=0>
<tr bgcolor=<%=myBorderColor%>>
<td>&nbsp;&nbsp;<b> 
<%If len(myOrder) = 0 Then%> 
<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> <%=my_File_Message_Folder%></font>
<%else%>
<a href="__Folders_List.asp?search=<%=mySearch%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> <%=my_File_Message_Folder%></font></a>
<%end if%>


</b></td>

<td> &nbsp;&nbsp;<b> 
<%If myOrder = "Folder_Short_Description" Then%> 
<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=my_File_Message_Description%>
<%else%></font>
<a href="__Folders_List.asp?order=Folder_Short_Description&search=<%=mySearch%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"><%=my_File_Message_Description%></font></a>
<%end if%>
</b></td>
<td><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> &nbsp;&nbsp;<b> <%=my_File_Message_Acces%></b></font></td>
<td><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> &nbsp;&nbsp;<b> <%=my_File_Message_Last_Upload%></b></font></td>
<td><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> &nbsp;&nbsp;<b> <%=myMessage_More%></b></font></td>
</tr>


<%
' LISTING
%>

<%	
do while not mySet_tb_Folders.eof  AND (myMaxRspByPage>j)
	j=j+1

'TEST ID USER HAS ACCESS
 mySQL_tb_Folders_Access = "Select * from tb_Folders_Access where Folder_ID = " & mySet_tb_Folders.fields("Folder_ID") & " AND Member_ID = "& myUser_ID &";"  
 set mySet_tb_Folders_Access = myConnection.Execute(mySQL_tb_Folders_Access)

if mySet_tb_Folders_Access.eof AND mySet_tb_Folders.fields("Folder_Public") = 0 AND mySet_tb_Folders.fields("Folder_Creator_ID") <> myUser_ID AND mySet_tb_Folders.fields("Folder_Responsible_ID") <> myUser_ID AND myUser_Type_ID > 1 Then 

else 
%>

<tr>
<td>&nbsp;&nbsp;
<a href="__Files_List.asp?Folder_ID=<%= mySet_tb_Folders.fields("Folder_ID") %>">
<img src="Images/Files_Folder.gif" alt="Explore" border=0></a>
&nbsp;<a href="__Files_List.asp?Folder_ID=<%= mySet_tb_Folders.fields("Folder_ID") %>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><b><%=mySet_tb_Folders.fields("Folder_Name")%></b></font></a></td>
<td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=mySet_tb_Folders.fields("Folder_Short_Description")%></font></td>
<%
if mySet_tb_Folders.fields("Folder_Public") = 1 Then 
%>
<td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=my_File_Message_Pulic_list%></font></td>
<%

else
%>
 <td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=my_File_Message_Private%></font></td>
<% 
end if
%>
<td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>">
<%
if len(mySet_tb_Folders.fields("Last_File_Date")) > 0 Then 
 response.write myDAte_Display(mySet_tb_Folders.fields("Last_File_Date"),2)
else 
 response.write my_File_Message_None
end if
%>

</font></td>


<td align=right>&nbsp;&nbsp;<a href="__Folders_Information.asp?ID=<%=mySet_tb_Folders.fields("Folder_ID")%>"><img src="Images/overapps-info.gif" border=0></a>&nbsp;
<%
'TEST IF USER CAN MODIFY THIS FOLDER
if  mySet_tb_Folders.fields("Folder_Creator_ID") = myUser_ID OR mySet_tb_Folders.fields("Folder_Responsible_ID") = myUser_ID OR myUser_Type_ID = 1 Then
%>
&nbsp;<a href="__Folders_Modification.asp?ID=<%=mySet_tb_Folders.fields("Folder_ID")%>"><img src="images/overapps-update.gif" border=0></a>&nbsp;
<%
end if
%>
</td>

</tr>
<%
end if
mySet_tb_Folders.movenext
loop
 
' Close Recordset
mySet_tb_Folders.close
Set mySet_tb_Folders=Nothing
' Close Connection 	
myConnection.Close
set myConnection = Nothing
%> 

</table>

<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"><b><%=myMessage_Page%>(S) :&nbsp; 
<%
myNbrPage=int((i+myMaxRspByPage-1)/myMaxrspbyPage)
indice=1
do While not indice>myNbrPage 
	if CInt(indice)=CInt(myNumPage) then
		%>
		[<%=indice%>]&nbsp; 
		<%
	else
		%>
		<a href="__Folders_List.asp?page=<%=indice%>&search=<%=mySearch%>&order=<%=myOrder%>"><Font Color="<%=myApplicationTextColor%>">[<%=indice%>]</FONT></a>&nbsp;
	<%
	end if	
	indice=indice+1
loop
%>
&nbsp;</b></font></TD></TR> 
<TR><TD bgcolor="<%=myBGColor%>">
&nbsp;<a href="__Folders_Modification.asp?myAction=New"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Add%></font></a>
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
'				    End Copyright												'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 
</body>
</html>

<html><script language="JavaScript"></script></html>