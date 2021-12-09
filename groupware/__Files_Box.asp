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

<%
' ------------------------------------------------------------
' Name 			: __Files_Boxt.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Folders Box for home page
' By			: Dania Tcherkezoff
' Company		: OverApps
' Date			: December ,10 2001
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------



Dim mySearch, myMaxRspByPage

Dim mySet_tb_Folders_Access, mySQL_tb_Folders_Access

Dim mySQL_Select_tb_Folders, mySet_tb_Folders

Dim myFolder_ID, myCount

%>

<table border="0" CELLPADDING="0" CELLSPACING="0" width="<%=myApplication_Width%>">
  <TR> 
    <TD colspan="5"><IMG SRC="Images/OverApps-transp.gif" WIDTH="<%=myApplication_Width%>" HEIGHT="1"></td>
  </tr>
  <tr> 
    <td align="center" bgcolor="<%=myApplicationColor%>" colspan="5"><B><font face="Arial, Helvetica, sans-serif"  color="<%=myApplicationTextColor%>"><%=myBox_Title%></font></b></td>
  </tr>
  
 </table> 







<%
' LISTING

' dbConnection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

		
' Get Folders Informations
mySQL_Select_tb_Folders = "SELECT Max(tb_Files.File_Modification_Date) as Last_File_Date ,Folder_Responsible_ID,Folder_Creator_ID,tb_Folders.Folder_Name, tb_Folders.Folder_Short_Description, Count(tb_Files.File_ID) AS NBFile, tb_Folders.Folder_Public, tb_Folders.Folder_ID FROM tb_Folders INNER  JOIN tb_Files ON tb_Folders.Folder_ID = tb_Files.File_Folder_ID  GROUP BY tb_Folders.Folder_Name, tb_Folders.Folder_Public, tb_Folders.Folder_ID, tb_Folders.Folder_Short_Description,Folder_Creator_ID,Folder_Responsible_ID ORDER BY tb_Folders.Folder_Name;"


' Execute
set mySet_tb_Folders = Server.CreateObject("ADODB.RecordSet")
set mySet_tb_Folders = myConnection.Execute(mySQL_Select_tb_Folders)

%>



<%
' Go to the current record

%> 

<%
' ROW TITLES
%>
<table width="<%=myApplication_Width%>" border=0 CELLPADDING="0" CELLSPACING="0"> 

<%
' LISTING
%>

<%	
mycount= 0 

do while not mySet_tb_Folders.eof and myCount < 8 


'TEST ID USER HAS ACCESS
 mySQL_tb_Folders_Access = "Select * from tb_Folders_Access where Folder_ID = " & mySet_tb_Folders.fields("Folder_ID") & " AND Member_ID = "& myUser_ID &" AND Site_ID=" & mySite_ID  
 set mySet_tb_Folders_Access = myConnection.Execute(mySQL_tb_Folders_Access)

if mySet_tb_Folders_Access.eof AND mySet_tb_Folders.fields("Folder_Public") = 0 AND mySet_tb_Folders.fields("Folder_Creator_ID") <> myUser_ID AND mySet_tb_Folders.fields("Folder_Responsible_ID") <> myUser_ID AND myUser_Type_ID > 1 Then 

else 
%>

<tr>
<td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="1" color="<%=myBGTextColor%>">
<%
if len(mySet_tb_Folders.fields("Last_File_Date")) > 0 Then 
 response.write myDate_Display(mySet_tb_Folders.fields("Last_File_Date"),2)
else 
 response.write my_File_Message_None
end if
%>

</font></td>

<td>&nbsp;&nbsp;
<a href="__Files_List.asp?Folder_ID=<%= mySet_tb_Folders.fields("Folder_ID") %>">
<img src="Images/Files_Folder.gif" alt="Explore" border=0 width=12 heith=12></a>
&nbsp;<a href="__Files_List.asp?Folder_ID=<%= mySet_tb_Folders.fields("Folder_ID") %>"><font face="Arial, Helvetica, sans-serif" size="1" color="<%=myBGTextColor%>"><b><%=mySet_tb_Folders.fields("Folder_Name")%></b></font></a></td>






<td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="1" color="<%=myBGTextColor%>"><%=mySet_tb_Folders.fields("Folder_Short_Description")%></td>

<td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="1" color="<%=myBGTextColor%>">
<%=mySet_tb_Folders.fields("NBFile")  &" " & my_File_Message_Files%> 
</font></td>

<%
if mySet_tb_Folders.fields("Folder_Public") = 1 Then 
%>
<td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="1" color="<%=myBGTextColor%>"><%=my_File_Message_Pulic_list%></td>
<%

else
%>
 <td>&nbsp;&nbsp;<font face="Arial, Helvetica, sans-serif" size="1" color="<%=myBGTextColor%>"><%=my_File_Message_Private%></td>
<% 
end if
%>
</tr>
<%
end if
myCount = mycount  + 1 
mySet_tb_Folders.movenext
loop
 
' Close Recordset
mySet_tb_Folders.close
Set mySet_tb_Folders=Nothing
' Close Connection 	
myConnection.Close
set myConnection = Nothing
%> 

  <tr BGCOLOR="#FFFFFF" ALIGN="RIGHT"> 
    <td bgcolor="<%=myBGColor%>" colspan="5" align=right><A href="__Folders_List.asp"><FONT SIZE="1" FACE="Arial, Helvetica, sans-serif" ><%=myMessage_More%> 
     <font size="1" face="Courier New, Courier, mono">--&gt;</font></FONT></A></td>
  </tr>

</table>


<html><script language="JavaScript"></script></html>