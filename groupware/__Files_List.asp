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
' Doesn't Work With PWS ?????
%>

<%
' ------------------------------------------------------------
' Name 			: __Files_list.asp
' Path   		: /
' Version 		: 1.16.0
' Description 	:  List all files from one folder or from a search result
' By			: Dania Tcherkezoff
' Company		: OverApps
' Date			: November ,15 2001
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Files_List.asp"

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


Dim mySearch, myMaxRspByPage, myImg

Dim i, j, mySet_tb_Folders_Access, mySQL_tb_Folders_Access, myFolder_ID

Dim mySQL_Select_tb_Folders, mySet_tb_Folders,myFolder_Name, myFolder_Long_Description, myFolder_Short_Description, myFolder_Creator

Dim myFoder_ID,mySQL_Select_tb_Folders_Access, myFolder_Creation_Date, myFolder_Responsible

Dim mySet_tb_Files, mySQL_Select_tb_Files, myFolder_Responsible_ID, myNumPage, myOrder,myNbrPage,indice, myRS

myFolder_ID = Request.QueryString("Folder_ID") 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHECK IF USER CAN ENTER THIS FOLDER
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
 myfolder_Responsible          = mySet_tb_Folders("Folder_Responsible_ID")
 
 
 mySet_tb_Folders.close
 set mySEt_tb_Folders = Nothing                   

 
 
 'MAX FILE PER PAGE
myMaxRspByPage=10

 
 
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
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%= my_File_Message_Folder %> : <%=myFolder_Name%></b></font></TD></TR> 
</table>
<BR> 
<div align=center><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myApplicationTextColor%>"><b><%= myFolder_Short_Description %></b></font></div>
<%
' ROW TITLES
%>
<table width="<%=myApplication_Width%>" border=0>
<tr bgcolor=<%=myBorderColor%>>
<td><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> &nbsp;&nbsp;<b> <%=my_File_Message%></b></font></td>
<td><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> &nbsp;&nbsp;<b> <%=myMessage_Presentation%></b></font></td>
<td><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> &nbsp;&nbsp;<b> <%=my_File_Message_Type%></b></font></td>
<td width=60><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> &nbsp;&nbsp;<b> <%=my_File_Message_Size%></b></font></td>
<td width=80><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> &nbsp;&nbsp;<b> <%=myMessage_More%></b></font></td>
</tr>
<%
'GET FILES INFORMATIONS


myNumPage=Request.QueryString("Page")
if Len(myNumPage)=0 then 
	myNumPage=1
end if

myOrder = Request.QueryString("order")


mySQL_Select_tb_Files = "Select * from tb_Files where File_Folder_ID = " & myFolder_ID  & " order by File_Name"
set mySet_tb_Files = Server.CreateObject("ADODB.RecordSet")
set mySet_tb_Files = myConnection.Execute(mySQL_Select_tb_Files)




IF mySet_tb_Files.eof Then %> 
 </table>
 <br>
 <div align=center><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><i><%= myFile_Message_Empty %></i></font></div>
<%
else 


' Go to the current record
i=0
myRs=(myNumPage-1)*myMaxRspByPage
j=0
if not mySet_tb_Files.bof then mySet_tb_Files.MoveFirst
do while not mySet_tb_Files.eof 
i=i+1
mySet_tb_Files.movenext
loop 
if not mySet_tb_Files.bof then 
mySet_tb_Files.MoveFirst
mySet_tb_Files.Move(myRs) 
end if




 Do while not mySet_tb_Files.eof %>
 <tr bgcolor=<%=myBGColor%> >
<td valign=middle>
<%
myImg=0

'LIST OF ICON , CAN ADD, MODIFY AS YOU LIKE

 IF mySEt_tb_Files.Fields("File_Type") =  "JPEG Image" OR mySEt_tb_Files.Fields("File_Type") =  "Gif Image" OR mySEt_tb_Files.Fields("File_Type") =  "PNG Image" Then %>
 <img src="Images/Overapps-Files_Image.gif" border=0 width=20 heigth=20 align="absmiddle">
<%
myImg = 1
end if%>

<% IF mySEt_tb_Files.Fields("File_Type") =  "Zip Archive" OR mySEt_tb_Files.Fields("File_Type") =  "Rar Archive"  Then %>
 <img src="Images/Overapps-Files_Zip.gif" border=0 width=20 heigth=20 align="absmiddle">
<%
myImg = 1
end if%>

<% IF mySEt_tb_Files.Fields("File_Type") =  "MS Word Document"  OR mySEt_tb_Files.Fields("File_Type") =  "Active Server Page" OR mySEt_tb_Files.Fields("File_Type") =  "XML" OR mySEt_tb_Files.Fields("File_Type") =  "Acrobat Reader Document" OR mySEt_tb_Files.Fields("File_Type") =  "Text"  OR mySEt_tb_Files.Fields("File_Type") =  "Web Document" Then %>
 <img src="Images/Overapps-Files_Text.gif" border=0 width=20 heigth=20 align="absmiddle">
<%
myImg = 1
end if%>

<% IF mySEt_tb_Files.Fields("File_Type") =  "MS Excel Document"   Then %>
 <img src="Images/Overapps-Files_Excel.gif" border=0 width=20 heigth=20 align="absmiddle">
<%
myImg = 1
end if%>

<% IF mySEt_tb_Files.Fields("File_Type") =  "Audio"   Then %>
 <img src="Images/Overapps-Files_Audio.gif" border=0 width=20 heigth=20 align="absmiddle">
<%
myImg = 1
end if%>

<% IF mySEt_tb_Files.Fields("File_Type") =  "Program"   Then %>
 <img src="Images/Overapps-Files_Exe.gif" border=0 width=20 heigth=20 align="absmiddle">
<%
myImg = 1
end if%>



<% IF myImg = 0   Then %>
 <img src="Images/Overapps-Files_Unknow.gif" border=0 width=20 heigth=20 align="absmiddle">
<%
myImg = 1
end if%>


&nbsp;<a target=_blank href="__Files_Download.asp?file_ID=<%=mySet_tb_Files.fields("File_ID")%>&Folder_ID=<%= myFolder_ID %>" ><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><b><%=mySet_tb_Files.fields("File_Name")%></b></font></a></td>
<td  valign=middle><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"> &nbsp; <%=mySet_tb_Files.fields("File_Short_Description")%></font></td>
<td  valign=middle><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"> &nbsp; <%=mySet_tb_Files.fields("File_Type")%></font></td>
<td  valign=middle align=right width=60><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"> 
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
<td  valign=middle width=80><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"> &nbsp;&nbsp;
&nbsp;&nbsp;<a href="__Files_Information.asp?File_ID=<%=mySet_tb_Files.fields("File_ID")%>&Folder_ID=<%= myFolder_ID %>"><img src="Images/overapps-info.gif" border=0></a>&nbsp;
<%
if  mySet_tb_Files.fields("File_Creator_ID") = myUser_ID or mySet_tb_Files.fields("File_Responsible_ID") =myUser_ID or myUser_TYPE_ID = 1 Then
%> 
&nbsp;<a href="__Files_Modification.asp?File_ID=<%=mySet_tb_Files.fields("File_ID")%>&Folder_ID=<%= myFolder_ID %>"><img src="images/overapps-update.gif" border=0></a>&nbsp;
<%end if%></font>
</td>
</tr>
<%
 mySet_tb_Files.MoveNext
Loop
%>
</table>
<%
end if
%>

<br>
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
&nbsp;<a href="__Files_Modification.asp?myAction=New&Folder_ID=<%=myFolder_ID%>"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Add%></font></a>
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
<html><script language="JavaScript"></script></html>