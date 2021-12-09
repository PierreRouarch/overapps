<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program is free software; you can redistribute it and/or modify
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
' Name : __Projects_Box
' Path : /
' Description : Projects List in a box (for home page)
' By : Pierre Rouarch
' Company : Overapps	
' Date : January, 4, 2001
' Version : 1.15.0
' Modify by :
' Company :
' Date :					
' ------------------------------------------------------------
 
' DB variables
Dim  mySQL_select_tb_projects, mySet_tb_projects

' Projects Variables
Dim myProject_ID, myProject_Name


%>
<table border="0" CELLPADDING="0" CELLSPACING="0"> <TR><TD><IMG SRC="Images/OverApps-transp.gif" WIDTH="<%=myApplication_Width%>" HEIGHT="1"></td></tr> 
<%
' Project TITLE
%> <tr bgcolor="<%=myApplicationColor%>"> <td ALIGN="Center"><b><font face="Arial, Helvetica, sans-serif"  color="<%=myApplicationTextColor%>"><%=myBox_Title%> 
</font> </b></td></tr> <tr BGCOLOR="<%=myBGColor%>" ALIGN="CENTER"> <td> <%	
' Connection 
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

' Read tb_projects
mySQL_Select_tb_Projects = "SELECT tb_projects.*,  tb_sites_Members.Member_Pseudo as Project_Leader_Pseudo FROM tb_Projects INNER JOIN tb_Sites_Members on tb_Projects.Project_leader_ID=tb_Sites_Members.Member_ID WHERE tb_Projects.Site_ID ="& mySite_ID


set mySet_tb_projects = myConnection.Execute(mySQL_Select_tb_projects)
if mySet_tb_projects.eof then %> <table> <tr ALIGN="CENTER"><td><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1"><%=myMessage_No_Project%></FONT></td></tr></table><%else %> 
<form method="Get" action="__Phases_List.asp" name=""> <table color="#d6d6ad" cellpadding="10"> 
<tr> <td bgcolor="<%=myBGColor%>"> <P><select name="Project_ID" size="" tabindex="1"> 
<option value=0 selected><%=myMessage_Select%></option> 

<%
' Projects List
%> <%do while not mySet_tb_Projects.eof%> <option value=<%=mySet_tb_Projects("Project_ID")%>><%=mySet_tb_Projects("Project_Name")%></option> 
<%
	mySet_tb_Projects.MoveNext
loop
%> </select><INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Go%>"> 
</P></td></tr> </table></form><% 
end if ' Eof / not Eof
' Close Recordset and Connection
mySet_tb_projects.close
Set mySet_tb_projects = Nothing
myConnection.Close
set myConnection = Nothing
%> </td></tr> <tr ALIGN="RIGHT"><td> <A HREF="__Projects_List.asp"><FONT SIZE="1" FACE="Arial, Helvetica, sans-serif" ><%=myMessage_More%><font size="1" face="Courier New, Courier, mono">--&gt;</font> </FONT> 
</A></td></tr> </table>








<html></html>