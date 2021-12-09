<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   nnnn OverApps nnnnn http://www.overapps.com
'
' This program "__News_Box" is free software; you can redistribute it and/or modify
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
'-----------------------------------------------------------------------------
%>
<%
' ------------------------------------------------------------
' Name : __News_Box.asp
' Path : /
' Description : Top 10 News Box for Home page
' By : Pierre Rouarch	
' Company : OverApps
' Date : December, 10,2001
' Version : 1.15.0
' Contributor : Dania Tcherekezoff
'
' Modify by :
' Company :
' Date :
' ------------------------------------------------------------
' DB Variables

Dim  mySQL_select_tb_News, mySet_tb_News, mySQL_select_tb_NewsWires, mySet_tb_NewsWires


' News Variables
Dim myNew_ID, myNew_Title, myNew_Description_Short, myNew_Description_Long, myNew_Date, myNewsWire_ID, myNewsWire_Name




%>
<table border="0" CELLPADDING="0" CELLSPACING="0" >
  <TR> 
    <TD colspan="5"><IMG SRC="Images/OverApps-transp.gif" WIDTH="<%=myApplication_Width%>" HEIGHT="1"></td>
  </tr>
  <tr> 
    <td align="center" bgcolor="<%=myApplicationColor%>" colspan="5"><B><font face="Arial, Helvetica, sans-serif"  color="<%=myApplicationTextColor%>"><%=myBox_Title%></font></b></td>
  </tr>
  <tr BGCOLOR="<%=myBGColor%>" ALIGN="CENTER"> 
    <td colspan="5"> 
      <%	
' Connection
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

' Select Top 10 news in newswires site selection
mySQL_Select_tb_News = "SELECT  TOP 10 * FROM tb_News INNER JOIN tb_NewsWires_Sites ON tb_News.NewsWire_ID=tb_Newswires_Sites.NewsWire_ID WHERE tb_NewsWires_Sites.Site_ID="&mySite_ID 

' Beginning Today
mySQL_Select_tb_News = mySQL_Select_tb_News &" ORDER BY tb_News.New_Date DESC"



	set mySet_tb_News = 	myConnection.Execute(mySQL_Select_tb_News)
if mySet_tb_News.eof then %>
  <tr> 
    <td bgcolor="<%=myBGColor%>" colspan="4"><div align=center><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" color="<%= myBGTextColor %>"><%=myMessage_No_News%></FONT></div></td>
  </tr>
  <%else  

	do while not mySet_tb_News.eof 
	myNew_ID = mySet_tb_News("New_ID")
	myNewsWire_ID = mySet_tb_News("NewsWire_ID")
	myNew_Title = mySet_tb_News("New_Title")
	myNew_Description_Short = mySet_tb_News("New_Description_Short")
	
	myNew_Date = myDate_Display(mySet_tb_News("New_Date"),2)
	
' Get Source Name
mySQL_Select_tb_NewsWires = "SELECT * FROM tb_NewsWires WHERE NewsWire_ID="&myNewsWire_ID 
	set mySet_tb_NewsWires = myConnection.Execute(mySQL_Select_tb_NewsWires)
	myNewsWire_Name = mySet_tb_NewsWires("Newswire_Name")
	%>
  <TR> 
    <TD bgcolor="<%=myBGColor%>" ><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1"> &nbsp;<%=myNew_Date%></font> </td> 
    <TD bgcolor="<%=myBGColor%>" ><font face="Arial, Helvetica, sans-serif" size="1"><b> <%=myNew_Title%></b></font>&nbsp;</td>
    <TD bgcolor="<%=myBGColor%>" align=left ><font face="Arial, Helvetica, sans-serif" size="1"><%=myNew_Description_Short%></font>
  	
      
	   <% if myNew_Description_Long <>"" then %>
	   &nbsp;<a href="__New_Information.asp?New_ID=<%=myNew_ID%>"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1"><%=myMessage_More%> <font size="1" face="Courier New, Courier, mono">--&gt;</font> </font></a>
	   <%else%>
	   &nbsp;  
	   <% end if %>
	   </td>
    </tr>
  <%	
    	mySet_tb_News.movenext
	loop 

end if 	
' Close Recordset and connection
mySet_tb_News.close
Set mySet_tb_News = Nothing
myConnection.Close
set myConnection = Nothing


%>
  <tr BGCOLOR="#FFFFFF" ALIGN="RIGHT"> 
    <td bgcolor="<%=myBGColor%>" colspan="5"> <A href="__News_List.asp"><FONT SIZE="1" FACE="Arial, Helvetica, sans-serif" ><%=myMessage_More%> 
      <font size="1" face="Courier New, Courier, mono">--&gt;</font> </FONT> </A></td>
  </tr>
</table>




