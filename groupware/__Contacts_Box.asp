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
' Name : __Contacts_Box.asp
' Path : /
' Description : Search Contact Box for Home Page
' By : Pierre Rouarch	
' Company : OverApps
' Date : December,10, 2001
' Versions:  1.15.0
' Contributor : Dania Tcherkezoff
' Modify by :
' Company :
' Date :
' ------------------------------------------------------------
%> <table border="0"  CELLPADDING="0" CELLSPACING="0" > <TR> <TD colspan="2" ><IMG SRC="Images/OverApps-transp.gif" WIDTH="<%=myApplication_Width%>" HEIGHT="1"></td></tr> 
<%
' Contacts TITLE
%> <tr bgcolor="<%=myApplicationColor%>"> <td ALIGN="Center" colspan=2 ><b><font face="Arial, Helvetica, sans-serif"  color="<%=myApplicationTextColor%>"><%=myBox_Title%> 
</font> </b></td></tr> <form method="post" action="__Contacts_List.asp" id=form1 name=form1> 
<tr BGCOLOR="#FFFFFF"> <td bgcolor="<%=myBGColor%>" > <font face="Arial,Helvetica" size=2><b><FONT SIZE="1"><%=myMessage_Company%> 
:</FONT></b></font> </td><td bgcolor="<%=myBGColor%>"> <input type="text" name="Contact_Company" size="20"> 
</td></tr> 


<tr BGCOLOR="#FFFFFF"> <td bgcolor="<%=myBGColor%>"> <font face="Arial,Helvetica" size=2><b><FONT SIZE="1"><%=myMessage_Name%> 
:</FONT> </b></font> </td><td bgcolor="<%=myBGColor%>"> <input type="text" name="Contact_Name" size="20"> 
</td></tr> <tr BGCOLOR="#FFFFFF"> <td bgcolor="<%=myBGColor%>"> <font face="Arial,Helvetica" size=2><b><FONT SIZE="1"><%=myMessage_City%> 
:</FONT> </b></font> </td><td bgcolor="<%=myBGColor%>"> <input type="text" name="Contact_City" size="20"> 
</td></tr> <trbgcolor="<%=myBGColor%>"> <td ALIGN="RIGHT" bgcolor="<%=myBGColor%>"> <FONT FACE="Arial, Helvetica, sans-serif" SIZE="1"><%=myMessage_You_Search_For%> 
: <br> <%=myMessage_A_Professionnal%><INPUT TYPE="radio" NAME="Method_Search" VALUE="Company" CHECKED><BR> 
<%=myMessage_An_Individual%><INPUT TYPE="radio" NAME="Method_Search" VALUE="Home"></FONT> 
</td><td ALIGN="CENTER" bgcolor="<%=myBGColor%>"><INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"></td></tr> 
</form>
<tr><td HEIGHT="100%" colspan="2"></td></Tr> <tr BGCOLOR="#FFFFFF"> <td colspan="2" ALIGN="right" bgcolor="<%=myBGColor%>"> 
<A HREF="__Contacts_List.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" ><%=myMessage_More%> 
<font size="1" face="Courier New, Courier, mono">--&gt;</font> </FONT></A></td></tr> </table>
<html><script language="JavaScript"></script></html>