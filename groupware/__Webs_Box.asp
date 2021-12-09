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
' Name : __Webs_Box.asp
' Path : /
' Description : Search Webs Box for Home Page
' By : Pierre Rouarch	
' Company : OverApps
' Date : January, 4 2001
' Version : 1.15.0
' Modify by :
' Company :
' Date :
' ------------------------------------------------------------


%>
<table border="0"  CELLPADDING="0" CELLSPACING="0" > <TR> 
<TD><IMG SRC="Images/transp.gif" WIDTH="<%=myApplication_Width%>" HEIGHT="1"></td></tr> 
<tr VALIGN="TOP"> <td align="center" bgcolor="<%=myApplicationColor%>"><B><font face="Arial, Helvetica, sans-serif"  color="<%=myApplicationTextColor%>"><%=myBox_Title%> 
</font></b></td></tr> <tr BGCOLOR="#FFFFFF" ALIGN="CENTER"> <td bgcolor="<%=myBGColor%>"> <form method="post" action="__Webs_List.asp" id=form1 name=form1> 
<br> &nbsp; <input type="text" name="search" size="30"> &nbsp;<INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"> 
</form></td></tr> <tr><td HEIGHT="100%"></td></Tr> <tr BGCOLOR="#FFFFFF" ALIGN="RIGHT"><td bgcolor="<%=myBGColor%>"> 
<A HREF="__Webs_List.asp"><FONT SIZE="1" FACE="Arial, Helvetica, sans-serif"><%=myMessage_More%> 
<font size="1" face="Courier New, Courier, mono">--&gt;</font> </FONT> </A></td></tr> </table>


<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>