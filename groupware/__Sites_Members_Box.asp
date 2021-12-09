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
'-----------------------------------------------------------------------------
%> 
<%
' ------------------------------------------------------------
' Name : __Sites_Members_Box.asp
' Path : /
' Description : Sites Members Search Box for Home Page
' By : Pierre Rouarch	
' Company : OverApps
' Date : September 12, 2001
' Version : 1.15.0
' Modify by :
' Company :
' Date :
' ------------------------------------------------------------
%> 
<table border="0" CELLPADDING="0" CELLSPACING="0">

<%
' Space
%>

<TR>
<TD><IMG SRC="Images/OverApps-transp.gif" WIDTH="<%=myApplication_Width%>" HEIGHT="1"></td>
</tr> 

<%
' BOX Title
%>



<tr>
<td align="center" bgcolor="<%=myApplicationColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="3" color="<%=myApplicationTextColor%>"><%=myBox_Title%></font></b>
</td>
</tr> 

<%
' Search
%>

<tr BGCOLOR="<%=myBGColor%>" ALIGN="CENTER">
<td>
<form method="post" action="__Sites_Members_list.asp" id=form1 name=form1> 
<br> &nbsp;<input type="text" name="search" size="30"> &nbsp; <INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"> 
</form>
</td>
</tr>
 
<%
' More 
%>

<tr>
<td HEIGHT="100%">
</td>
</Tr>
<tr BGCOLOR="<%=myBGColor%>"> 
<td ALIGN="right"> 
<A HREF="__Sites_Members_List.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBGTextColor%>"><%=myMessage_More%><font size="1" face="Courier New, Courier, mono">--&gt;</font></FONT></A>
</td>
</tr>

</table>

<html></html>