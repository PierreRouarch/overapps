<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001  -+ OverApps - http://www.overapps.com
'
' This program "Left.asp" is free software; you can redistribute it and/or modify
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
' 	" Copyright (C) 2001 OverApps & contributors "
' at the bottom of the page with an active link from the name "OverApps" to 
' the Address http://www.overapps.com where the user (netsurfer) could find the 
' appropriate information about this license. 
'-----------------------------------------------------------------------------
%>
<%
' ------------------------------------------------------------
' Nom : Left.asp
' Path : /_borders
' Description : Navigation Menu
' Company : OverApps
' By : Pierre Rouarch
' Update : May, 9, 2001

' ------------------------------------------------------------

%>
<TABLE  BORDER="0" CELLPADDING="0" CELLSPACING="0" BGCOLOR="<%=myBorderColor%>" HEIGHT="100%" Width=100%> 

<TR ALIGN="RIGHT"><TD><img src="Images/transp.gif" width="1" height="10" align="absmiddle"></TD><TD><img src="Images/transp.gif" width="3" height="1" align="absmiddle"></TD></TR> 

<%

myNumber_User_Applications=0
Do While myNumber_User_Applications<=myMax_User_Applications

%>

<TR ALIGN="RIGHT"><TD><B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><A HREF="<%=myUser_Application_Entry_Page(myNumber_User_Applications)%>" ><FONT COLOR="<%=myBorderTextColor%>" STYLE="text-decoration:none"><%=myUser_Application_Title(myNumber_User_Applications)%> 
</FONT></A></FONT></B> </TD><TD><img src="Images/transp.gif" width="3" height="1" align="absmiddle"></TD></TR> <% 
myNumber_User_Applications=myNumber_User_Applications+1

Loop 
%>


 

 <TR ALIGN="RIGHT"><TD><img src="Images/transp.gif" width="1" height="10" align="absmiddle"></TD><TD><img src="Images/transp.gif" width="3" height="1" align="absmiddle"></TD></TR> 


<TR ALIGN="RIGHT"><TD><img src="Images/transp.gif" width="1" height="10" align="absmiddle"></TD><TD><img src="Images/transp.gif" width="3" height="1" align="absmiddle"></TD></TR> 

<%if myUser_type_ID=7 then %>
<TR ALIGN="RIGHT"><TD ><B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><A HREF="__identification_Site.asp"> 
<FONT COLOR="<%=myBorderTextColor%>" STYLE="text-decoration:none"><%=myMessage_Identification%> </FONT></A></FONT></B></TD><TD><img src="Images/transp.gif" width="3" height="1" align="absmiddle"></TD></TR>
<%end if%>


 
<TR ALIGN="RIGHT"><TD HEIGHT="100%" ><B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"></FONT></B></TD><TD><img src="Images/transp.gif" width="3" height="1" align="absmiddle"></TD></TR> 
</TABLE>