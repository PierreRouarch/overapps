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
' 	" Copyright (C) 2001-2002  OverApps "
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
' Cache non géré par PWS
%>
<!-- #include file="_INCLUDE/Global_Parameters.asp" -->


<%
' ------------------------------------------------------------
' Nom 		: __Accueil.asp
' Chemin    : /Accueil
' Description 	: page d'accueil d'un site
' Société 	: OverApps
' Développeur 	: Pierre Rouarch
' Version 1.15.0
'  
' ------------------------------------------------------------
' changement de Site
Dim myPage, myNewSite_ID
myPage = "../Accueil/__Accueil.asp"

myNewSite_ID = Request.form("Site_ID")

if len(myNewSite_ID)<>0 then
	mySite_ID = myNewSite_ID
	mySite_ID = CInt(mySite_ID)
	Session("Site_ID") = mySite_ID
end if 


	
%>


<!-- #include file="_INCLUDE/DB_Environment.asp" -->


<HTML><HEAD></HEAD><TITLE>OverApps - Selection d'un Projet</TITLE>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR><TD> </TD></TR> <TR><TD> <!-- #include file="_borders/Top.asp" --> 
</TD></TR> <TR><TD WIDTH="<%=myGlobal_Width%>"> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR VALIGN="TOP" ALIGN="CENTER"><TD WIDTH="<%=myLeft_Width%>"> 
<!-- Navigation !--> <!-- #include file="_borders/Left.asp" --> </TD><!-- Application !--> 
<TD> <TABLE WIDTH="468" BORDER="0" BGCOLOR="#FFFFFF" CELLPADDING="0" CELLSPACING="0"> 
<TR ALIGN="CENTER"><TD VALIGN="TOP"> <% if (mySite_ID<>0 and mySite_Projects_Open = True) or (myUser_ID<>0 and mySite_ID = 0 and myUser_Projects_Open = True) then %> 
<!-- #include file="__Projects_Box.asp" --> <% End if %> <br> <A HREF="__Projects_List.asp"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">Liste 
de tous les projets </FONT></A></TD></TR> </TABLE></TD></TR> </TABLE></TD></TR> 
<TR><TD> <!-- #include file="_borders/Down.asp" --> </TD></TR> <TR><TD> 
</TD></TR> <TR><TD> </TD></TR></Table>

</BODY>
</HTML>
