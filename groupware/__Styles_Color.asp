<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Styles_Color.asp" is free software; 
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
' Name 			: __Styles_Color.asp  
' Path   		: /
' Version 		: 1.15.0
' Description 	        : Display of color for style 
' By			: Dania TCHERKEZOFF
' Company		: OverApps
' Date			: October 24, 2001
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Styles_Color.asp"

Dim myPage_Application
myPage_Application="Styles"
	
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

Dim mySearch, myMaxRspByPage, mySortStyleDirectory_ID, myStyleDirectory_ID, myStyleDirectory_Name, mySortStyle_Name, mySortStyle_Description_Short, myOrder, myRs,  myNumPage, myNbrPage, indice, myStyle_URL, myInfo, myModif

Dim mySQL_Select_tb_Styles_Color,mySet_tb_Styles_Colo,mySet_tb_Styles_Color,myRedirection, myStyleError


Dim i, j, myAction, myParameter, myColor, Style_Update_Parameter, mySet_tb_Styles_update, myUpdate_Style_Name, myMessage_Color 


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET PARAMETER								      '				
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

myParameter = Request.QueryString("myParameter")
myStyle_ID=Request.queryString("myStyle_ID")
myColor=Request.QueryString("myColor")
myAction=Request.QueryString("action")
myStyleError = Request.QueryString("myStyleError")


if myParameter = "BGColor" Then myMessage_Color=myStyles_Message_BGColor
if myParameter = "BGTextColor" Then myMessage_Color=myStyles_Message_BGTextColor
if myParameter = "BorderColor" Then myMessage_Color=myStyles_Message_BorderColor
if myParameter = "BorderTextColor" Then myMessage_Color=myStyles_Message_BorderTextColor
if myParameter = "ApplicationColor" Then myMessage_Color=myStyles_Message_ApplicationColor
if myParameter = "ApplicationTextColor" Then myMessage_Color=myStyles_Message_ApplicationColor

%>

<html>

<head>
<title><%=mySite_Name%>  Styles - Modification -</title>
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

<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 

<%
' APPLICATION TITLE
%>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></TD></TR> 
</table>


<div align=center><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><b><%=myStyles_Message_ChooseColor%> : <%=myMessage_Color%></b></font></div><br>

<%
' CENTER APPLICATION

''''''''''''''''''''''''''''''''''''''''''''
'DISPLAY OF THE COLOR
'''''''''''''''''''''''''''''''''''''''''''''

'Open Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Styles_Color = "SELECT * FROM tb_Styles_Color" 
set mySet_tb_Styles_Color = myConnection.Execute(mySQL_Select_tb_Styles_Color)



if not mySet_tb_Styles_Color.bof then 
	 mySet_tb_Styles_Color.MoveFirst
end if
%>
<table width="375" border=1 align=center>
<%

   
	 
' COLOR LISTING
	
i=0
do while not mySet_tb_Styles_Color.eof 
  if ( i mod 15 )= 0 then response.write("<TR>")	
%>	
<td bgcolor="<%=mySet_tb_Styles_Color("Color_Code")%>" width=25 height=20>
<a href="__Styles_Modification.asp?action=changecolor&myStyle_id=<%=myStyle_ID%>&myColorParameter=<%=myParameter%>&myColor=<%=mySet_tb_Styles_Color("Color_Code")%>&myStyleError=<%=myStyleError%>">

<img src=images/overapps-transp.gif border=0 height=20 width=25></a></td>
<%
if ( i mod 15 ) = 14 then response.write("</TR>")	

i=i+1
mySet_tb_Styles_Color.movenext
loop

' Close Connection 
		myConnection.close
		set myConnection = nothing



%>

</table>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>">&nbsp;</font></TD></TR> 
</table>
</td>
</tr>
</table>

<!-- #include file="_borders/Down.asp" --> 

<% 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	 '
' license's compliances.							                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
<html></html>