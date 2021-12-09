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
' Doesn't Work With PWS ????
%>

<%
' ------------------------------------------------------------
' Name 			: __Styles_list.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: Styles List
' By			: Pierre Rouarch
' Company		: OverApps
' Date			: February 1, 2001
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage, myNewSite_ID
myPage = "__Styles_List.asp"

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

Dim mySearch, myMaxRspByPage, mySortStyleDirectory_ID, myStyleDirectory_ID, myStyleDirectory_Name, mySortStyle_Name, mySortStyle_Description_Short, myOrder, myRs,  myNumPage, myNbrPage, indice, myStyle_URL,  myInfo, myModif


Dim i, j

Dim myList_Styles_Name, myList_Styles_Global_Width , myList_Styles_Left_Width , myList_Styles_Right_Width, myList_Styles_Application_Width, myList_Styles_BGColor, myList_Styles_BGImage, myList_Styles_BGTextColor, myList_Styles_BorderColor, myList_Styles_BorderImage, myList_Styles_BorderTextColor, myList_Styles_ApplicationColor, myList_Styles_ApplicationImage, myList_Styles_ApplicationTextColor, myList_Styles_Date_Update, myList_Styles_Author_Update,myList_Styles_ID


%>
<html>

<head>
<title><%=mySite_Name%>  Styles - List -</title>
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
' Get Parameters

mySearch=Replace(Request.querystring("Search"),"'","''")
if mySearch="" then
		mySearch=Replace(Request.form("Search"),"'","''")
end if

myNumPage=Request("Page")
if Len(myNumPage)=0 then 
	myNumPage=1
end if

myOrder = Request.QueryString("order")
		
myMaxRspByPage=3



' Prepare Sort 

mySortStyle_Name = "<a href=""__Styles_List.asp?order=Style_Name&Page="&myNumPage&"&search="&mySearch&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Style&"</font></a>"

mySortStyle_Description_Short = "<a href=""__Styles_List.asp?order=Style_Description_Short&Page="&myNumPage&"&search="&mySearch&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</font></a>"

Select case myOrder
	case "Style_Name"
		mySortStyle_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Style&"</FONT>"
	case "Style_Description_Short"
		mySortStyle_Description_Short = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</FONT>"
	case else
		myOrder="Style_Name"
		mySortStyle_Name = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Style&"</FONT>"
End Select

' dbConnection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

		
' Get Styles Informations
mySQL_Select_tb_Styles = "SELECT * FROM tb_Styles WHERE tb_Styles.Site_ID="&mySite_ID

' Search Purposes
if mySearch<>"" then
	mySQL_Select_tb_Styles=mySQL_Select_tb_Styles & " AND (Style_Name LIKE '%"&mySearch&"%')"
end if


' ORDER
If myOrder <> "Style_Name" Then
 mySQL_Select_tb_Styles=mySQL_Select_tb_Styles & " ORDER BY " & myOrder &", Style_Name"
else  
 mySQL_Select_tb_Styles=mySQL_Select_tb_Styles & " ORDER BY Style_Name"
end if 
 
' Execute
set mySet_tb_Styles = myConnection.Execute(mySQL_Select_tb_Styles)

%> 



<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 

<%
' APPLICATION TITLE
%>

<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></TD></TR> 
</table>

<%
' SEARCH BOX
%>


<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0"> 
<tr ALIGN="CENTER">

<td> 
<form method="post" action="__Styles_List.asp" id=form1 name=form1> 
<br> &nbsp; <input type="text" name="search" size="30" VALUE="<%=mySearch%>"> &nbsp; <INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>"> 
</form>
</td>
</tr>
</table>


<BR> 

<%
' LIST
%>



<%
' Go to the current record
i=0
myRs=(myNumPage-1)*myMaxRspByPage
j=0
if not mySet_tb_Styles.bof then mySet_tb_Styles.MoveFirst
do while not mySet_tb_Styles.eof 
i=i+1
mySet_tb_Styles.movenext
loop 
if not mySet_tb_Styles.bof then 
mySet_tb_Styles.MoveFirst
mySet_tb_Styles.Move(myRs) 
end if
%> 

<%
' ROW TITLES
%>

<%
' LISTING
%>

<%	
do while not mySet_tb_Styles.eof AND (myMaxRspByPage>j)

	j=j+1
	myList_Styles_ID           = mySet_tb_Styles("Style_ID")
	myList_Styles_Name         = mySet_tb_Styles("Style_Name")
	myList_Styles_Global_Width = mySet_tb_Styles("Style_Global_Width")
	myList_Styles_Left_Width	 = mySet_tb_Styles("Style_Left_Width")
	myList_Styles_Right_Width	 = mySet_tb_Styles("Style_Right_Width")
	myList_Styles_Application_Width= mySet_tb_Styles("Style_Application_Width")
	myList_Styles_BGColor		 = mySet_tb_Styles("Style_BGColor")
	myList_Styles_BGImage		 = mySet_tb_Styles("Style_BGImage")
	myList_Styles_BGTextColor  = mySet_tb_Styles("Style_BGTextColor")
	myList_Styles_BorderColor  = mySet_tb_Styles("Style_BorderColor")
	myList_Styles_BorderImage  = mySet_tb_Styles("Style_BorderImage")
	myList_Styles_BorderTextColor = mySet_tb_Styles("Style_BorderTextColor")
	myList_Styles_ApplicationColor= mySet_tb_Styles("Style_ApplicationColor")
	myList_Styles_ApplicationImage= mySet_tb_Styles("Style_ApplicationImage")
	myList_Styles_ApplicationTextColor = mySet_tb_Styles("Style_ApplicationTextColor")
	myList_Styles_Date_Update  = mySet_tb_Styles("Style_Date_Update")
	myList_Styles_Author_Update = mySet_tb_Styles("Style_Author_Update")


		
	myInfo  = "<a href=""__Style_Information.asp?Style_ID=" & myStyle_ID & """>" & "<img border=""0"" src=""images/overapps-info.gif"" WIDTH=""20"" HEIGHT=""20"" " & " alt="" " & myStyle_Name & """></a>"
	myModif = "<a href=""__Style_Modification.asp?Style_ID=" & myStyle_ID & """>" 	& "<img border=""0"" src=""images/overapps-update.gif"" WIDTH=""20"" HEIGHT=""22"" " & " alt="" " & myStyle_Name & """></a>"
	%> 

<table border="0" Width="300" cellpadding="1" cellspacing="1" align=center>
<TR><TD bgcolor="<%=myList_Styles_ApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4"  color="<%=myList_Styles_ApplicationTextColor%>"><b><%=myList_Styles_Name%></b></font></TD></TR> 

<tr> 
<td valign="top" align="left" bgcolor="<%=myList_Styles_BorderColor%>">
<b><font face="Arial, Helvetica, sans-serif" size="2"  color="<%=myList_Styles_BorderTextColor%>"><%=myStyles_Message_Border%></font></b>
</td></tr>
<tr><td bgcolor="<%=myList_Styles_BGColor%>">
<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myList_Styles_BGTextColor%>"><strong>
<%=myStyles_Message_Global_Width%> : <%=myList_Styles_Global_Width%><br>
<%=myStyles_Message_Left_Width%> : <%=myList_Styles_Left_Width%><br>
<%=myStyles_Message_Application_Width%> : <%=myList_Styles_Application_Width%></td></tr>
<tr><td bgcolor="<%=myBorderColor%>"><a href="__Styles_Modification.asp?myStyle_id_selected=<%=myList_Styles_ID%>"><font face="Arial, Helvetica, sans-serif" size="2"  COLOR="<%=myBorderTextColor%>"><%=myStyles_Message_Modify%></font></a>
</td></tr>
</table>

<br>
<br>

<%
mySet_tb_Styles.movenext
loop
 
' Close Recordset
mySet_tb_Styles.close
Set mySet_tb_Styles=Nothing
' Close Connection 	
myConnection.Close
set myConnection = Nothing

%> 
<table  border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" >
      <tr>
<td align="left" valign="middle" colspan="3" bgcolor="<%=myApplicationColor%>"> 
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myApplicationTextColor%>"><b><%=myMessage_Page%>(s) :&nbsp; 
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
		<a href="__Styles_List.asp?page=<%=indice%>&search=<%=mySearch%>&order=<%=myOrder%>"><Font Color="<%=myApplicationTextColor%>">[<%=indice%>]</FONT></a>&nbsp;
	<%
	end if	
	indice=indice+1
loop
%>
&nbsp;</b></FONT>
</td>
</tr>
<tr><td><a href=__Styles_Modification.asp?myAction=Add><font face="Arial, Helvetica, sans-serif" size="2"  ><%=myStyles_Message_Add%></a></td></tr>
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


<html></html>
<html><script language="JavaScript"></script></html>