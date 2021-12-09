<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - O_v_erA_pps - http://www.overapps.com
'
' This program "__News_list.asp" is free software; 
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
'-----------------------------------------------------------------------------
%>
<% 	Option Explicit
	Response.Buffer = true
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache"
' Doesn't Work with PWS ???
%>

<%
' ------------------------------------------------------------
' Name 			: __News_list.asp
' Path   		: /
' Version 		: 1.15.0
' Description 	: News List
' by			: Pierre Rouarch
' Company		: OverApps
' Date			: December 10, 2001
' Contributor : Dania Tcherkezoff
'
' Modify by		:
' Company		:
' Date			:
' ------------------------------------------------------------
Dim myPage
myPage="__News_List.asp"

Dim myPage_Application
myPage_Application="News"
	
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


Dim mySearch, myMaxRspByPage, mySortNew_Date, mySortNew_Title, mySortNew_Description_Short,  myOrder, myRs,  myNumPage, myNbrPage, indice,  myRole, myInfo, myModif

Dim  myNew_Site_ID, myNew_Member_ID, myNew_ID, myNew_Title, myNew_Description_Short, myNew_Description_Long, myNew_Date, myNew_Date_End

Dim myNewsWire_ID, myNewsWire_Name

Dim i, j

Dim mySQL_Select_tb_News, mySet_tb_News, mySQL_Select_tb_NewsWires, mySet_tb_NewsWires

Dim myList

%>
<html>

<head>
<title><%=mySite_Name%> - News List</title>
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

<TD WIDTH="<%=myLeft_Width%>">
<!-- #include file="_borders/Left.asp" -->
</td>

<%
' CENTER Application
%> 

<%

mySearch=Replace(Request.querystring("Search"),"'","''")

if mySearch="" then
	mySearch=Replace(Request.form("Search"),"'","''")
end if

myNumPage=Request("Page")
if Len(myNumPage)=0 then
	myNumPage=1
end if
myMaxRspByPage=10

myList=Request("List")


' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


mySortNew_Date = "<a href=""__News_List.asp?order=New_Date&Page="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Date&"</font></a>"

mySortNew_Title = "<a href=""__News_List.asp?order=New_Title&Page="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Title&"</font></a>"

mySortNew_Description_Short = "<a href=""__News_List.asp?order=New_Description_Short&Page="&myNumPage&"""><FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</font></a>"

' Get Sort MEthod
myOrder = Request.QueryString("order")
	Select case myOrder
		case "New_Title"
			mySortNew_Title = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Title&"</FONT>"
		case "New_Description_Short"
			mySortNew_Description_Short = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Presentation&"</FONT>"
		case "New_Date"
			mySortNew_Date = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Date&"</FONT>"

		case else
			myOrder="New_Date DESC"
			mySortNew_Date = "<FONT Face=""Arial,Helvetica"" size=""2"" color="&myBorderTextColor&">"&myMessage_Date&"</FONT>"
	End Select

' Read in Database - NewsWire_ID=1 in this Version


mySQL_Select_tb_News = "SELECT *,New_Description_Long FROM tb_News INNER JOIN tb_NewsWires_sites ON tb_News.NewsWire_ID=tb_NewsWires_Sites.NewsWire_ID WHERE tb_NewsWires_sites.Site_ID="&mySite_ID


if mySearch<>"" then
	mySQL_Select_tb_News=mySQL_Select_tb_News & " AND (tb_News.New_Title LIKE '%"&mySearch&"%' OR tb_News.New_Description_Short LIKE '%"&mySearch&"%')"
end if
mySQL_Select_tb_News=mySQL_Select_tb_News & " ORDER BY " & myOrder &", tb_News.New_Title"

set mySet_tb_News = myConnection.Execute(mySQL_Select_tb_News)
%>

<TD VALIGN="top" ALIGN="left" bgcolor="<%=mybgColor%>" Width="<%=myApplication_Width%>">

<TABLE WIDTH="<%=myApplication_Width%>" BORDER="0" BGCOLOR="<%=myApplicationColor%>">
 <TR>
  <TD>
  <font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>">
   <b>
    <%=myApplication_Title%>
    </b>
  </font>
  </TD>
 </TR>
</TABLE>

<br>

<table border="0" Width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0">
 <tr ALIGN="CENTER">
  <td>
   <form method="post" action="<%=myPage%>" id=form1 name=form1> 
   <input type="text" name="Search" size="30" Value="<%=mySearch%>">
   &nbsp;<INPUT TYPE="submit" NAME="Submit" VALUE="<%=myMessage_Search%>">
   </form>
  </td>
 </tr>
</table>

<BR> 

<%
'Changing Record Set Number Depending On Page Number
i=0
j=0

myRs=(myNumPage-1)*myMaxRspByPage

if not mySet_tb_News.bof then mySet_tb_News.MoveFirst

do while not mySet_tb_News.eof
 i=i+1
 mySet_tb_News.movenext
loop

if not mySet_tb_News.bof then
 mySet_tb_News.MoveFirst
 mySet_tb_News.Move(myRs)
end if

%> 

<table border="0" cellpadding="5" cellspacing="1" Width="<%=myApplication_Width%>">
 <tr>
  <td valign="top" align="left" bgcolor="<%=myBorderColor%>">
   <b>
    <font face="Arial, Helvetica, sans-serif" size="2">
	  <%=mySortNew_Date%>
	</font>
   </b>
  </td>
  <td valign="top" align="left" bgcolor="<%=myBorderColor%>">
   <b>
    <font face="Arial, Helvetica, sans-serif" size="2">
	 <%=mySortNew_Title%>
	</font>
   </b>
  </td>
  <td valign="top" align="left" bgcolor="<%=myBorderColor%>">
   <b>
    <font face="Arial, Helvetica, sans-serif" size="2">
	 <%=mySortNew_Description_Short%>
	</font>
   </b>
  </td>
  <td valign="top" align="left" bgcolor="<%=myBorderColor%>">
   <b>
    <font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>">
      <%=myMessage_More%>
	</font>
   </b>
   </td>
  </tr>

<%
'IF NO NEWS
if mySet_tb_News.eof Then 
%>
  <tr>
    <td  valign="middle" align=center colspan=4>
	 <font face="Arial, Helvetica, sans-serif" size="2">
	  <i>
	   <%=myMessage_No_News%>
	  </i>
	 </font>
	</td>
  </tr>
<%
end if	


'DISPLAY ALL NEWS	
do while not mySet_tb_News.eof AND (myMaxRspByPage>j)
	j=j+1
	myNew_Site_ID = mySet_tb_News("Site_ID")
	myNew_Member_ID = mySet_tb_News("Member_ID")
	myNewsWire_ID   = mySet_tb_News("NewsWire_ID")
	myNew_ID        = mySet_tb_News("New_ID")
	myNew_Title         = mySet_tb_News("New_Title")
	myNew_Description_Short  = mySet_tb_News("New_Description_Short")
	myNew_Description_Long  = mySet_tb_News("New_Description_Long")
	myNew_Date =  myDate_Display(mySet_tb_News("New_Date"),2)

	if len(myNew_Description_Long) <> 0 then 
		myInfo  = "<a href=""__New_Information.asp?List="&myList&"&New_ID="&myNew_ID&""">" & "<img border=""0"" src=""images/overapps-info.gif"" WIDTH=""20"" HEIGHT=""20"" " & " alt="" " & myNew_Title & """></a>"
	
	else 
		myInfo=""
	
	end if	

	myModif = "<a href=""__New_Modification.asp?Action=Update&List=" & myList&"&New_ID=" & myNew_ID & """>" 	& "<img border=""0"" src=""images/overapps-update.gif"" WIDTH=""20"" HEIGHT=""22"" " & " alt="" " & myNew_Title & """></a>"
	%> 

 <tr>
  <td align="left" valign="middle">
	<font face="Arial, Helvetica, sans-serif" size="2">
	 <strong>
	  <%=myNew_Date%>
	   <%if myNew_Date_End>myNew_Date then %> -> <%=myNew_Date_End%><%end if%>
	 </strong>
	</font>
   </td>

	<td align="left" valign="middle">
	<font face="Arial, Helvetica, sans-serif" size="2"><strong>
	  <% = myNew_Title %>
	  </strong></font>
	</td>

	<td valign="middle" align="left">
	<font face="Arial, Helvetica, sans-serif" size="2"><%=myNew_Description_Short%> </font>
	</td>

	<td align="right" valign="middle">
	<%=myInfo%>&nbsp;&nbsp;<%=myModif%>
	</td>

	</tr>

	<%	
	mySet_tb_News.movenext
loop

' Close Recordsert
mySet_tb_News.close
Set mySet_tb_News=nothing

' Close Connection
myConnection.Close
set myConnection = Nothing


%>


<tr>
<td align="left" valign="middle" colspan="4" bgcolor="<%=myApplicationColor%>">
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myApplicationTextColor%>"><b>PAGE(S)
:&nbsp; 

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
		<a href="__News_List.asp?Page=<%=indice%>&search=<%=mySearch%>&oder=<%=myOrder%>"><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myApplicationTextColor%>"><b>[<%=indice%>]</b></font></a>&nbsp;
	<%
	end if
	indice=indice+1
loop
%>
&nbsp;</b></FONT>
</td>
</tr>

</table>

<%
' ADMINISTRATION
' EveryBody Can Add An Article in this Version
%> 
<table border="0" >
<tr>
<td>
&nbsp;<font face="Arial, Helvetica, sans-serif" size="2"><a href="__New_Modification.asp?Action=New&NewsWire_ID=<%=myNewsWire_ID%>"><font face="Arial, Helvetica, sans-serif" size="2"><%=myMessage_Add%>&nbsp;<%=myMessage_Article%></font></a></font>
</td>
</tr>
</table>

</td>
</TR>
</TABLE>




<!-- #include file="_borders/Down.asp" -->

<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.														'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0">
<TR ALIGN="RIGHT">
<TD>
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</FONT>
</TD>
</TR>
</TABLE>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 

</body>
</html>
<%

%>

<html></html>
<html><script language="JavaScript"></script></html>