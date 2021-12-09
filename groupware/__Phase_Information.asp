<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - O-v-e-r=A+p+p+s+ - http://www.overapps.com
'
' This program "__Phase_Information.asp" is free software; 
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
' Does n't Work With PWS ????
%>

<%
' ------------------------------------------------------------
' Name			: __Phase_Information.asp
' PAth   		: /
' Version 		: 1.15.0
' Description 	: Information about a planning Phase
' By		 	: Pierre Rouarch											
' Company		: OverApps
' Date			: October, 4, 2001
' Contributions	:
'
' Modify by		: 
' Company 		: 
' Date			: 
' ------------------------------------------------------------

' Page
Dim myPage
myPage = "__Phase_Information.asp"
Dim myPage_Application
myPage_Application="Projects"
	
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

' Variables

Dim  mySQL_Select_tb_Phases, mySet_tb_Phases, mySQL_Select_tb_Projects, mySet_tb_Projects

Dim myProject_ID, MyProject_Name, myProject_Member_ID, myProject_Leader_ID


Dim myPhase_Project_Name, myPhase_Member_ID, myPhase_ID, myPhase_Name, myPhase_Parent_ID, myPhase_Parent_Name, myPhase_Date_Beginning, myPhase_Date_End, myPhase_Date_Beginning2, myPhase_Date_End2,  myPhase_Progress,  myPhase_Presentation, myPhase_Leader_ID, myLeader_Pseudo,   myPhase_Author_Update, myPhase_Date_Update	

Dim myList, myNumPage, mySearch


' Get Parameters

myProject_ID=request("Project_ID")
if  Len(myProject_ID)=0 then
		Response.Redirect("__Project_List.asp")
end if

myPage=myPage&"?Project_ID="&myProject_ID
	
myList = ""
myList = request("List")


myPhase_ID = request("Phase_ID")
if len(myPhase_ID)=0 then 
			Response.Redirect("__Project_List.asp?Project_ID="&myProject_ID&"")
end if

myPage=myPage&"&Phase_ID="&myPhase_ID


' Read Information in DB

' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Phases   = "SELECT tb_phases.*, tb_Sites_Members.Member_Pseudo AS Leader_Pseudo, tb_phases.Phase_ID, tb_phases_1.Phase_Name AS Phase_Parent_Name FROM (tb_phases LEFT JOIN tb_Sites_Members ON tb_phases.Phase_Leader_ID = tb_Sites_Members.Member_ID)  LEFT JOIN tb_phases AS tb_phases_1 ON tb_phases.Phase_Parent_ID = tb_phases_1.Phase_ID  WHERE tb_phases.Phase_ID = " & myPhase_ID & " AND tb_phases.Project_ID = " & myProject_ID

set mySet_tb_Phases = myConnection.execute(mySQL_Select_tb_Phases)

'If Nothing GO BACK
if mySet_tb_Phases.eof then
	mySet_tb_Phases.close
	Set mySet_tb_Phases=nothing
	myConnection.close
	set myConnection = nothing
	Response.Redirect("__Phases_List.asp?Project_ID="&myProject_ID&"")

else
	' Get Information
	myPhase_Member_ID=mySet_tb_Phases("Member_ID")
	myPhase_Name    = mySet_tb_Phases("Phase_Name")
	myPhase_Parent_Name=mySet_tb_Phases("Phase_Parent_Name")
	myPhase_Presentation = mySet_tb_Phases("Phase_Presentation")
	myPhase_Date_Beginning  = mySet_tb_Phases("Phase_Date_Beginning")
	myPhase_Date_End    = mySet_tb_Phases("Phase_date_End")
	myPhase_Date_Beginning2 = mySet_tb_Phases("Phase_Date_Beginning2")
	myPhase_Date_End2   = mySet_tb_Phases("Phase_date_End2")
	myLeader_Pseudo	= mySet_tb_Phases("Leader_Pseudo")
	myPhase_Author_Update	= mySet_tb_Phases("Phase_Author_Update")
	myPhase_Date_Update= mySet_tb_Phases("Phase_Date_Update")
end if

mySet_tb_Phases.close
set mySet_tb_Phases=nothing
myConnection.close
set myConnection = nothing

'''''''''''''''''''''''''''
' PROJECT Information     '
'''''''''''''''''''''''''''
' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Projects = "Select * From tb_Projects Where Project_ID="&myProject_ID
	set mySet_tb_Projects = myConnection.execute(mySQL_Select_tb_Projects)
myProject_Member_ID=mySet_tb_Projects("Member_ID")
myProject_Leader_ID=mySet_tb_Projects("Project_Leader_ID")
myPhase_Project_Name=mySet_tb_Projects("Project_Name")

mySet_tb_Projects.close
Set mySet_tb_Projects=nothing
myConnection.close
set myConnection = nothing




%>

<HTML>
<HEAD>
</HEAD>
<TITLE><%=mySite_Name%> - Phase - Information</TITLE>

<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">


<%
' TOP
%>

<!-- #include file="_borders/Top.asp" -->


<%
' CENTER
%>

<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" CELLPADDING="0" CELLSPACING="0">

<TR VALIGN="TOP">

<%
' CENTER LEFT
%>


<TD WIDTH="<%=myLeft_Width%>" BGColor="<%=myBorderColor%>"> 
<!-- #include file="_borders/Left.asp" --> 
</TD>

<%
' CENTER APPLICATION
%>


<TD WIDTH="<%=myApplication_Width%>" ALIGN="CENTER" BGCOLOR="<%=myBGColor%>">

<table border="0" cellpadding="6" cellspacing="1" width="<%=MyApplication_Width%>"> 

<%
' Project and Phase Name
%>

<tr align="center"> 
<td colspan="2" bgcolor="<%=myApplicationColor%>">
<font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><%=myMessage_Project%>&nbsp;:&nbsp;<%=myPhase_Project_Name%>&nbsp;/&nbsp;<%=MyMessage_Phase%>&nbsp;:&nbsp;<%=myPhase_Name%></font>
</td>
</tr>

<%
' Presentation
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Presentation%></b></font>
</td>
<td align="left">
<font size="2" face="Arial, Helvetica, sans-serif"><%=myPhase_Presentation%></font>
</td>
</tr>  

<%
' Parent
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Parent%></b></font>
</td>
<td align="left">
<font size="2" face="Arial, Helvetica, sans-serif"><%=myPhase_Parent_Name %></font>
</td>
</tr>


<%
' Date Beginning
%>

<tr>
<td align="right"  bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Beginning%></b></font>
</td>
<td align="left">
<font size="2" face="Arial, Helvetica, sans-serif"><%=myDate_Display(myPhase_Date_Beginning,1)%></font>
</td>
</tr>

<%
' Date End
%>
 
<tr>
<td align="right"  bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_End%></b></font>
</td>
<td align="left">
<font size="2" face="Arial, Helvetica, sans-serif"><%=myDate_Display(myPhase_Date_End,1)%></font>
</td>
</tr> 

<%
' Revised Beginning
%>

<tr>
<td align="right"  bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Beginning%>&nbsp;(<%=myMessage_Correction%>)</b></font>
</td>
<td align="left">
<font size="2" face="Arial, Helvetica, sans-serif"><%=myDate_Display(myPhase_Date_Beginning2,1)%></font>
</td>
</tr> 


<%
' Revised End
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_End%>&nbsp;(<%=myMessage_Correction%>)</b></font>
</td>
<td align="left">
<font size="2" face="Arial, Helvetica, sans-serif"><%=myDate_Display(myPhase_Date_End2,1)%></font>
</td>
</tr>
 
<%
' Leader
%>


<tr>
<td align="right"  bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Leader%></b></font>
</td>
<td align="Left">
<font size="2" face="Arial, Helvetica, sans-serif"><% = myLeader_Pseudo %></font>
</td>
</tr> 

<%
' Date Author
%>


<tr> 
<td align="Center" valign="top" bgcolor="<%=myApplicationColor%>" colspan="2">
<font face="Arial, Helvetica, sans-serif" size="1" color="<%=myApplicationTextColor%>"> 
<%=myDate_Display(myPhase_Date_Update,2)%> -- <%=myPhase_Author_Update%></font>
</td>
</tr>


</table>

<%
' NAVIGATION
%> 

<table border="0"  width="100% cellpadding="0" cellspacing="0"> 

<tr>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;<a href="__Projects_List.asp?List=<%=myList%>&Numpage=<%=myNumPage%>&Search=<%=mySearch%>"><%=myMessage_Project%>s</a>
,&nbsp;<a href="__Phases_List.asp?Project_ID=<%=myProject_ID%>"><%=myMessage_Phase%>s</a></font>
</td>
</tr> 

<tr>
<td align="left">
<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;<a href="__Phase_Modification.asp?Action=New&Project_ID=<%=myProject_ID%>"><%=myMessage_Add%>&nbsp;<%=myMessage_Phase%></a></font>
</td>
</tr> 

<%
' Can be Modify or Delete by Project Author, Project Leader, Phase Author, Phase Leader or Administrator
%> 

<% if myProject_Member_ID=myUser_ID or myProject_Leader_ID=myUser_ID or myPhase_Member_ID=myUser_ID or myPhase_Leader_ID=myUser_ID or myUser_type_ID=1 then %> 
<tr>
<td align="left">
<FONT SIZE="2" FACE="Arial, Helvetica, sans-serif">&nbsp;<A HREF="__Phase_Modification.asp?Phase_ID=<%=myPhase_ID%>&Project_ID=<%=myProject_ID%>"><%=myMessage_Modify%></A> , <A HREF="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Phase_Modification.asp?action=Delete&amp;Phase_ID=<%=myPhase_ID%>&Project_ID=<%=myProject_ID%>';"><%=myMessage_Delete%></A></FONT>
</td>
</tr>
<%End If%>


</table>


</TD>
</TR>

</TABLE>

<%
' DOWN
%>



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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</font></A> & contributors</FONT>
</TD>
</TR>
</TABLE>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</BODY>
</HTML>
