<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - O~ver~Apps - http://www.overapps.com
'
' This program "__Project_modification.asp" is free software; 
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
' Doesn't Work With PWS ???
%>

<%
' ------------------------------------------------------------
' 
' Name		 	: __Project_modification.asp
' Path		    : /
' Version 		: 1.15.0
' Description 	: Add, Modify, Delete a Project
' By			: Pierre Rouarch	
' Company		: OverApps
' Date			:November,21, 2001
'
' Contributions : Christophe Humbert, Jean-Luc Lesueur, Dania Tcherkezoff
'
' Modify by		:	
' Company		:
' Date			:
' ------------------------------------------------------------

Dim myPage
myPage="__Project_modification.asp"

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

' LOCAL VARIABLES DEFINITIONS

Dim myProject_Site_ID, myProject_Member_ID, myProject_Parent_ID, myProject_Parent_ID_Selected, myProject_ID, myProject_Name,   myProject_Presentation, myProject_Theme_ID, myProject_Theme_ID_Selected,  myProject_Type_ID, myProject_Type_ID_Selected, myProject_Theme_Name,  myProject_Type_Name, myProject_Public, myProject_Top,  myProject_Date_Beginning, myProject_Date_End, myDBeginning, myMBeginning, myYBeginning, myDEnd, myMEnd, myYEnd, myProject_Date_Beginning2, myProject_Date_End2, myDBeginning2, myMBeginning2, myYBeginning2, myDEnd2, myMEnd2, myYEnd2, myProject_Status_ID,myProject_Status_ID_Selected, myProject_Leader_ID, myProject_Leader_ID_Selected , myProject_Priority_ID, myProject_Priority_ID_Selected, myProject_Progress, myProject_Personnal, myProject_Date_Update , myProject_Author_Update 

Dim myStrProject_Date_Beginning, myStrProject_Date_End, myStrProject_Date_Beginning2, myStrProject_Date_End2



Dim myAction, myList, myTitle, myNumPage, mySearch



	
Dim mySQL_Select_Tb_Projects, mySQL_Select_Tb_Projects_Status, mySQL_Select_Tb_Projects_Priorities, mySQL_Insert_Tb_Projects, mySQL_Update_Tb_Projects, mySQL_Delete_Tb_Projects, mySet_Tb_Projects, mySet_Tb_Projects_Status, mySet_Tb_Projects_Priorities, mySQL_Select_Tb_Projects_themes, mySet_Tb_Projects_themes, mySQL_Select_Tb_Projects_Types, mySet_Tb_Projects_Types, mySQL_Insert_tb_Projects_Sites, mySQL_Insert_tb_Projects_Members



''''''''''''''''''''''''''''''''''''''''' 
' Get Parameters						'
'''''''''''''''''''''''''''''''''''''''''

myList = ""
myList = request("List")
myAction = request("Action")
if len(myAction)=0 then 
	myAction="Update"
end if

myProject_ID = request("Project_ID")
if len(myProject_ID)=0 then 
	myAction="New"
end if

if myAction="New" then
	myProject_Date_Beginning=myDate_Now()
	myDBeginning	 = Day(myProject_Date_Beginning)
	myMBeginning	 = Month(myProject_Date_Beginning)
	myYBeginning	 = Year(myProject_Date_Beginning)
	myStrProject_Date_Beginning=Year(myProject_Date_Beginning)&"/"&Month(myProject_Date_Beginning)&"/"&Day(myProject_Date_Beginning)
	myProject_Date_End    = now()
	myDEnd	 = Day(myProject_Date_End)
	myMEnd	 = Month(myProject_Date_End)
	myYEnd	 = Year(myProject_Date_End)
	myStrProject_Date_End=Year(myProject_Date_End)&"/"&Month(myProject_Date_End)&"/"&Day(myProject_Date_End)
	myProject_Date_Beginning2=Now()
	myDBeginning2	 = Day(myProject_Date_Beginning2)
	myMBeginning2	 = Month(myProject_Date_Beginning2)
	myYBeginning2	 = Year(myProject_Date_Beginning2)
	myStrProject_Date_Beginning2=Year(myProject_Date_Beginning2)&"/"&Month(myProject_Date_Beginning2)&"/"&Day(myProject_Date_Beginning2)
	myProject_Date_End2    = now()
	myDEnd2	 = Day(myProject_Date_End2)
	myMEnd2	 = Month(myProject_Date_End2)
	myYEnd2	 = Year(myProject_Date_End2)
	myStrProject_Date_End2=Year(myProject_Date_End2)&"/"&Month(myProject_Date_End2)&"/"&Day(myProject_Date_End2)
	myProject_Leader_ID_Selected=myUser_ID

end if 



myNumPage=Request("NumPage")
if Len(myNumPage)=0 then 
		myNumPage=1
end if
mySearch=Replace(Request("search"),"'","''")
		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Delete
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction = "Delete" then

	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

	' SEE RIGTHS 
	mySQL_Select_tb_Projects="SELECT * FROM tb_Projects WHERE Project_ID = "& myProject_ID&" and Site_ID="&mySite_ID	
	myConnection.Execute(mySQL_Select_Tb_Projects)
	set mySet_tb_Projects = myConnection.Execute(mySQL_Select_tb_Projects)
	' if eof go back
	if mySet_tb_Projects.eof then
		' Close Recordset 
		mySet_tb_Projects.close
		Set mySet_tb_Projects=Nothing
		
		' Close Connection
		myConnection.close
		set myConnection = nothing
		Response.Redirect("__Projects_List.asp?List="&myList&"&Numpage="&myNumPage&"&Search="&mySearch&"")
	end if

	myProject_Member_ID=mySet_tb_Projects("Member_ID")
	myProject_Leader_ID=mySet_tb_Projects("Project_Leader_ID")
	
	if myProject_Member_ID=myUser_ID or myProject_Leader_ID=myUser_ID or myUser_type_ID=1 then
		mySQL_Delete_Tb_Projects = "DELETE * FROM tb_Projects WHERE Project_ID = "& myProject_ID
	end if
	
	myConnection.Execute(mySQL_Delete_Tb_Projects)
	'Close Connection 
	myConnection.Close
	set myConnection = Nothing
	' and go Back
	Response.Redirect("__Projects_List.asp?List="&myList&"&Numpage="&myNumPage&"&Search="&mySearch&"")
end if


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


if Request.form("Validation")=myMessage_Go then

	' Get Entries
	' Not Used in this version 
	myProject_Parent_ID = Request.Form("Project_Parent_ID")
	myProject_ID  = Request.Form("Project_ID")
	myProject_Name = Replace(Request.Form("Project_Name"),"'"," ")
	myProject_Presentation=Replace(Request.Form("Project_Presentation"),"'"," ")

	myDBeginning=Request.Form("JBeginning")
	myMBeginning=Request.Form("MBeginning")
	myYBeginning=Request.Form("YBeginning")
	if (myDBeginning<> "" and myMBeginning <> "" and myYBeginning <> "") then
		myProject_Date_Beginning=myDate_Construct(myYBeginning,myMBeginning,myDBeginning,0,0,0)
	else 
		myProject_Date_Beginning=Null
	end if

	myDEnd=Request.Form("DEnd")
	myMEnd=Request.Form("MEnd")
	myYEnd=Request.Form("YEnd")
	if (myDEnd<> "" and myMEnd <> "" and myYEnd <> "") then
		myProject_Date_End =myDate_Construct(myYEnd,myMEnd,myDEnd,0,0,0)
	else 
		myProject_Date_End=Null
	end if

	myDBeginning2=Request.Form("JBeginning2")
	myMBeginning2=Request.Form("MBeginning2")
	myYBeginning2=Request.Form("YBeginning2")
	if (myDBeginning2<> "" and myMBeginning2 <> "" and myYBeginning2 <> "") then
		myProject_Date_Beginning2=myDate_Construct(myYBeginning2,myMBeginning2,myDBeginning2,0,0,0)
	else 
		myProject_Date_Beginning2=Null
	end if
	myDEnd2=Request.Form("DEnd2")
	myMEnd2=Request.Form("MEnd2")
	myYEnd2=Request.Form("YEnd2")
	if (myDEnd2<> "" and myMEnd2 <> "" and myYEnd2 <> "") then
		myProject_Date_End2=myDate_Construct(myYEnd2,myMEnd2,myDEnd2,0,0,0)
	else 
		myProject_Date_End2=Null
	end if


	myProject_Status_ID=Request.Form("Project_Status_ID")
	myProject_Leader_ID=Request.Form("Project_Leader_ID")
	myProject_Priority_ID=Request.Form("Project_Priority_ID")
	myProject_Progress=Request.Form("Project_Progress")
	if len(myProject_Progress)=0 then
		myProject_Progress=0
	end if
	myProject_Personnal=Request.Form("Project_Personnal")
	if len(myProject_Personnal)>0 then 
		myProject_Personnal=True
	else
		myProject_Personnal=False
	end if


	' Test Entries
	Call myFormSetEntriesInString

	myFormCheckEntry null, "Project_Name",true,null,null,0,100
	myFormCheckEntry null, "Project_Leader_ID",true,1,200000,0,6


	if not myform_entry_error  then

		myStrProject_Date_Beginning=myProject_Date_Beginning
	    myStrProject_Date_End=myProject_Date_End
		myStrProject_Date_Beginning2=myProject_Date_Beginning2
	    myStrProject_Date_End2=myProject_Date_End2

		
		
		' DB Connection 
		set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String

	
		myProject_Author_Update    = myUser_Pseudo
		myProject_Date_Update      = myDate_Now()


		if myAction = "New" then

			' Insert
			mySQL_Select_tb_Projects = "SELECT * FROM tb_Projects"
			Set mySet_tb_Projects = server.createobject("adodb.recordset")
			mySet_tb_Projects.open mySQL_Select_tb_Projects, myConnection, 3, 3
			mySet_tb_Projects.AddNew

			mySet_tb_Projects.fields("Site_ID")=mySite_ID
			mySet_tb_Projects.fields("Member_ID")=myUser_ID
			mySet_tb_Projects.fields("Project_Parent_ID")=myProject_Parent_ID
			mySet_tb_Projects.fields("Project_Name")=myProject_Name
			mySet_tb_Projects.fields("Project_Presentation")=myProject_Presentation
			mySet_tb_Projects.fields("Project_Date_Beginning")=myStrProject_Date_Beginning
			mySet_tb_Projects.fields("Project_Date_End")=myStrProject_Date_End	
			' Not Used
			mySet_tb_Projects.fields("Project_Date_Beginning2")=myStrProject_Date_Beginning2
			mySet_tb_Projects.fields("Project_Date_End2")=myStrProject_Date_End2
			mySet_tb_Projects.fields("Project_Status_ID")=myProject_Status_ID
			mySet_tb_Projects.fields("Project_Leader_ID")=myProject_Leader_ID
			mySet_tb_Projects.fields("Project_Priority_ID")=myProject_Priority_ID
			mySet_tb_Projects.fields("Project_Progress")=myProject_Progress
			mySet_tb_Projects.fields("Project_Personnal")=myProject_Personnal
			' /Not Used
			mySet_tb_Projects.fields("Project_Date_Update")=myProject_Date_Update
			mySet_tb_Projects.fields("Project_Author_Update")=myProject_Author_Update
	
			mySet_tb_Projects.Update
			' Close Recordset 
			mySet_tb_Projects.close
			Set mySet_tb_Projects = Nothing


	elseif myAction = "Update" then
	' Update in Database

			mySQL_Select_tb_Projects = "SELECT * FROM tb_Projects WHERE Project_ID =" & myProject_ID
			Set mySet_tb_Projects = server.createobject("adodb.recordset")
			mySet_tb_Projects.open mySQL_Select_tb_Projects, myConnection, 3, 3
	
			mySet_tb_Projects.fields("Project_Parent_ID")=myProject_Parent_ID
			mySet_tb_Projects.fields("Project_Name")=myProject_Name
			mySet_tb_Projects.fields("Project_Presentation")=myProject_Presentation
			mySet_tb_Projects.fields("Project_Date_Beginning")=myStrProject_Date_Beginning
			mySet_tb_Projects.fields("Project_Date_End")=myStrProject_Date_End
			' Not Used
			mySet_tb_Projects.fields("Project_Date_Beginning2")=myStrProject_Date_Beginning2
			mySet_tb_Projects.fields("Project_Date_End2")=myStrProject_Date_End2
			mySet_tb_Projects.fields("Project_Status_ID")=myProject_Status_ID
			mySet_tb_Projects.fields("Project_Leader_ID")=myProject_Leader_ID
			mySet_tb_Projects.fields("Project_Priority_ID")=myProject_Priority_ID
			mySet_tb_Projects.fields("Project_Progress")=myProject_Progress
			mySet_tb_Projects.fields("Project_Personnal")=myProject_Personnal
			'/Not Used
			mySet_tb_Projects.fields("Project_Date_Update")=myProject_Date_Update
			mySet_tb_Projects.fields("Project_Author_Update")=myProject_Author_Update
			mySet_tb_Projects.Update
			' Close Recordset 
			mySet_tb_Projects.close
			Set mySet_tb_Projects = Nothing

		end if


		' Close Connection
		myConnection.close
		set myConnection = nothing	
	

	Response.Redirect("__Projects_List.asp?List="&myList&"")
 	


end if 

end if 


%>
<html>

<head>
<title><%=mySite_Name%> - Project - Add/Modify/Delete</title>

</head>
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

<TD WIDTH="<%=myLeft_Width%>" BGColor="<%=myBorderColor%>">
<!-- #include file="_borders/Left.asp" -->
</td>
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form															'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


if myAction = "Update" then


	mySQL_Select_Tb_Projects = "SELECT * FROM tb_Projects WHERE Project_ID = " &myProject_ID
	set mySet_Tb_Projects = myConnection.Execute(mySQL_Select_Tb_Projects)

	' If Nothing Go Back
	if mySet_Tb_Projects.eof then
		Response.Redirect("__Projects_List.asp?List="&myList&"")
	end if

	' preparation des données
	myProject_Site_ID = mySet_Tb_Projects("Site_ID")
	myProject_Member_ID = mySet_Tb_Projects("Member_ID")
	' Not Used
	myProject_Parent_ID_Selected = mySet_Tb_Projects("Project_Parent_ID")
	if len(myProject_Parent_ID)=0 then 
		myProject_Parent_ID=0
	end if
	' / Not Used
	myProject_ID = mySet_Tb_Projects("Project_ID")
	myProject_Name  = mySet_Tb_Projects("Project_Name")
	myProject_Presentation  = mySet_Tb_Projects("Project_Presentation")
	' Not Used
	myProject_Date_Beginning=mySet_Tb_Projects("Project_Date_Beginning")
	myDBeginning	 = Day(myProject_Date_Beginning)
	myMBeginning	 = Month(myProject_Date_Beginning)
	myYBeginning	 = Year(myProject_Date_Beginning)
	myProject_Date_End=mySet_Tb_Projects("Project_Date_End")
	myDEnd	 = Day(myProject_Date_End)
	myMEnd	 = Month(myProject_Date_End)
	myYEnd	 = Year(myProject_Date_End)
	myProject_Date_Beginning2=mySet_Tb_Projects("Project_Date_Beginning2")
	myDBeginning2	 = Day(myProject_Date_Beginning2)
	myMBeginning2	 = Month(myProject_Date_Beginning2)
	myYBeginning2	 = Year(myProject_Date_Beginning2)
	myProject_Date_End2=mySet_Tb_Projects("Project_Date_End2")
	myDEnd2	 = Day(myProject_Date_End2)
	myMEnd2	 = Month(myProject_Date_End2)
	myYEnd2	 = Year(myProject_Date_End2)
	myProject_Status_ID_Selected = mySet_Tb_Projects("Project_Status_ID")
	if len(myProject_Status_ID_Selected)=0 then
		myProject_Status_ID_Selected=0
	end if 
	myProject_Priority_ID_Selected = mySet_Tb_Projects("Project_Priority_ID")
	if len(myProject_Priority_ID_Selected)=0 then 
		myProject_Priority_I_SelectedD = 0
	end if
	myProject_Progress = mySet_Tb_Projects("Project_Progress")
	if len(myProject_Progress)=0 then 
		 myProject_Progress = 0
	end if
	myProject_Personnal = mySet_Tb_Projects("Project_Personnal")
	' / Not USed


	myProject_Leader_ID_Selected = mySet_Tb_Projects("Project_Leader_ID")
	if len(myProject_Leader_ID_Selected)=0 or myProject_Leader_ID_Selected=0 then 
		myProject_Leader_ID_Selected = myUser_ID
	end if



myProject_Author_Update = mySet_Tb_Projects("Project_Author_Update")				
	myProject_Date_Update = mySet_Tb_Projects("Project_Date_Update")
	end if

' Close Connection
myConnection.close
set myConnection = nothing

%> 



<td valign="top">
<form method="POST" action="__Project_Modification.asp" name="myForm" > 
<table border="0" cellpadding="5" cellspacing="1" width="<%=myApplication_Width%>">


<%
' Title and hidden Fields
%>


<tr>
<td align="center" colspan="2" bgcolor="<%=myApplicationColor%>">
<font face="Arial, Helvetica, sans-serif" size="4"><b><%=myApplication_Title%></b></font>

<INPUT TYPE="hidden" NAME="Project_ID" VALUE="<%=myProject_ID%>">
<INPUT TYPE="hidden" NAME="Action" VALUE="<%=myAction%>"> 
<INPUT TYPE="hidden" NAME="JBeginning" VALUE="<%=myDBeginning%>">
<INPUT TYPE="hidden" NAME="MBeginning" VALUE="<%=myMBeginning%>"> 
<INPUT TYPE="hidden" NAME="YBeginning" VALUE="<%=myYBeginning%>">
<INPUT TYPE="hidden" NAME="DEnd" VALUE="<%=myDEnd%>"> 
<INPUT TYPE="hidden" NAME="MEnd" VALUE="<%=myMEnd%>">
<INPUT TYPE="hidden" NAME="YEnd" VALUE="<%=myYEnd%>"> 

</td>
</tr>

<%
' Project Name
%>
 
<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Name%>*<BR> 
<%=myFormGetErrMsg("Project_Name")%></B></font>
</td>
<td align="left" valign="top"> 
<input type="text" size="30" name="Project_Name" value="<%=myProject_Name%>"> 
</td>
</tr>

<%
' Presentation
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Presentation%></b></font>
</td>
<td align="left" valign="top">
<TEXTAREA COLS="65" NAME="Project_Presentation" ROWS="4" WRAP="PHYSICAL"><%=myProject_Presentation%></TEXTAREA>
</td>
</tr>


<%
' Leader
%>

<tr>
<td align="right"  bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" color="<%=myBorderTextColor%>"><b><%=myMessage_Leader%>*<BR> <%=myFormGetErrMsg("Project_Leader_ID")%></b></font>
</td>
            <td align="Left"> <b>
              <select name="Project_Leader_ID" size="1" tabindex="1">
                <%
'if myProject_Leader_ID_Selected=0 then
'	Response.Write "<option value=""0"">"&myMessage_Select&"</option>"
'else
'	Response.Write "<option selected value=""0"">"&myMessage_Select&"</option>"
'end if		

	' DB connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String
mySQL_Select_Tb_Sites_Members = "SELECT * FROM tb_Sites_Members Where Site_ID="&mySite_ID&" order by Member_Pseudo" 
set mySet_Tb_Sites_Members = myConnection.Execute(mySQL_Select_Tb_Sites_Members)

mySet_tb_Sites_Members.movefirst
do while not mySet_tb_Sites_Members.eof
	if myProject_Leader_ID_Selected = mySet_tb_Sites_Members("Member_ID")  then
		 Response.Write "<option selected value=" & mySet_tb_Sites_Members("Member_ID") & ">" & mySet_tb_Sites_Members("Member_Pseudo") & "</option>"
	else
		Response.Write "<option value=" & mySet_tb_Sites_Members("Member_ID") & ">" & mySet_tb_Sites_Members("Member_Pseudo") & "</option>"
	end if
mySet_tb_Sites_Members.movenext
loop


	' Close Recordset 
	mySet_tb_Sites_Members.close
	Set mySet_tb_Sites_Members=Nothing

	' Close Connection
	myConnection.close
	set myConnection = nothing
%> 
              </select>
              </b> </td>
</tr>


<%
' Validation
%>

<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">&nbsp;
</td>
<td valign="top" colspan="3" nowrap align="left"> 
<input type="submit" value="<%=myMessage_Go%>" name="Validation">
</td>
</tr>

<%
' Date Author
%>

<tr> 
<td align="CENTER" valign="top" colspan="4" bgcolor="<%=myApplicationColor%>">
<font face="Arial, Helvetica, sans-serif" size="1" Color="<%=myApplicationTextColor%>"> 
<%=myDate_Display(myProject_Date_Update,2) %> -- <%=myProject_Author_Update %></font>
</td>
</tr>

</table>

</form>


<%
' NAVIGATION
%>

<table border="0" width="90%" cellpadding="3" cellspacing="0"> <tr> <td><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="__Projects_List.asp?List=<%=myList%>&Numpage=<%=myNumPage%>&Search=<%=mySearch%>"><%=myMessage_Project%>s</a> 
<% if myAction="Update" then %> ,&nbsp;<a href="__Phases_List.asp?Project_ID=<%=myProject_ID%>"><%=myMessage_Phase%>s</a></font> 
<% end if%> </td></tr> <% if myAction="Update" then %> <tr> <td><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="__Project_Modification.asp?Action=New"><%=myMessage_Add%>&nbsp;<%=myMessage_Project%></a></font> 
</td></tr> <%
' Can be Modify or Delete by Author, Leader or Administrator
%> <% if myProject_Member_ID=myUser_ID or myProject_Leader_ID_Selected=myUser_ID or myUser_type_ID=1 then %> 
<tr> <td><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Project_Modification.asp?action=Delete&Project_ID=<%=myProject_ID%>'"><%=myMessage_Delete%></a></font> 
</td></tr>
<% End IF %>
<%End If%>
</table>
</td>
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
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>"> 
Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</FONT></A> & contributors</FONT>
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


<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>