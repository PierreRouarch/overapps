<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - O|verAp|ps - http://www.overapps.com
'
' This program "__Phase_Modification.asp" is free software; you can redistribute it and/or modify
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
' Does n't Work with PWS ????
%>

<%
' ------------------------------------------------------------
' Name			: __Phase_Modification.asp
' Path	 		: /
' Version 		: 1.15.0
' Description 	: Add/Modify or Delete a Phase
' 
' By 			: Pierre Rouarch	
' Company		: OverApps
' Date			: November, 21, 2001
'
' Contributions : Dania Tcherkezoff
'
' Update by		:
' Company		:
' Date			:
' Modifications :
' ------------------------------------------------------------

Dim myPage
myPage = "__Phase_Modification.asp"

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
' Local Variables


Dim myProject_ID, myProject_Member_ID, myProject_Date_Beginning, myStrProject_Date_Beginning, myProject_Date_End, myStrProject_Date_End, myProject_Date_Beginning2,  myProject_Date_End2, myProject_Leader_ID, myProject_Leader_ID_Selected, myProject_Date_Update, myProject_Author_Update 


Dim myPhase_Site_ID, myPhase_Member_ID, myPhase_Project_ID, myPhase_Project_Name, myPhase_Parent_ID,  myPhase_Parent_ID_Selected, myPhase_Parent_Name, myPhase_ID, myPhase_Path,  myPhase_Name, myPhase_Presentation, myPhase_Date_Beginning, myStrPhase_Date_Beginning, myPhase_Date_End, myStrPhase_Date_End, myPhase_Date_Beginning2, myStrPhase_Date_Beginning2, myPhase_Date_End2, myStrPhase_Date_End2, myPhase_Status_ID, myPhase_Status_ID_Selected, myPhase_Leader_ID, myPhase_Leader_ID_Selected, myPhase_Priority_ID, myPhase_Priority_ID_Selected, myPhase_Progress, myPhase_Personnal, myPhase_Date_Update, myPhase_Author_Update


Dim  myDay_Beginning, myMonth_Beginning, myYear_Beginning, myDay_End, myMonth_End, myYear_End, myDay_Beginning2, myMonth_Beginning2, myYear_Beginning2, myDay_End2, myMonth_End2, myYear_End2



Dim  mySQL_Select_tb_Phases, mySQL_Select_tb_Phases2, mySet_tb_Phases, mySet_tb_Phases2, mySQL_Delete_tb_Phases, mySQL_Select_tb_Projects, mySet_tb_Projects,  mySQL_Select_tb_Phases_Status, mySet_tb_Phases_Status,  mySQL_Select_tb_Phases_Priorities, mySet_tb_Phases_Priorities



Dim myTitle, mySubmit

Dim myAction, myList, myNumPage, mySearch



' Get Parameters

myProject_ID=request("Project_ID")
if  Len(myProject_ID)=0 then
		Response.Redirect("__Projects_List.asp")
end if

myPage=myPage&"?Project_ID="&myProject_ID
	
myList = ""
myList = request("List")

myNumPage = request("NumPage")

myAction = request("Action")
if len(myAction)=0 then 
	myAction="Update"
end if

myPhase_ID = request("Phase_ID")
if len(myPhase_ID)=0 then 
	myAction="New"
end if

if myAction="New" then
	myPhase_Date_Beginning=Now()
	myDay_Beginning	 = Day(myPhase_Date_Beginning)
	myMonth_Beginning	 = Month(myPhase_Date_Beginning)
	myYear_Beginning	 = Year(myPhase_Date_Beginning)
	myStrPhase_Date_Beginning=Year(myPhase_Date_Beginning)&"/"&Month(myPhase_Date_Beginning)&"/"&Day(myPhase_Date_Beginning)
	myPhase_Date_End    = now()
	myDay_End	 = Day(myPhase_Date_End)
	myMonth_End	 = Month(myPhase_Date_End)
	myYear_End	 = Year(myPhase_Date_End)
	myStrPhase_Date_End=Year(myPhase_Date_End)&"/"&Month(myPhase_Date_End)&"/"&Day(myPhase_Date_End)
	myPhase_Date_Beginning2=Null
	myDay_Beginning2	 = null
	myMonth_Beginning2	 = null
	myYear_Beginning2	 = null
	myPhase_Date_End2    = null
	myDay_End2	 = null
	myMonth_End2	 = null
	myYear_End2	 = null
	myPhase_Leader_ID=myUSer_ID
	myPhase_Leader_ID_Selected=myPhase_Leader_ID
end if 




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sub Programm - UPDATE PROJECTS IN DATABASE 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub myProc_Update_Project ()

myProject_Date_Update=myDate_Now()
myProject_Author_Update=myUser_Pseudo

' Update Beginning Date 

' Get First Beginning Date in Phases

mySQL_Select_tb_Phases = "SELECT * FROM tb_phases WHERE Project_ID="&myProject_ID&" AND Phase_Date_Beginning <> '' ORDER BY Phase_Date_Beginning "
set mySet_tb_Phases = myConnection.Execute(mySQL_Select_tb_Phases) 

if not mySet_tb_Phases.eof then

	myProject_Date_Beginning = mySet_tb_Phases("Phase_Date_Beginning")

	mySQL_Select_tb_Phases2 = "SELECT * FROM tb_phases WHERE Project_ID="&myProject_ID&" AND Phase_Date_Beginning2 <> '' ORDER BY Phase_Date_Beginning2 "
	set mySet_tb_Phases2 = myConnection.Execute(mySQL_Select_tb_Phases2)		

	if not mySet_tb_Phases2.eof then	

		if isDate(mySet_tb_Phases2("Phase_Date_Beginning2")) AND myProject_Date_Beginning>mySet_tb_Phases2("Phase_Date_Beginning2") then 
				myProject_Date_Beginning=mySet_tb_Phases2("Phase_Date_Beginning2")
		end if

		mySet_tb_Phases2.close
		Set mySet_tb_Phases2 = Nothing
	end if

	mySet_tb_Phases.close
	Set mySet_tb_Phases = Nothing

	' Update Beginning Date in tb_Projects
	mySQL_Select_tb_Projects = "SELECT * FROM tb_Projects WHERE Project_ID =" & myProject_ID
	Set mySet_tb_Projects = server.createobject("adodb.recordset")
	mySet_tb_Projects.open mySQL_Select_tb_Projects, myConnection, 3, 3

	mySet_tb_Projects.fields("Project_Date_Beginning")=myProject_Date_Beginning

	mySet_tb_Projects.Update

	' Close Recordset 
	mySet_tb_Projects.close
	Set mySet_tb_Projects = Nothing

end if



' Update End Project 
		
mySQL_Select_Tb_Phases = "SELECT * FROM tb_phases WHERE Project_ID="&myProject_ID&" AND Phase_Date_End <> '' ORDER BY Phase_Date_End DESC "
set mySet_tb_Phases = myConnection.Execute(mySQL_Select_Tb_Phases) 
if not mySet_tb_Phases.eof then
	myProject_Date_End = mySet_tb_Phases("Phase_Date_End")
	mySQL_Select_Tb_Phases2 = "SELECT * FROM tb_phases WHERE Project_ID="&myProject_ID&" AND Phase_Date_End2 <> '' ORDER BY Phase_Date_End DESC "
	set mySet_tb_Phases2 = myConnection.Execute(mySQL_Select_Tb_Phases2) 
	if not mySet_Tb_Phases2.eof then
		if isDate(mySet_tb_Phases2("Phase_Date_End2")) and  myProject_Date_End < mySet_tb_Phases2("Phase_Date_End2") then
					myProject_Date_End = mySet_tb_Phases2("Phase_Date_End2")
		end if
		mySet_tb_Phases2.close
		Set mySet_tb_Phases2 = Nothing
	end if

	mySet_tb_Phases.close
	Set mySet_tb_Phases = Nothing

	' Update End Date in tb_Projects (execute does n't work Well  with PWS ???)
	mySQL_Select_tb_Projects = "SELECT * FROM tb_Projects WHERE Project_ID =" & myProject_ID
	Set mySet_tb_Projects = server.createobject("adodb.recordset")
	mySet_tb_Projects.open mySQL_Select_tb_Projects, myConnection, 3, 3

	mySet_tb_Projects.fields("Project_Date_End")=myProject_Date_End

	mySet_tb_Projects.Update

	' Close Recordset 
	mySet_tb_Projects.close
	Set mySet_tb_Projects = Nothing

end if



' Update Author and Date Update
mySQL_Select_tb_Projects = "SELECT * FROM tb_Projects WHERE Project_ID =" & myProject_ID
Set mySet_tb_Projects = server.createobject("adodb.recordset")
mySet_tb_Projects.open mySQL_Select_tb_Projects, myConnection, 3, 3

mySet_tb_Projects.fields("Project_Date_Update")=myProject_Date_Update
mySet_tb_Projects.fields("Project_Author_Update")=myProject_Author_Update
mySet_tb_Projects.Update

' Close Recordset 
mySet_tb_Projects.close
Set mySet_tb_Projects = Nothing

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' End Sub Programm UPDATE PROJECTS IN DATABASE 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Delete Phase
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


if myAction = "Delete" then
	' DB Connection
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String
	mySQL_Select_tb_Phases = "Select * FROM tb_phases WHERE Phase_ID = " & myPhase_ID
	set mySet_tb_Phases = 	myConnection.Execute(mySQL_Select_Tb_Phases)
	myProject_ID= mySet_tb_Phases("Project_ID")
	mySQL_Delete_tb_Phases = "DELETE FROM tb_phases WHERE Phase_ID = " & myPhase_ID
	myConnection.Execute(mySQL_Delete_Tb_Phases)
	mySet_tb_Phases.close
	Set mySet_tb_Phases=nothing

	' Update Project
	Call myProc_Update_Project()
	myConnection.Close
	set myConnection = Nothing
	Response.Redirect("__Phases_List.asp?Project_ID="&myProject_ID&"")
end if			
		



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


if Request.form("Validation")=myMessage_Go then


	' GET ENTRIES 	
	myPhase_Name = Replace(Request.Form("Phase_Name"),"'"," ")
	myPhase_Presentation = Replace(Request.Form("Phase_Presentation"),"'"," ")

	myPhase_Parent_ID	= Request.Form("Phase_Parent_ID")
	myPhase_Parent_ID_Selected=myPhase_Parent_ID

	myPhase_Path=myPhase_Parent_ID&"/"&myPhase_ID 
	myDay_Beginning=Request.Form("Day_Beginning")
	myMonth_Beginning=Request.Form("Month_Beginning")
	myYear_Beginning=Request.Form("Year_Beginning")
	if (myDay_Beginning<> "" and myMonth_Beginning <> "" and myYear_Beginning <> "") then
		myPhase_Date_Beginning= myDate_Construct(myYear_Beginning,myMonth_Beginning,myDay_Beginning,0,0,0)
	else 
		myPhase_Date_Beginning=Null
	end if
	myDay_End=Request.Form("Day_End")
	myMonth_End=Request.Form("Month_End")
	myYear_End=Request.Form("Year_End")
	if (myDay_End<> "" and myMonth_End <> "" and myYear_End <> "") then
		myPhase_Date_End= myDate_Construct(myYear_End,myMonth_End,myDay_End,0,0,0)
	else 
		myPhase_Date_End=Null
	end if


	myDay_Beginning2=Request.Form("Day_Beginning2")
	myMonth_Beginning2=Request.Form("Month_Beginning2")
	myYear_Beginning2=Request.Form("Year_Beginning2")
	if (myDay_Beginning2<>"" and myMonth_Beginning2 <> "" and myYear_Beginning2 <> "") then
		myPhase_Date_Beginning2=myDate_Construct(myYear_Beginning2,myMonth_Beginning2,myDay_Beginning2,0,0,0)
	else 
		myPhase_Date_Beginning2=Null
	end if
	myDay_End2=Request.Form("Day_End2")
	myMonth_End2=Request.Form("Month_End2")
	myYear_End2=Request.Form("Year_End2")
	if (myDay_End2<> "" and myMonth_End2 <> "" and myYear_End2 <> "") then
		myPhase_Date_End2=myDAte_Construct(myYear_End2,myMonth_End2,myDay_End2,0,0,0)
	else 
		myPhase_Date_End2=Null
	end if
	

	myPhase_Leader_ID= Request.Form("Phase_Leader_ID")
	myPhase_Leader_ID_Selected=myPhase_Leader_ID


	' TEST ENTRIES
 
	Call myFormSetEntriesInString

	myFormCheckEntry null, "Phase_Name",true,null,null,0,100
	myFormCheckEntry myErr_Numerical, "Day_Beginning",true,1,31,0,2
	myFormCheckEntry null, "Month_Beginning",true,1,12,0,2
	myFormCheckEntry myErr_Numerical, "Year_Beginning",true,1999,2011,4,4
	myFormCheckEntry myErr_Numerical, "Day_End",true,1,31,0,2
	myFormCheckEntry null, "Month_End",true,1,12,0,2
	myFormCheckEntry myErr_Numerical, "Year_End",true,1999,2011,4,4

	myFormCheckEntry myErr_Numerical, "Day_Beginning2",false,1,31,0,2
	myFormCheckEntry null, "Month_Beginning2",false,0,12,0,2
	myFormCheckEntry myErr_Numerical, "Year_Beginning2",false,1999,2011,4,4
	myFormCheckEntry myErr_Numerical, "Day_End2",false,1,31,0,2
	myFormCheckEntry null, "Month_End2",false,0,12,0,2
	myFormCheckEntry myErr_Numerical, "Year_End2",false,1999,2011,4,4


	' Other Errors

	if (not isDate(myPhase_Date_Beginning)) then 
		myform_entry_error=true
	end if

	if (not isDate(myPhase_Date_End)) then 
		myform_entry_error=true
	end if

	if len(myPhase_Date_Beginning)>0 and len(myPhase_Date_End)>0 then
		if (isDate(myPhase_Date_Beginning) and isDate(myPhase_Date_End)) then 
	 		if cdate(myPhase_Date_End)<cDate(myPhase_Date_Beginning) then 
				myform_entry_error=true
			end if 
		end if
	end if

	if (not isDate(myPhase_Date_Beginning) and isDate(myPhase_Date_End)) then 
		myform_entry_error=true
	end if

	if (isDate(myPhase_Date_Beginning) and not isDate(myPhase_Date_End)) then 
		myform_entry_error=true
	end if


	if (len(myPhase_Date_Beginning2)>0 And not isDate(myPhase_Date_Beginning2)) then 
		myform_entry_error=true
	end if

	if (len(myPhase_Date_End2)>0 AND not isDate(myPhase_Date_End2)) then 
		myform_entry_error=true
	end if

	if len(myPhase_Date_Beginning2)>0 and len(myPhase_Date_End2)>0 then
		if (isDate(myPhase_Date_Beginning2) and isDate(myPhase_Date_End2)) then
 			if cDate(myPhase_Date_End2)<CDate(myPhase_Date_Beginning2) then 
				myform_entry_error=true
			end if 
		end if
	end if

	if (not isDate(myPhase_Date_Beginning2) and isDate(myPhase_Date_End2)) then 
		myform_entry_error=true
	end if

	if (isDate(myPhase_Date_Beginning2) and not isDate(myPhase_Date_End2)) then 
		myform_entry_error=true
	end if

	'''''''''''''''''''''''''''''''''''''
	' / TEST ENTRIES					'
	'''''''''''''''''''''''''''''''''''''


	if not myform_entry_error then
		
		myStrPhase_Date_Beginning=myPhase_Date_Beginning
	    myStrPhase_Date_End=myPhase_Date_End
	
		myStrPhase_Date_Beginning2=null
		if isDate(myPhase_Date_Beginning2) then 
			myStrPhase_Date_Beginning2=myPhase_Date_Beginning2
		end if
		myStrPhase_Date_End2=null
		if isDate(myPhase_Date_End2) then 	
	    	myStrPhase_Date_End2=myPhase_Date_End2
		end if
		myPhase_Author_Update = myUser_Pseudo
		myPhase_Date_Update = myDate_Now()

		set myConnection = Server.CreateObject("ADODB.Connection")
		myConnection.Open myConnection_String



		if myAction = "New" then

			' Insert
			mySQL_Select_tb_Phases = "SELECT * FROM tb_Phases"
			Set mySet_tb_Phases = server.createobject("adodb.recordset")
			mySet_tb_Phases.open mySQL_Select_tb_Phases, myConnection, 3, 3
			mySet_tb_Phases.AddNew

			mySet_tb_Phases.fields("Site_ID")=mySite_ID
			mySet_tb_Phases.fields("Member_ID")=myUser_ID
			mySet_tb_Phases.fields("Project_ID")=myProject_ID

			mySet_tb_Phases.fields("Phase_Parent_ID")=myPhase_Parent_ID
			mySet_tb_Phases.fields("Phase_Name")=myPhase_Name
			mySet_tb_Phases.fields("Phase_Presentation")=myPhase_Presentation
			mySet_tb_Phases.fields("Phase_Date_Beginning")=myStrPhase_Date_Beginning
			mySet_tb_Phases.fields("Phase_Date_End")=myStrPhase_Date_End	
			

			mySet_tb_Phases.fields("Phase_Date_Beginning2")=myStrPhase_Date_Beginning2
			mySet_tb_Phases.fields("Phase_Date_End2")=myStrPhase_Date_End2


			mySet_tb_Phases.fields("Phase_Leader_ID")=myPhase_Leader_ID


			mySet_tb_Phases.fields("Phase_Date_Update")=myPhase_Date_Update
			mySet_tb_Phases.fields("Phase_Author_Update")=myPhase_Author_Update
	
			mySet_tb_Phases.Update
			' Close Recordset 
			mySet_tb_Phases.close
			Set mySet_tb_Phases = Nothing

	
	elseif  myAction = "Update" then

			'UPDATE
			Response.write "UPDATE"
			mySQL_Select_tb_Phases = "SELECT * FROM tb_Phases Where Phase_ID="&myPhase_ID&" AND Project_ID="&myProject_ID

			Set mySet_tb_Phases = server.createobject("adodb.recordset")
			mySet_tb_Phases.open mySQL_Select_tb_Phases, myConnection, 3, 3

			mySet_tb_Phases.fields("Site_ID")=mySite_ID
			mySet_tb_Phases.fields("Member_ID")=myUser_ID
			mySet_tb_Phases.fields("Project_ID")=myProject_ID

			mySet_tb_Phases.fields("Phase_Parent_ID")=myPhase_Parent_ID
			mySet_tb_Phases.fields("Phase_Name")=myPhase_Name
			mySet_tb_Phases.fields("Phase_Presentation")=myPhase_Presentation
			mySet_tb_Phases.fields("Phase_Date_Beginning")=myStrPhase_Date_Beginning
			mySet_tb_Phases.fields("Phase_Date_End")=myStrPhase_Date_End	
			


			mySet_tb_Phases.fields("Phase_Date_Beginning2")=myStrPhase_Date_Beginning2
			mySet_tb_Phases.fields("Phase_Date_End2")=myStrPhase_Date_End2


			mySet_tb_Phases.fields("Phase_Leader_ID")=myPhase_Leader_ID



			mySet_tb_Phases.fields("Phase_Date_Update")=myPhase_Date_Update
			mySet_tb_Phases.fields("Phase_Author_Update")=myPhase_Author_Update
	
			mySet_tb_Phases.Update
			' Close Recordset 
			mySet_tb_Phases.close
			Set mySet_tb_Phases = Nothing


	
		end if

		' Update Data in tb_Projects
		Call myProc_Update_Project()

		myConnection.close
		set myConnection = nothing

	Response.Redirect("__Phases_List.asp?Project_ID="&myProject_ID&"")



end if ' No Entry Error

end if ' End Form Validation



%> 

<html> 

<head> 
<title><%=mySite_Name%> - Phase - Add/Modify/Delete </title> 
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

<%
' CENTER LEFT
%> 

<TD WIDTH="<%=myLeft_Width%>" BGCOLOR="<%=myBorderColor%>">
<!-- #include file="_borders/Left.asp" -->
</td>
<%
' CENTER APPLICATION
%> 

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	Populate Form																  '	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myAction="Update"  and not Request.form("Validation")=myMessage_Go then 

	' DB Connection
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String

	mySQL_Select_tb_Phases   = "SELECT * FROM tb_phases WHERE Phase_ID = " & myPhase_ID & " AND Project_ID = " & myProject_ID
	set mySet_tb_Phases = myConnection.execute(mySQL_Select_tb_Phases)

	' If Nothing Go Out
	if mySet_tb_Phases.eof then
		mySet_tb_Phases.close
		set mySet_tb_Phases=nothing
		myConnection.close
		set myConnection = nothing
		Response.Redirect("__Phases_List.asp?Project_ID="&myProject_ID&"")
	else
	' Get Information

		myPhase_Site_ID=mySet_tb_Phases("Site_ID")
		myPhase_Member_ID=mySet_tb_Phases("Member_ID")
		myPhase_Project_ID=mySet_tb_Phases("Project_ID")

		myPhase_Parent_ID= mySet_tb_Phases("Phase_Parent_ID")
		myPhase_Parent_ID_Selected=myPhase_Parent_ID

		myPhase_ID  = mySet_tb_Phases("Phase_ID")
		myPhase_Name = mySet_tb_Phases("Phase_Name")
		myPhase_Presentation = mySet_tb_Phases("Phase_Presentation")
		myPhase_Date_Beginning = mySet_tb_Phases("Phase_Date_Beginning")
		myDay_Beginning = Day(myPhase_Date_Beginning)
		myMonth_Beginning = Month(myPhase_Date_Beginning)
		myYear_Beginning = Year(myPhase_Date_Beginning)
		myPhase_Date_End = mySet_tb_Phases("Phase_Date_End")
		myDay_End=Day(myPhase_Date_End)
		myMonth_End=Month(myPhase_Date_End)
		myYear_End = Year(myPhase_Date_End)
		myPhase_Date_Beginning2 = mySet_tb_Phases("Phase_Date_Beginning2")
		myDay_Beginning2 = Day(myPhase_Date_Beginning2)
		myMonth_Beginning2 = Month(myPhase_Date_Beginning2)
		myYear_Beginning2 = Year(myPhase_Date_Beginning2)
		myPhase_Date_End2 = mySet_tb_Phases("Phase_Date_End2")
		myDay_End2	 = Day(myPhase_Date_End2)
		myMonth_End2	 = Month(myPhase_Date_End2)
		myYear_End2	 = Year(myPhase_Date_End2)

		myPhase_Leader_ID=mySet_tb_Phases("Phase_Leader_ID")
		myPhase_Leader_ID_Selected=myPhase_Leader_ID

		' Not Used
		myPhase_Status_ID = mySet_tb_Phases("Phase_Status_ID")
		myPhase_Status_ID_Selected=myPhase_Status_ID
		myPhase_Priority_ID=mySet_tb_Phases("Phase_Priority_ID")
		myPhase_Progress=mySet_tb_Phases("Phase_Progress")
		myPhase_Personnal=mySet_tb_Phases("Phase_Personnal")
		' Not Used

		myPhase_Date_Update= mySet_tb_Phases("Phase_Date_Update")
		myPhase_Author_Update= mySet_tb_Phases("Phase_Author_Update")
	
	end if
	mySet_tb_Phases.close
	set mySet_tb_Phases=Nothing
	myConnection.close
	set myConnection = nothing
end if


''''''''''''''''''''
' PROJECT NAME     '
''''''''''''''''''''
' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Projects = "Select * From tb_Projects Where Project_ID="&myProject_ID
	set mySet_tb_Projects = myConnection.execute(mySQL_Select_tb_Projects)
myPhase_Project_Name=mySet_tb_Projects("Project_Name")


mySet_tb_Projects.close
Set mySet_tb_Projects=nothing
myConnection.close
set myConnection = nothing

%> 

<td bgcolor="<%=myBGColor%>" align="left" valign="top">

<form method="POST" action="<%=myPage%>" name="myForm"> 
<input type="hidden" name="Action" value="<%=myAction%>">
<input type="hidden" name="Phase_ID" value="<%=myPhase_ID%>"> 
<INPUT TYPE="hidden" NAME="Project_ID" VALUE="<%=myProject_ID%>"> 


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
' Phase Name
%>

<tr>
<td  align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Name%>*<br><%=myFormGetErrMsg("Phase_Name")%></b></FONT>
</td>
<td align="left"> 
<input type="text" name="Phase_Name" size="40" value="<%=myPhase_Name%>">
</td>
</tr>


<%
' Phase Presentation
%>

 
<tr>
<td align="right" valign="top" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Presentation%></b></font>
</td>
<td align="left"> 
<textarea rows="4" name="Phase_Presentation" cols="56" WRAP="PHYSICAL"><%=myPhase_Presentation%></textarea> 
</td>
</tr>
 
<%
' Parent
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Parent%></b></font>
</td>

<td align="Left"> 
<select name="Phase_Parent_ID" size="1" tabindex="1"> 



<%
'''''''''''''''''''''''''''''''''''''
' PARENT'S PHASE					'	
'''''''''''''''''''''''''''''''''''''

' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

' If it's a phase (not a sub-phase) Write Select
if (len(myPhase_Parent_ID_Selected)=0 or myPhase_Parent_ID_Selected=0) then
	Response.Write "<option selected value=""0"">"&myMessage_Select&"</option>"
else
	Response.Write "<option value=""0"">"&myMessage_Select&"</option>"
end if		

' Phases List Level 1

mySQL_Select_tb_Phases2 = "SELECT * FROM tb_phases WHERE Phase_Parent_ID=0 AND Project_ID = " & myProject_ID
set mySet_tb_Phases2 = myConnection.Execute(mySQL_Select_tb_Phases2)

do while not mySet_tb_Phases2.eof

	' A Phase can't be its own parent 
	if myPhase_ID <> mySet_tb_Phases2("Phase_ID") then
		if CInt(myPhase_Parent_ID_Selected) = mySet_tb_Phases2("Phase_ID") then
			Response.Write "<option selected value=" & mySet_tb_Phases2("Phase_ID") & ">" & 			mySet_tb_Phases2("Phase_Name") & "</option>"
		else
			Response.Write "<option value=" & mySet_tb_Phases2("Phase_ID") & ">" & 			mySet_tb_Phases2("Phase_Name") & "</option>"
		end if
	end if
	mySet_tb_Phases2.movenext
loop

mySet_tb_Phases2.close
Set mySet_tb_Phases2=Nothing
myConnection.close
set myConnection = nothing
%>

</select>
</td>
</tr>

<%
' Date Beginning
%>

<tr>
<td align="right"  bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Beginning%>*</b></font> 
</td>
<td align="Left"> 
<P>
<SELECT NAME="Month_Beginning">
<OPTION VALUE="0" <%if (myMonth_Beginning=0 or Len(myMonth_Beginning)=0) then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMonth_Beginning=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if myMonth_Beginning=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if myMonth_Beginning=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if myMonth_Beginning=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if myMonth_Beginning=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if myMonth_Beginning=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if myMonth_Beginning=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if myMonth_Beginning=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if myMonth_Beginning=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if myMonth_Beginning=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if myMonth_Beginning=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if myMonth_Beginning=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>
&nbsp;
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> 
<%=myFormGetErrMsg("Month_Beginning")%></FONT></B>

<INPUT TYPE="text" SIZE="2" MAXLENGTH="2" NAME="Day_Beginning" VALUE="<%=myDay_Beginning%>"> 
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myFormGetErrMsg("Day_Beginning")%></FONT></B> 
, 

<INPUT TYPE="text" SIZE="4" MAXLENGTH="4" NAME="Year_Beginning" VALUE="<%=myYear_Beginning%>"> 
&nbsp;<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myFormGetErrMsg("Year_Beginning")%></FONT></B> 
</P>&nbsp; 

<%if not isDate(myPhase_Date_Beginning) then %> 
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myError_Message_Invalid_Date%></FONT></B> 
<%end if %>

</td>

<% 
' Date End
%>

</tr> <tr><td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_End%>*</b></font>
</td>

<td align="Left"> 
<P>
<SELECT NAME="Month_End"> 
<OPTION VALUE="0" <%if (myMonth_End=0 or Len(MyMonth_End)=0) then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMonth_End=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if myMonth_End=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if myMonth_End=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if myMonth_End=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if myMonth_End=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if myMonth_End=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if myMonth_End=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if myMonth_End=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if myMonth_End=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if myMonth_End=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if myMonth_End=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if myMonth_End=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>

&nbsp;
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> 
<%=myFormGetErrMsg("Month_End")%></FONT></B>

<INPUT TYPE="text" SIZE="2" MAXLENGTH="2" NAME="Day_End" VALUE="<%=myDay_End%>"> 
<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myFormGetErrMsg("Day_End")%></FONT></B>
 , 
<INPUT TYPE="text" SIZE="4" MAXLENGTH="4" NAME="Year_End" VALUE="<%=myYear_End%>"> 
&nbsp; <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> 
<%=myFormGetErrMsg("Year_End")%></FONT></B> </P>
&nbsp; 
<%if not isDate(myPhase_Date_End) then %> 
	<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myError_Message_Invalid_Date%></FONT></B> 
<%
end if
 %>

<% 
if len(myPhase_Date_End2)>0 and len(myPhase_Date_Beginning2)>0 then
	if (isDate(myPhase_Date_End) and isDate(myPhase_Date_Beginning)) then 
		if cdate(myPhase_Date_End)<cdate(myPhase_Date_Beginning) then
		%>
		<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myError_Message_Date_End_Before_Date_Beginning%></FONT></B> 
		<%
		end if
	end if 
end if 
%> 

</td>
</tr>

<%
' Beginning Revised
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Beginning%>&nbsp;(<%=myMessage_Correction%>)</b></font>
</td>
<td align="Left">
<P>
<SELECT NAME="Month_Beginning2"> 
<OPTION VALUE="0" <%if (myMonth_Beginning2=0 or len(myMonth_Beginning2)=0) then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMonth_Beginning2=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if myMonth_Beginning2=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if myMonth_Beginning2=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if myMonth_Beginning2=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if myMonth_Beginning2=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if myMonth_Beginning2=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if myMonth_Beginning2=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if myMonth_Beginning2=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if myMonth_Beginning2=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if myMonth_Beginning2=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if myMonth_Beginning2=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if myMonth_Beginning2=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>

&nbsp;
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<INPUT TYPE="text" SIZE="2" MAXLENGTH="2" NAME="Day_Beginning2" VALUE="<%=myDay_Beginning2%>"> 
, 
<INPUT TYPE="text" SIZE="4" MAXLENGTH="4" NAME="Year_Beginning2" VALUE="<%=myYear_Beginning2%>"> 
&nbsp;</FONT></P>&nbsp; 

<%
if (myDay_Beginning2<>"" or (myMonth_Beginning2<>"" and myMonth_Beginning2<>"0") or myYear_Beginning2<>"") AND  not isDate(myPhase_Date_Beginning2) then
	 %> 
	<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myError_Message_Invalid_Date%></FONT></FONT></B> 
	<%
end if
%>

<%
 if (isDate(myPhase_Date_End2) and Not isDate(myPhase_Date_Beginning2)) then
	%> 
	<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myError_Message_Enter_Date_Beginning%></FONT></B> 
	<%
end if
%>
</td>
</tr>


<%
' End Revised
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_End%>&nbsp;(<%=myMessage_Correction%>)</b></font>
</td>
<td align="Left"> 
<P>
<SELECT NAME="Month_End2">
<OPTION VALUE="0" <%if (myMonth_End2=0 or len(myMonth_End2)=0)then%>Selected<%end if%>><%=myMessage_Select%></OPTION> 
<OPTION VALUE="1" <%if myMonth_End2=1 then%>Selected<%end if%>><%=myMessage_January%></OPTION> 
<OPTION VALUE="2" <%if myMonth_End2=2 then%>Selected<%end if%>><%=myMessage_February%></OPTION> 
<OPTION VALUE="3" <%if myMonth_End2=3 then%>Selected<%end if%>><%=myMessage_March%></OPTION> 
<OPTION VALUE="4" <%if myMonth_End2=4 then%>Selected<%end if%>><%=myMessage_April%></OPTION> 
<OPTION VALUE="5" <%if myMonth_End2=5 then%>Selected<%end if%>><%=myMessage_May%></OPTION> 
<OPTION VALUE="6" <%if myMonth_End2=6 then%>Selected<%end if%>><%=myMessage_June%></OPTION> 
<OPTION VALUE="7" <%if myMonth_End2=7 then%>Selected<%end if%>><%=myMessage_July%></OPTION> 
<OPTION VALUE="8" <%if myMonth_End2=8 then%>Selected<%end if%>><%=myMessage_August%></OPTION> 
<OPTION VALUE="9" <%if myMonth_End2=9 then%>Selected<%end if%>><%=myMessage_September%></OPTION> 
<OPTION VALUE="10" <%if myMonth_End2=10 then%>Selected<%end if%>><%=myMessage_October%></OPTION> 
<OPTION VALUE="11" <%if myMonth_End2=11 then%>Selected<%end if%>><%=myMessage_November%></OPTION> 
<OPTION VALUE="12" <%if myMonth_End2=12 then%>Selected<%end if%>><%=myMessage_December%></OPTION> 
</SELECT>
&nbsp; 
<FONT FACE="Arial, Helvetica, sans-serif" SIZE="2">
<INPUT TYPE="text" SIZE="2" MAXLENGTH="2" NAME="Day_End2" VALUE="<%=myDay_End2%>"> 
, 
<INPUT TYPE="text" SIZE="4" MAXLENGTH="4" NAME="Year_End2" VALUE="<%=myYear_End2%>"> 
&nbsp;
</FONT>
</P>
&nbsp; 
<%
if (myDay_End2<>"" or (myMonth_End2<>"" and myMonth_End2<>"0") or myYear_End2<>"") AND  not  isDate(myPhase_Date_End2) then 
	%> 
	<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myError_Message_Invalid_Date%></FONT></B>
	<%
end if
%>

<%
if len(myPhase_Date_End2)>0 and len(myPhase_Date_Beginning2)>0 then
	if (isDate(myPhase_Date_End2) and isDate(myPhase_Date_Beginning2)) then 
		if cdate(myPhase_Date_End2)<cdate(myPhase_Date_Beginning2) then
		%>
		<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"><%=myError_Message_Date_End_Before_Date_Beginning%></FONT></B>
		<%	
		end if
	end if 
end if 
%>
<%
if ((len(myPhase_Date_End2)=0 or myMonth_End2="0") and len(myPhase_Date_Beginning2)>0 and isDate(myPhase_Date_Beginning2)) then
	%> 
	<B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2"> <%=myError_Message_Enter_Date_End%></FONT></B> 
	<%
end if
%>
</td>
</tr>

<%
' Leader
%>

<tr>
<td align="right" bgcolor="<%=myBorderColor%>">
<font size="2" face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>"><b><%=myMessage_Leader%>&nbsp;(<%=myMessage_Phase%>)</b></font>
</td>
<td align="Left">

<select name="Phase_Leader_ID"> 
<%
if (len(myPhase_Leader_ID_Selected)=0 or myPhase_Leader_ID_Selected=0) then 
	Response.Write "<option selected value=""0"">"&myMessage_Select&"</option>"
else
	Response.Write "<option value=""0"">"&myMessage_Select&"</option>"
end if	

' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

mySQL_Select_tb_Sites_Members = "SELECT * FROM tb_Sites_Members Where Site_ID="&mySite_ID&" ORDER BY Member_Pseudo"
	
set mySet_tb_Sites_Members = myConnection.Execute(mySQL_Select_tb_Sites_Members)

mySet_tb_Sites_Members.movefirst
do while not mySet_tb_Sites_Members.eof
	if CInt(myPhase_Leader_ID_Selected) = mySet_tb_Sites_Members("Member_ID") then
		 Response.Write "<option selected value=" & mySet_tb_Sites_Members("Member_ID") & ">" & mySet_tb_Sites_Members("Member_Pseudo") & "</option>"
	else
		Response.Write "<option value=" & mySet_tb_Sites_Members("Member_ID") & ">" & mySet_tb_Sites_Members("Member_Pseudo") & "</option>"
	end if
mySet_tb_Sites_Members.movenext
loop

mySet_tb_Sites_Members.close
Set mySet_tb_Sites_Members=nothing
myConnection.close
set myConnection = nothing
%> 
</select>
</td>
</tr>


<%
' Validation
%>

<tr>
<td align="right"  valign="top" bgcolor="<%=myBorderColor%>">
<font face="Arial, Helvetica, sans-serif" Color="<%=myBorderTextColor%>">&nbsp;</font>
</td>
<td align="left"> 
<input type="submit" value="<%=myMessage_Go%>" name="Validation">
</td>
</tr>

<%
' Date and Author
%>

<tr> 
<td align="Center" valign="top" bgcolor="<%=myApplicationColor%>" colspan="2">
<font face="Arial, Helvetica, sans-serif" size="1" color="<%=myApplicationTextColor%>"> 
<%=myDate_Display(myPhase_Date_Update,2)%> -- <%=myPhase_Author_Update%></font>
</td>
</tr>
</table>
</form>

<%
' NAVIGATION
%>

<table border="0" width="90%" cellpadding="3" cellspacing="0">
<tr>
<td>
<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="__Projects_List.asp?List=<%=myList%>&Numpage=<%=myNumPage%>&Search=<%=mySearch%>"><%=myMessage_Project%>s</a> 
<% 
if myAction="Update" then 
	%> 
	,&nbsp;
	<a href="__Phases_List.asp?Project_ID=<%=myProject_ID%>"><%=myMessage_Phase%>s</a>
	<% 
end if
%>
</font>
</td>
</tr>

<%
' Can be Modify or Delete by Project Author, Project Leader, Phase Author, Phase Leader or Administrator
%> 

<%
if myAction="Update" then 
	%> 

	<% 
	if myProject_Member_ID=myUser_ID or myProject_Leader_ID_Selected=myUser_ID or myPhase_Member_ID=myUser_ID or myPhase_Leader_ID_Selected=myUser_ID or  myUser_type_ID=1 then
		 %> 
		<tr>
		<td>
		<font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myMessage_Delete%> ?'))document.location='__Project_Modification.asp?action=Delete&Project_ID=<%=myProject_ID%>'"><%=myMessage_Delete%></a></font> 
		</td>
		</tr>
		<%
	 End IF
End If
%>
</table>
</td>
</tr>
</table>

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
</body>
</html>






<html></html>
<html><script language="JavaScript"></script></html>