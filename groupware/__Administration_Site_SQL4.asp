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

<% 	'Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Cache non géré par PWS

mySQL_Dont_Connect = 1
%>
<!-- #include file="_INCLUDE/Global_Parameters.asp" -->


<%
' ------------------------------------------------------------
' Name		: __Administration_SQL.asp
' Path		: /
' Description 	: SQL SERVER Administration Home Page
' By		: Pierre Rouarch, Dania Tcherkezoff	
' Company 	: OverApps
' Date		: December, 11, 2001
' Version   : 1.17.0
'
' Modify by	: 
' Company	:
' Date
' ------------------------------------------------------------

Dim myPage
myPage = "__Administration_Site_SQL4.asp"

%>


<!-- #include file="_INCLUDE/DB_Environment.asp" -->

<%
If myUser_Type_ID<>1 then
	response.redirect("__Home.asp")
End if
%>



<HTML><HEAD></HEAD><TITLE><%=mySite_Name%> - Site Administration </TITLE>
<BODY BackGround="<%=myBGImage%>" bgColor=<%=myBGColor%> Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">
<%
' TOP
%> <!-- #include file="_borders/Top.asp" --> <%
' CENTER
%> <TABLE WIDTH="<%=myGlobal_Width%>" BGCOLOR="<%=myBorderColor%>" BORDER="0" CELLPADDING="0" CELLSPACING="0"> 
<TR VALIGN="TOP"> <%
' CENTER LEFT
%> <TD WIDTH="<%=myLeft_Width%>"> <!-- #include file="_borders/Left.asp" --> </TD><%
'CENTER APPLICATION
%> <TD WIDTH="<%=myApplication_Width%>" BGCOLOR="<%=myBGCOLOR%>"> 
<%''''''''''''''''''''''''''''''''''START OF SCRIPT''''''''''''''''%>



<html>
<head>
	<title>SQL SERVER SETUP</title>
</head>

<body>

<%
on error resume next
set myNew_DB_Connection = Server.CreateObject("ADODB.Connection")
myNew_DB_Connection.Open myConnection_String_SQL

'Copy Table Function 
sub Table_Copy(myTable_Name,myFields_Number,myAuto_Num) 
 Dim mySQL_New_DB, mySQL_Old_DB
 Dim mySet_New_DB, mySet_Old_DB
 Dim del,x,y,i, myCurrent_Auto_Num, temp, memo, memo2

  mySQL_New_DB = "SELECT * FROM " & myTable_Name
 Set mySet_New_DB = server.createobject("adodb.recordset")
 mySet_New_DB.open mySQL_New_DB, myNew_DB_Connection, 3,3

 if myAuto_Num >= 0 Then 
 mySQL_Old_DB = "SELECT * FROM " & myTable_Name & " ORDER BY " & mySet_New_DB.fields(myAuto_Num).name
else 
 mySQL_Old_DB = "SELECT * FROM " & myTable_Name
end if 

 set mySet_Old_DB = server.createobject("adodb.recordset")
 set mySet_Old_DB = myOld_DB_Connection.Execute(mySQL_Old_DB)

 response.write myTable_Name 

 del = 0
 if not mySet_Old_DB.eof Then mySet_Old_DB.MoveFirst
 i=1

		do while not mySet_Old_DB.eof 
      if (del <> 0 ) Then 
	 		 myNew_DB_Connection.Execute("Delete From " & myTable_Name & " where "& mySet_New_DB.fields(myAuto_Num).name & " = " & del )
			  			 		 		   
      end if
			mySet_New_DB.AddNew 
			x=0   
			y=0
			del = 0   
		  if	MyAuto_Num >= 0  Then 
					myCurrent_Auto_Num = mySet_Old_DB.fields(myAuto_Num)
			else 
			    myCurrent_Auto_Num = -1
			end if
			
			if i = myCurrent_Auto_Num or myCurrent_Auto_Num < 0 Then
			 do while (x < myFields_Number)           
					if x <> myAuto_Num And  mySet_New_DB.fields(x).name <> "Meeting.Public" Then

					   temp =mySet_Old_DB.fields(mySet_New_DB.fields(x).name) 
					    
						if mySet_New_DB.fields(x).name = "NewsGroup_Message"     Then memo  = temp
						if mySet_New_DB.fields(x).name = "Contact_Comments"      Then memo  = temp
						if mySet_New_DB.fields(x).name = "Web_Description_Long"  Then memo  = temp
						if mySet_New_DB.fields(x).name = "Meeting_Agenda"        Then memo  = temp
						if mySet_New_DB.fields(x).name = "Meeting_Comments"      Then memo2 = temp
						if mySet_New_DB.fields(x).name = "New_Description_Long"  Then memo  = temp
						if mySet_New_DB.fields(x).name = "Member_Comments"       Then memo  = temp
					
						
					 'CHECK FOR DATE AND ADAPT THEM
					
					 
						  mySet_New_DB(x) = temp
					 
					end if	
															
					'response.write "<br>" & mySet_New_DB.fields(x).name & " = = "& mySet_New_DB(mySet_New_DB.fields(x).name) 
             		x =x + 1		 		
					y=y+1	
						
			 loop 
				
		mySet_Old_DB.movenext
      	 
	     del = 0 
			 else del= i
			 end if  			  
			 
			 If myTable_Name = "tb_Newsgroups_messages" Then
			  mySet_New_DB.fields("NewsGroup_Message") = memo
			 end if
			 				 
			 If myTable_Name = "tb_contacts" Then
			  mySet_New_DB.fields("Contact_Comments") = memo
			 end if
			 
			 If myTable_Name = "tb_Webs" Then
			  mySet_New_DB.fields("Web_Description_Long") = memo
			 end if
			 
			 If myTable_Name = "tb_Meetings" Then
			  mySet_New_DB.fields("Meeting_Comments") = memo
			 end if
			 
			 If myTable_Name = "tb_Meetings" Then
			  mySet_New_DB.fields("Meeting_Agenda") = memo2
			 end if
			 
			 If myTable_Name = "tb_News" Then
			  mySet_New_DB.fields("New_Description_Long") = memo
			 end if
			 
			 If myTable_Name = "tb_Sites_Members" Then
			  mySet_New_DB.fields("Member_Comments") = memo
			 end if
			
			 If myTable_Name = "tb_Meetings" Then
			  mySet_New_DB.fields("Meeting_Public") = 0
			 end if			 					 
			 
			 mySet_New_DB.Update
			  
 	  	 i=i+1			
  	loop		   		 
       
		mySet_Old_DB.close
		mySet_New_DB.close
		
		set	mySet_Old_DB=nothing
		set	mySet_New_DB=nothing 
		Response.Write "...." & myDB_Message_Copied & "<br>"	 
end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Copy Table Function 2 For 1.13.X and 1.14.X Version
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub Table_Copy2(myTable_Name,myFields_Number,myAuto_Num) 
 Dim mySQL_New_DB, mySQL_Old_DB
 Dim mySet_New_DB, mySet_Old_DB
 Dim del,x,y,i, myCurrent_Auto_Num, temp, memo,memo2

 mySQL_New_DB = "SELECT * FROM " & myTable_Name
 Set mySet_New_DB = server.createobject("adodb.recordset")
 mySet_New_DB.open mySQL_New_DB, myNew_DB_Connection, 3,3

 if myAuto_Num >= 0 Then 
  mySQL_Old_DB = "SELECT * FROM " & myTable_Name & " ORDER BY " & mySet_New_DB.fields(myAuto_Num).name
 else 
  mySQL_Old_DB = "SELECT * FROM " & myTable_Name
 end if 

 set mySet_Old_DB = server.createobject("adodb.recordset")
 set mySet_Old_DB = myOld_DB_Connection.Execute(mySQL_Old_DB)

 response.write myTable_Name 

 del = 0
 if not mySet_Old_DB.eof Then mySet_Old_DB.MoveFirst
 i=1

  do while not mySet_Old_DB.eof AND i < 500
      if (del <> 0 ) Then 
	 		 myNew_DB_Connection.Execute("Delete From " & myTable_Name & " where "& mySet_New_DB.fields(myAuto_Num).name & " = " & del )
			  			 		 		   
      end if
      mySet_New_DB.AddNew 
	  x=0   
	  y=0
	  del = 0   
	  if MyAuto_Num >= 0  Then 
		myCurrent_Auto_Num = mySet_Old_DB.fields(myAuto_Num)
	  else 
		myCurrent_Auto_Num = -1
	  end if
			
	  if i = myCurrent_Auto_Num or myCurrent_Auto_Num < 0 Then
	   do while (x < myFields_Number)           
		if x <> myAuto_Num AND mySet_New_DB.fields(x).name <> "Site_ID" Then
		   temp =mySet_Old_DB.fields(mySet_New_DB.fields(x).name)
           if mySet_New_DB.fields(x).name = "File_Long_Description"  Then memo = temp
		   if mySet_New_DB.fields(x).name = "Folder_Long_Description"  Then memo = temp
				
					 
					 mySet_New_DB(x) = temp					 
				end if	 	
																				
             	    x = x + 1		 		
					y = y + 1							
			 loop 
			 
		mySet_Old_DB.movenext
      	 
	     del = 0 
			 else del= i
			 end if  			  
			 
			 If myTable_Name = "tb_Files" Then
			  mySet_New_DB.fields("File_Long_Description") = memo
			 end if
			 
			 If myTable_Name = "tb_Folders" Then
			  mySet_New_DB.fields("Folder_Long_Description") = memo
			 end if
			 
		 mySet_New_DB.Update
			  			 			 			
 	  	 i=i+1			
  	loop		   		 
        
		mySet_Old_DB.close
		mySet_New_DB.close
		
		set	mySet_Old_DB=nothing
		set	mySet_New_DB=nothing 
		Response.Write "...." & myDB_Message_Copied & "<br>"	 

end sub




'on error resume next



Dim myAction, myFile_System_Object,myConnection_SQL
''''''''''''''''''''''''
'GET PARAMETERS        '
''''''''''''''''''''''''
myAction = Request.QueryString("Action")

'Connect to db


set myConnection_SQL = Server.CreateObject("ADODB.Connection")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DATA IMPORTATION FROM OLD BASES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myAction = "import" then

myVersion = Request.QueryString("Version")

'SET OLD DBB PATH 
myOld_Database_Path = server.mapPath("..") & "\DB\empty.mdb"


'OPEN THE OLD DB
set myOld_DB_Connection = Server.CreateObject("ADODB.Connection")
myOld_DB_Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & myOld_Database_Path
'OPEN NEW DB



'INSERT FIELDS FOR TB_APPLICATIONS
 Set mySet_New_DB = server.createobject("adodb.recordset")
 mySet_New_DB.open "SELECT * from tb_Applications", myNew_DB_Connection, 3,3
 
 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  = 1 
 mySet_New_DB("Application_Name")            = "Agenda"
 mySet_New_DB("Application_Presentation")    = "Shared Agenda between Members of a site"
 mySet_New_Db("Application_Title")           = "Agenda Title"
 mySet_New_DB("Application_Entry_Page")      = "__Agenda_Day.asp"
 mySet_New_DB("Application_Include_Box")     = "__Agenda_Box.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update             
 
 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  = 2 
 mySet_New_DB("Application_Name")            = "Projects"
 mySet_New_DB("Application_Presentation")    = "Planning Projects"
 mySet_New_DB("Application_Title")           = "Projects Title"
 mySet_New_DB("Application_Entry_Page")      = "__Projects_List.asp"
 mySet_New_DB("Application_Include_Box")     = "__Projects_Box.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update             

 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  = 3 
 mySet_New_DB("Application_Name")            = "Members"
 mySet_New_DB("Application_Presentation")    = "Members visulation for members"
 mySet_New_DB("Application_Title")           = "Members Title"
 mySet_New_DB("Application_Entry_Page")      = "__Sites_Members_List.asp"
 mySet_New_DB("Application_Include_Box")     = "__Members_Box"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update  


 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  = 4 
 mySet_New_DB("Application_Name")            = "Contacts"
 mySet_New_DB("Application_Presentation")    = "Contacts Directory"
 mySet_New_DB("Application_Title")           = "Contacts Title"
 mySet_New_DB("Application_Entry_Page")      = "__Contacts_List.asp"
 mySet_New_DB("Application_Include_Box")     = "__Contacts_box.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update   

 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID") 			 = 5 
 mySet_New_DB("Application_Name")            = "Webs"
 mySet_New_DB("Application_Presentation")    = "Web Directory"
 mySet_New_DB("Application_Title")           = "Web Title"
 mySet_New_DB("Application_Entry_Page")      = "__Webs_List.asp"
 mySet_New_DB("Application_Include_Box")     = "__Webs_Box.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update    

 mySet_New_DB.AddNew   
 mySet_New_DB("Application_ID")  = 6 
 mySet_New_DB("Application_Name")            = "News"
 mySet_New_DB("Application_Presentation")    = "News and NewsWires"
 mySet_New_DB("Application_Title")           = "News Title"
 mySet_New_DB("Application_Entry_Page")      = "__News_List.asp"
 mySet_New_DB("Application_Include_Box")     = "__News_box.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update    

 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")              = 7
 mySet_New_DB("Application_Name")            = "Events"
 mySet_New_DB("Application_Presentation")    = "Calendars and events"
 mySet_New_DB("Application_Title")           = "Events Titles"
 mySet_New_DB("Application_Entry_Page")      = "__Events_List.asp"
 mySet_New_DB("Application_Include_Box")     = "__Events_Box.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1 
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update    

 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  = 8 
 mySet_New_DB("Application_Name")            = "Newsgroups"
 mySet_New_DB("Application_Presentation")    = "Sample Newsgroup"
 mySet_New_DB("Application_Title")           = "Forum Title"
 mySet_New_DB("Application_Entry_Page")      = "__Newsgroup_Messages_List.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update    

 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  = 9
 mySet_New_DB("Application_Name")            = "Styles"
 mySet_New_DB("Application_Presentation")    = "Styles"
 mySet_New_DB("Application_Title")           = "Styles Title"
 mySet_New_DB("Application_Entry_Page")      = "__Styles_List.asp" 
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1 
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update    

 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  = 10
 mySet_New_DB("Application_Name")            = "Files"
 mySet_New_DB("Application_Presentation")    = "Shared Files"
 mySet_New_DB("Application_Title")           = "Files Title"
 mySet_New_DB("Application_Entry_Page")      = "__Folders_List.asp"
 mySet_New_DB("Application_Include_Box")     = "__Files_Box.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3 
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update    

 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  = 11
 mySet_New_DB("Application_Name")            = "About"
 mySet_New_DB("Application_Presentation")    = "Intranet Général Informations" 
 mySet_New_DB("Application_Title")           = "About Title"
 mySet_New_DB("Application_Entry_Page")      = "__Intranet_Information.asp"
 mySet_New_DB("Application_Public_Type_ID")  = 3
 mySet_New_DB("Application_Opened")          = 1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update    

 mySet_New_DB.AddNew
 mySet_New_DB("Application_ID")  =  12
 mySet_New_DB("Application_Name")            =  "Quit"
 mySet_New_DB("Application_Presentation")    =  "Disconnect"
 mySet_New_DB("Application_Title")           =  "Quit Title"
 mySet_New_DB("Application_Entry_Page")      =  "__Quit.asp"
 mySet_New_DB("Application_Public_Type_ID")  =  3
 mySet_New_DB("Application_Opened")          =  1
 mySet_New_DB("Application_Author_Creation") = "Administrator"
 mySet_New_DB("Application_Date_Creation")   = "01/01/01"
 mySet_New_DB.Update    



 




'COPY THE 'TABLES FROM OLD TO THE NEW DB

Table_Copy "tb_Newsgroups_messages",10,3
Table_Copy "tb_projects",18,3
Table_Copy "tb_Newsgroups_Sites",6,1
Table_Copy "tb_Meetings_members", 4,-1
Table_Copy "tb_Meetings",16,4
Table_Copy "tb_Sites_members",40,1
Table_Copy "tb_Calendars_Members",9,-1
Table_Copy "tb_Calendars",10,2
Table_Copy "tb_Calendars_Sites",3,-1
Table_Copy "tb_Calendars_Types",3,0
Table_Copy "tb_Contacts_Activities",6,1
Table_Copy "tb_Countries",3,-1
Table_Copy "tb_Directories",10,2
Table_Copy "tb_Directories_members",9,-1
Table_Copy "tb_Directories_Sites",3,-1
Table_Copy "tb_Directories_Types",3,0
Table_Copy "tb_Events",10,3
Table_Copy "tb_News",13,3
Table_Copy "tb_Newsgroups",9,3
Table_Copy "tb_NewsWires",10,2
Table_Copy "tb_NewsWires_Members",9,-1
Table_Copy "tb_NewsWires_Sites",3,-1
Table_Copy "tb_NewsWires_Types",3,0
Table_Copy "tb_phases",19,4
Table_Copy "tb_Phases_Members",3,-1
Table_Copy "tb_Phases_Priorities",6,4
Table_Copy "tb_Phases_status",6,4
Table_Copy "tb_Projects_Members",3,-1
Table_Copy "tb_Projects_Priorities",6,4
Table_Copy "tb_Projects_Status",6,4
Table_Copy "tb_Sites_Activities",6,1
Table_Copy "tb_Sites_Members_Activities",6,1
Table_Copy "tb_tasks",20,5
Table_Copy "tb_Tasks_Members",3,-1
Table_Copy "tb_Tasks_Priorities",6,4
Table_Copy "tb_Tasks_Status",6,4
Table_Copy "tb_Webdirectories",10,0
Table_Copy "tb_Webdirectories_Members",9,-1
Table_Copy "tb_Webdirectories_Sites",3,-1
Table_Copy "tb_Webdirectories_types",3,0
Table_Copy "tb_Webs_Categories",7,0
Table_Copy "tb_contacts",35,3 
Table_Copy "tb_Webs",13,4





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SPECIAL ADAPTATION FOR 1.13.X and 1.14.X
if myVersion < 20 then


 Table_Copy "tb_Styles",19,2

 
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TB SITE APPLICATIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 

 Table_Copy "tb_Sites_Applications",7,-1 



'SPECIAL FUNCTION FOR TB_SITES
''''''''''''''''''''''''''''''''''''''''''''''''''''
myTable_Name = "tb_Sites"
myFields_Number = 27
myAuto_Num = 3


mySQL_New_DB = "SELECT * FROM " & myTable_Name
Set mySet_New_DB = server.createobject("adodb.recordset")
mySet_New_DB.open mySQL_New_DB, myNew_DB_Connection, 3,3

mySQL_Old_DB = "SELECT * FROM " & myTable_Name  & " ORDER BY Site_ID "

Set mySet_Old_DB = server.createobject("adodb.recordset")
mySet_Old_DB.open mySQL_Old_DB, myOld_DB_Connection, 3,3
set mySet_Old_DB = myOld_DB_Connection.Execute(mySQL_Old_DB)

response.write myTable_Name 

del = 0
mySet_Old_DB.MoveFirst
i=1

		do while not mySet_Old_DB.eof 

      if (del <> 0 ) 	 Then 
	 		 myNew_DB_Connection.Execute("Delete From " & myTable_Name & " where "& mySet_New_DB.fields(myAuto_Num).name & " = " & del )				 		 		   
      end if
			mySet_New_DB.AddNew 
			x=0   
			y=0
			del = 0   
		  if	MyAuto_Num >= 0  Then 
					myCurrent_Auto_Num = mySet_Old_DB.fields(myAuto_Num)
			else 
			    myCurrent_Auto_Num = -1
			end if
			
			if i = myCurrent_Auto_Num or myCurrent_Auto_Num < 0 Then

			mySet_New_DB.fields("Engine_ID") =   mySet_Old_DB.fields("Engine_ID") 
			mySet_New_DB.fields("Distributor_ID") = mySet_Old_DB.fields("Distributor_ID")
			mySet_New_DB.fields("NetWork_ID") = mySet_Old_DB.fields("NetWork_ID")
			mySet_New_DB.fields("Site_URL") = mySet_Old_DB.fields("Site_URL")
			mySet_New_DB.fields("Site_Name") = mySet_Old_DB.fields("Site_Name")
			mySet_New_DB.fields("Site_Public_Type_ID") = 3
			mySet_New_DB.fields("Site_Presentation") = mySet_Old_DB.fields("Site_Presentation")
			mySet_New_DB.fields("Site_Language_ID") = mySet_Old_DB.fields("Site_Language_ID")
			mySet_New_DB.fields("Site_Use_ID") = mySet_Old_DB.fields("Site_Use_ID")
			mySet_New_DB.fields("Site_Company") = mySet_Old_DB.fields("Site_Company")
			mySet_New_DB.fields("Site_Address") = mySet_Old_DB.fields("Site_Address")
			mySet_New_DB.fields("Site_zip") = mySet_Old_DB.fields("Site_zip")
			mySet_New_DB.fields("Site_State") = mySet_Old_DB.fields("Site_State")
			mySet_New_DB.fields("Site_Country_ID") = mySet_Old_DB.fields("Site_Country_ID")
			mySet_New_DB.fields("Site_Phone") = mySet_Old_DB.fields("Site_Phone")
			mySet_New_DB.fields("Site_Fax") = mySet_Old_DB.fields("Site_Fax")
			mySet_New_DB.fields("Site_Web") = mySet_Old_DB.fields("Site_Web")
			mySet_New_DB.fields("Site_Email") = mySet_Old_DB.fields("Site_Email")
			mySet_New_DB.fields("Site_Style_ID") = mySet_Old_DB.fields("Site_Style_ID")
			mySet_New_DB.fields("Site_Public_Directory") = mySet_Old_DB.fields("Site_Public_Directory")
			mySet_New_DB.fields("Site_Author_Update") = mySet_Old_DB.fields("Site_Author_Update")
			mySet_New_DB.fields("Site_Date_Update") = Now()
			mySet_New_DB.fields("Site_Maximum_Files_Size") = mySet_Old_DB.fields("Site_Maximum_Files_Size")
			mySet_New_DB.fields("Site_Date_Format")       = mySet_Old_DB.fields("Site_Date_format")
			mySet_New_DB.fields("Site_Hour_Format")       = mySet_Old_DB.fields("Site_Hour_Format")
			mySet_New_DB.fields("Site_Agenda_Start")      = 7
			mySet_New_DB.fields("Site_Agenda_End")        = 20
			mySet_New_DB.fields("Site_Agenda_Week_Start") = 0
			
			
      mySet_Old_DB.movenext
      	 
	     del = 0 
			 else del = i
			 end if  			  
			 mySet_New_DB.Update			
 	  	 i=i+1			
    	loop		   		 
 
		mySet_Old_DB.close
		mySet_New_DB.close
 		
		set	mySet_Old_DB=nothing
		set	mySet_New_DB=nothing

Table_Copy2 "tb_Files" , 12 , 0
Table_Copy "tb_Files_Extensions" , 4 , 0
Table_Copy2 "tb_Folders" , 10 , 0
Table_Copy2 "tb_Folders_Access", 3 , -1



end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myNew_DB_Connection.Close
myOld_DB_Connection.Close

%>
<script language="javascript">
<!--
top.location.href="__Administration_Site.asp"	
-->
</script>
<%

end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'END IMPORTATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



%>
<table width=100%>
          <tr> 
            <td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"><b><font face="Arial, Helvetica, sans-serif" size="3" color="<%=myApplicationTextColor%>"> 
              SQL SERVER</font></b> </td>
          </tr>
</table>	<br><br>
<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myMessage_Warning3%></font>
<br><br>
&nbsp;<font face="Arial, Helvetica, sans-serif" size="3" color="<%=myBGTextColor%>"><b><%=myMessage_Choose%> :</b></font><br>

<%
 If myAction <> "Show_Import" OR myAction <> "New" Then
%>

<%

'TEST ON DATABASES FILES
myFile_Test = 1
set myFile_System_Object=server.createobject("scripting.FileSystemObject")
myVersion=0


if myFile_System_Object.FileExists(myOld_Database_Path_OS & "overapps-software.mdb") Then
 %> 
 <div align=left>&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?Version=11';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.9.5)</b></font></A></div><br>
 <%
 myVersion   = 11 
end if 

if myFile_System_Object.FileExists(myOld_Database_Path_OS & "overapps-software_V1110.mdb") Then 
 myVersion   = 12 
%> 
 <div align=left>&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?Version=<%= myVersion %>';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.11.0)</b></font></A></div>
<br>
<%
end if
 
  
 if myFile_System_Object.FileExists(myOld_Database_Path_OS & "overapps-software_V1111.mdb") Then 
  myVersion=13
   
 %> 
 <div align=left>&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?Version=<%= myVersion %>';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.11.1)</b></font></A></div>
<br>
<%
end if

if myFile_System_Object.FileExists(myOld_Database_Path_GW & "overapps_V113X.mdb") Then
 %> 
 <div align=left>&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?Version=14';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.13.X)</b></font></A></div><br>
 <%
 myVersion   = 14
end if 

if myFile_System_Object.FileExists(myOld_Database_Path_GW & "overapps_V114X.mdb") Then
 %> 
 <div align=left>&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?Version=15';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.14.X)</b></font></A></div><br>
 <%
 myVersion   = 15
end if 

if myFile_System_Object.FileExists(myOld_Database_Path_GW & "overapps_V115X.mdb") Then
 %> 
 <div align=left>&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?Version=16';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.15.X)</b></font></A></div><br>
 <%
 myVersion   = 16
end if 

if myFile_System_Object.FileExists( myOld_Database_Path_GW &  "overapps_V116X.mdb") Then
 %> 
 <div align=left>&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?Version=17';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> (1.16.X)</b></font></A></div><br>
 <%
 myVersion   = 17
end if 

if myFile_System_Object.FileExists( Server.MapPath("..") & "\DB\Overapps_V118X.mdb") Then
 %> 
 <div align=left>&nbsp;&nbsp;<a href="Javascript:if(confirm('<%=myDB_Message_Question%>'))document.location='__DB_Update_Begin.asp?Version=18';"><font face="Arial, Helvetica, sans-serif" size="2" color="<%=MyBGTextColor%>"><b><%=myDB_Message_Import%> &nbsp;<%=myMessage_current_version%> (1.18.X)</b></font></A></div><br>
 <%
 myVersion   = 18
end if 



if myVersion = 0 Then
%>
<div align=left><font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><% =myDB_Message_CheckFile %></font></div><br>
<%
end if



set myFile_System_Object=nothing
%>

<%
end if
%>



<%''''''''''''''''''''''''''''''''''END OF SCRIPT''''''''''''''''''''%>
</TD></TR> <%
' / CENTER APPLICATION
%> </TABLE><%
' /CENTER
%> <%
' DOWN
%> <!-- #include file="_borders/Down.asp" --> 
<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.							'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</BODY>
</HTML>

<html><script language="JavaScript"></script></html>