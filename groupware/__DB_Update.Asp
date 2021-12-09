<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001  - OverApps - http://www.overapps.com
'
' This program "__Contact_Information.asp" is free software; 
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
' 	" Copyright (C) 2001 OverApps & contributors "
' at the bottom of the page with an active link from the name "OverApps" to 
' the Address http://www.overapps.com where the user (netsurfer) could find the 
' appropriate information about this license. 
'
'-----------------------------------------------------------------------------
%>

<% 
	'Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1 
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Cache non géré par PWS
mySQL_Dont_Connect = 1
%>

<%
' ------------------------------------------------------------
' Name			: __DB_Update.asp
' Path   		: /
' Vertsion 		: 1.15.0
' Description 	: Tools for updating old DB to version 1.11.0
' By			  : Dania TCHERKEZOFF												
' Company		: OverApps
' Date			: November 21, 2001
' ------------------------------------------------------------

Dim myPage
myPage = "__DB_Update.asp"

	
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
Dim myNew_DB_Connection, myOld_DB_Connection
Dim myTable_Name,myFields_Number,myAuto_Num 
Dim mySQL_New_DB, mySQL_Old_DB
Dim mySet_New_DB, mySet_Old_DB
Dim del,x,y,i,j,myCurrent_Auto_Num
Dim myFile_System_Object
Dim date_temp, myVersion

Dim myOpen_Agenda(20)    
Dim myOpen_Projects(20)  
Dim myOpen_Tasks(20)     
Dim myOpen_Members(20)   
Dim myOpen_Contacts(20)  
Dim myOpen_Webs(20)      
Dim myOpen_News(20)      
Dim myOpen_Events(20)    
Dim myOpen_NewsGroups(20)

 


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CHECK IF THE USER CAN ENTER IN THIS APPLICATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if myUser_Type_ID <> 1 then
	Response.redirect("__Quit.asp")
end if


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET PARAMETERS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET VERSION OF THE OLD BASE
' 1 for version  1.9.5 AND PREVIOUS
' 2 For 1.10.X 
' 3   For 1.11.X
myVersion = Request.QueryString("Version")
if len(myVersion)= 0 then Response.Redirect "__DB_Update_adv.asp"

'SET OLD DBB PATH DEPENDING ON VERSION

If myVersion = 1 or myVersion =11 Then myOld_Database_Path = myOld_Database_Path_OS & "overapps-software.mdb"
If myVersion = 2 or myVersion =12 Then myOld_Database_Path = myOld_Database_Path_OS & "overapps-software_v1110.mdb"
If myVersion = 3 or myVersion =13 Then myOld_Database_Path = myOld_Database_Path_OS & "overapps-software_v1111.mdb"
If myVersion = 4 or myVersion =14 Then myOld_Database_Path = myOld_Database_Path_GW & "overapps_v113x.mdb"
If myVersion = 5 or myVersion =15 Then myOld_Database_Path = myOld_Database_Path_GW & "overapps_v114x.mdb"
If myVersion = 6 or myVersion =16 Then myOld_Database_Path = myOld_Database_Path_GW & "overapps_v115x.mdb"
If myVersion = 7 or myVersion =17 Then myOld_Database_Path = myOld_Database_Path_GW & "overapps_v116x.mdb"
If myVersion = 8 or myVersion =18 Then myOld_Database_Path = Server.MapPath("..") & "\DB\Overapps_V118X.mdb"

If myVersion > 10 Then  
 myNew_Database_Path = server.mapPath("..") & "\DB\empty.mdb"
end if 



'IF OPERATION AS ALREADY BE DONE, OLD FILE IS RENANMED
set myFile_System_Object=server.createobject("scripting.FileSystemObject")

if myversion < 10 Then
 
 if Request.QueryString("done") < 0 Then
  myFile_System_Object.MoveFile  myNew_Database_Path & ".bak", myNew_Database_Path
 end if
 
 myFile_System_Object.CopyFile   myNew_Database_Path , myNew_Database_Path & ".bak"
end if

if myVersion > 10 Then
 if myFile_System_Object.FileExists(myNew_Database_Path) Then myFile_System_Object.DeleteFile(myNew_Database_Path)
 myFile_System_Object.CopyFile myNew_Database_Path&".bak",myNew_Database_Path 
end if

set myFile_System_Object=nothing
 
 'OPEN THE NEW DB
set myNew_DB_Connection = Server.CreateObject("ADODB.Connection")
myNew_DB_Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & myNew_Database_Path

'OPEN THE OLD DB
set myOld_DB_Connection = Server.CreateObject("ADODB.Connection")
myOld_DB_Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & myOld_Database_Path


'Copy Table Function 
sub Table_Copy(myTable_Name,myFields_Number,myAuto_Num) 
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
					    
						if mySet_New_DB.fields(x).name = "NewsGroup_Message"  Then memo = temp
						if mySet_New_DB.fields(x).name = "Contact_Comments"  Then memo = temp
						if mySet_New_DB.fields(x).name = "Web_Description_Long"  Then memo = temp
						if mySet_New_DB.fields(x).name = "Meeting_Agenda"  Then memo = temp
						if mySet_New_DB.fields(x).name = "Meeting_Comments"  Then memo2 = temp
						if mySet_New_DB.fields(x).name = "New_Description_Long"  Then memo = temp
						if mySet_New_DB.fields(x).name = "Member_Comments"  Then memo = temp
					
						
					 'CHECK FOR DATE AND ADAPT THEM
					 if IsDate(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)) AND (mySet_New_DB.fields(x).name) <> "NewsGroup_Message_Thread"  AND (mySet_New_DB.fields(x).name) <> "Project_Path" Then
					   
					   date_temp = myDate_Construct(Year(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),Month(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),Day(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),hour(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),minute(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),second(mySet_Old_DB.fields(mySet_New_DB.fields(x).name))) 

					   if len(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)) <= 10 Then Date_temp = left(date_temp,10)
					   mySet_New_DB.fields(x) =date_temp
			    	 
					 else
					 
						  mySet_New_DB(x) = temp
					 
					 end if
					end if	
					
										
					response.write "<br>" & mySet_New_DB.fields(x).name & " = = "& mySet_New_DB(mySet_New_DB.fields(x).name) 
             		
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
		  if	MyAuto_Num >= 0  Then 
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
						
					 'CHECK FOR DATE AND ADAPT THEM
					 if IsDate(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)) AND (mySet_New_DB.fields(x).name) <> "NewsGroup_Message_Thread"  AND (mySet_New_DB.fields(x).name) <> "Project_Path" Then
					   
					   date_temp = myDate_Construct(Year(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),Month(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),Day(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),hour(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),minute(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)),second(mySet_Old_DB.fields(mySet_New_DB.fields(x).name))) 

					   if len(mySet_Old_DB.fields(mySet_New_DB.fields(x).name)) <= 10 Then Date_temp = left(date_temp,10)
					   mySet_New_DB.fields(x) =date_temp
			    	 
					 else
					 
						  mySet_New_DB(x) = temp
					 
					 end if
					end if	
					
										
					response.write "<br>" & mySet_New_DB.fields(x).name & " = = "& mySet_New_DB(mySet_New_DB.fields(x).name) 
             		
					  x =x + 1		 		
					  y=y+1	
						
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

'COPY THE 'TABLES FROM OLD TO THE NEW DB

Table_Copy "tb_Newsgroups_messages",10,3
Table_Copy "tb_projects",18,3
'Table_Copy "tb_Newsgroups_Sites",6,1
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
'SPECIAL ADAPTATION FOR 1.10.X and 1.11.X
if myVersion = 2 OR myVersion = 3 or myVersion = 12 OR myVersion = 13 Then


 Table_Copy "tb_Styles",19,2

 
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TB SITE APPLICATIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 

 Table_Copy "tb_Sites_Applications",7,-1 

 myNew_DB_Connection.Execute ("Delete from tb_Sites_Applications where Application_ID = 10")


 'NEED TO ADD 3 FIELDS  
 myTable_Name = "tb_Sites_Applications"
 mySQL_New_DB = "SELECT * FROM " & myTable_Name
 Set mySet_New_DB = server.createobject("adodb.recordset")
 mySet_New_DB.open mySQL_New_DB, myNew_DB_Connection, 3,3
 
 mySet_New_DB.AddNew
 
  mySet_New_DB.fields("Site_ID") = 1
  mySet_New_DB.fields("Application_ID") = 10
  mySet_New_DB.fields("Site_Application_Title") = "Files"
  mySet_New_DB.fields("Site_Application_Opened") = 1	
  mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3
  mySet_New_DB.fields("Site_Application_Author_Update") = "Administrator"
  mySet_New_DB.fields("Site_Application_Date_Update")  = myDate_Now()
 
  mySet_New_DB.Update
 
 
  mySet_New_DB.AddNew
 
  mySet_New_DB.fields("Site_ID") = 1
  mySet_New_DB.fields("Application_ID") = 11
  mySet_New_DB.fields("Site_Application_Title") = "About"
  mySet_New_DB.fields("Site_Application_Opened") = 1	
  mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3
  mySet_New_DB.fields("Site_Application_Author_Update") = "Administrator"
  mySet_New_DB.fields("Site_Application_Date_Update") =  myDate_Now()
 
  mySet_New_DB.Update
 
  mySet_New_DB.AddNew
 
  mySet_New_DB.fields("Site_ID") = 1
  mySet_New_DB.fields("Application_ID") = 12
  mySet_New_DB.fields("Site_Application_Title") = "Quit"
  mySet_New_DB.fields("Site_Application_Opened") = 1	
    mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3
  mySet_New_DB.fields("Site_Application_Author_Update") = "Administrator"
  mySet_New_DB.fields("Site_Application_Date_Update") = myDate_Now()
 
  mySet_New_DB.Update
 
 
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
			mySet_New_DB.fields("Site_Extranet_Code") = 0
    		mySet_New_DB.fields("Site_Presentation") = mySet_Old_DB.fields("Site_Presentation")
			mySet_New_DB.fields("Site_Use_ID") = mySet_Old_DB.fields("Site_Use_ID")
			mySet_New_DB.fields("Site_Company") = mySet_Old_DB.fields("Site_Company")
			mySet_New_DB.fields("Site_Address") = mySet_Old_DB.fields("Site_Address")
			mySet_New_DB.fields("Site_Language_ID") = mySet_Old_DB.fields("Site_language_ID")
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
			mySet_New_DB.fields("Site_Date_Update") = myDate_Now()
			mySet_New_DB.fields("Site_Maximum_Files_Size") = 2097152
			mySet_New_DB.fields("Site_Date_Format")       = 1
			mySet_New_DB.fields("Site_Hour_Format")       = 1
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
 
		set	mySet_Old_DB=nothing
		set	mySet_New_DB=nothing

end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'SPECIAL ADAPTATION FOR 1.9.5 and earlier
if myVersion = 1  or myVersion = 11 Then

''''''''''''''''''''''''''''''''''''''''''
'INSERTING 4 STYLES
''''''''''''''''''''''''''''''''''''''''''


 myTable_Name = "tb_Styles"
 mySQL_New_DB = "SELECT * FROM " & myTable_Name
 Set mySet_New_DB = server.createobject("adodb.recordset")
 mySet_New_DB.open mySQL_New_DB, myNew_DB_Connection, 3,3

 
 
 mySet_New_DB.AddNew
 mySet_New_DB.fields("Site_ID") = 1
 mySet_New_DB.fields("Member_ID") = 2
 mySet_New_DB.fields("Style_Name")= "Overapps"
 mySet_New_DB.fields("Style_Global_Width")= 768
 mySet_New_DB.fields("Style_Left_Width")= 150
 mySet_New_DB.fields("Style_Application_Width")=618
 mySet_New_DB.fields("Style_Right_Width")=0
 mySet_New_DB.fields("Style_BGColor")="#FFFFFF"
 mySet_New_DB.fields("Style_BGImage")= ""
 mySet_New_DB.fields("Style_BGTextColor")="#000000"
 mySet_New_DB.fields("Style_BorderTextColor")="#000000"
 mySet_New_DB.fields("Style_BorderColor")="#cccc00"
 mySet_New_DB.fields("Style_BorderImage")=""
 mySet_New_DB.fields("Style_ApplicationColor")="#663399"
 mySet_New_DB.fields("Style_ApplicationTextColor")="#FFFFFF"
 mySet_New_DB.fields("Style_ApplicationImage")= ""
 mySet_New_DB.fields("Style_Author_Update")="Administrator"
 mySet_New_DB.fields("Style_Date_Update")=myDate_Now()
 mySet_New_DB.Update
 
 mySet_New_DB.AddNew
 mySet_New_DB.fields("Site_ID") = 1
 mySet_New_DB.fields("Member_ID") = 2
 mySet_New_DB.fields("Style_Name")= "Overapps 1024"
 mySet_New_DB.fields("Style_Global_Width")= 990
 mySet_New_DB.fields("Style_Left_Width")= 150
 mySet_New_DB.fields("Style_Application_Width")=840
 mySet_New_DB.fields("Style_Right_Width")=0
 mySet_New_DB.fields("Style_BGColor")="#FFFFFF"
 mySet_New_DB.fields("Style_BGImage")= ""
 mySet_New_DB.fields("Style_BGTextColor")="#000000"
 mySet_New_DB.fields("Style_BorderTextColor")="#000000"
 mySet_New_DB.fields("Style_BorderColor")="#cccc00"
 mySet_New_DB.fields("Style_BorderImage")=""
 mySet_New_DB.fields("Style_ApplicationColor")="#663399"
 mySet_New_DB.fields("Style_ApplicationTextColor")="#FFFFFF"
 mySet_New_DB.fields("Style_ApplicationImage")= ""
 mySet_New_DB.fields("Style_Author_Update")="Administrator"
 mySet_New_DB.fields("Style_Date_Update")=myDate_Now()
 mySet_New_DB.Update

mySet_New_DB.AddNew
 mySet_New_DB.fields("Site_ID") = 1
 mySet_New_DB.fields("Member_ID") = 2
 mySet_New_DB.fields("Style_Name")= "Blue Grey"
 mySet_New_DB.fields("Style_Global_Width")= 768
 mySet_New_DB.fields("Style_Left_Width")= 150
 mySet_New_DB.fields("Style_Application_Width")=618
 mySet_New_DB.fields("Style_Right_Width")=0
 mySet_New_DB.fields("Style_BGColor")="#FFFFFF"
 mySet_New_DB.fields("Style_BGImage")= ""
 mySet_New_DB.fields("Style_BGTextColor")="#000000"
 mySet_New_DB.fields("Style_BorderTextColor")="#FFFFFF"
 mySet_New_DB.fields("Style_BorderColor")="#333399"
 mySet_New_DB.fields("Style_BorderImage")=""
 mySet_New_DB.fields("Style_ApplicationColor")="#999999"
 mySet_New_DB.fields("Style_ApplicationTextColor")="#000000"
 mySet_New_DB.fields("Style_ApplicationImage")= ""
 mySet_New_DB.fields("Style_Author_Update")="Administrator"
 mySet_New_DB.fields("Style_Date_Update")=myDate_Now()
 mySet_New_DB.Update
 
  mySet_New_DB.AddNew
 mySet_New_DB.fields("Site_ID") = 1
 mySet_New_DB.fields("Member_ID") = 2
 mySet_New_DB.fields("Style_Name")= "BlueGrey1024"
 mySet_New_DB.fields("Style_Global_Width")= 990
 mySet_New_DB.fields("Style_Left_Width")= 150
 mySet_New_DB.fields("Style_Application_Width")=840
 mySet_New_DB.fields("Style_Right_Width")=0
 mySet_New_DB.fields("Style_BGColor")="#FFFFFF"
 mySet_New_DB.fields("Style_BGImage")= ""
 mySet_New_DB.fields("Style_BGTextColor")="#000000"
 mySet_New_DB.fields("Style_BorderTextColor")="#FFFFFF"
 mySet_New_DB.fields("Style_BorderColor")="#333399"
 mySet_New_DB.fields("Style_BorderImage")=""
 mySet_New_DB.fields("Style_ApplicationColor")="#999999"
 mySet_New_DB.fields("Style_ApplicationTextColor")="#000000"
 mySet_New_DB.fields("Style_ApplicationImage")= ""
 mySet_New_DB.fields("Style_Author_Update")="Administrator"
 mySet_New_DB.fields("Style_Date_Update")=myDate_Now()
 mySet_New_DB.Update
  
 
 mySet_New_DB.close()
 set  mySet_New_DB = NOTHING
 

''''''''''''''''''''''''''''''''''''''''''''''''''''
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
			mySet_New_DB.fields("Site_Presentation") = mySet_Old_DB.fields("Site_Presentation")
			mySet_New_DB.fields("Site_URL") = mySet_Old_DB.fields("Site_URL")
			mySet_New_DB.fields("Site_Name") = mySet_Old_DB.fields("Site_Name")
			mySet_New_DB.fields("Site_Public_Type_ID") = 3
			mySet_New_DB.fields("Site_Extranet_Code") = 0
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
			mySet_New_DB.fields("Site_Date_Update") = myDate_Now()
			mySet_New_DB.fields("Site_Maximum_Files_Size") = 2097152
			mySet_New_DB.fields("Site_Date_Format")       = 1
			mySet_New_DB.fields("Site_Hour_Format")       = 1
			mySet_New_DB.fields("Site_Agenda_Start")      = 7
			mySet_New_DB.fields("Site_Agenda_End")        = 20
			mySet_New_DB.fields("Site_Agenda_Week_Start") = 0
			mySet_New_DB.fields("Site_Style_ID") = 3
    		mySet_New_DB.fields("Site_Language_ID") = mySet_Old_DB.fields("Site_language_ID")
			myOpen_Agenda(i)     = mySet_Old_DB.fields("Site_Agenda_Open")
			myOpen_Projects(i)   = mySet_Old_DB.fields("Site_Projects_Open")
			myOpen_Tasks(i)      = mySet_Old_DB.fields("Site_Tasks_Open")
			myOpen_Members(i)    = mySet_Old_DB.fields("Site_Members_Open")
			myOpen_Contacts(i)   = mySet_Old_DB.fields("Site_Contacts_Open")
			myOpen_Webs(i)       = mySet_Old_DB.fields("Site_Webs_Open")
			myOpen_News(i)       = mySet_Old_DB.fields("Site_News_Open")
			myOpen_Events(i)     = mySet_Old_DB.fields("Site_Events_Open")
			myOpen_NewsGroups(i) = mySet_Old_DB.fields("Site_NewsGroups_Open")
			
			
			
      mySet_Old_DB.movenext
      	 
	     del = 0 
			 else del = i
			 end if  			  
			 mySet_New_DB.Update			
 	  	 i=i+1			
    	loop		   		 
 
		set	mySet_Old_DB=nothing
		set	mySet_New_DB=nothing
		
''''''''''''''''''''''''''''''''''''''
'OPEN APPLICATION, tb_sites_applications
''''''''''''''''''''''''''''''''''''''		
mySQL_New_DB = "SELECT * FROM tb_Sites_Applications"  
Set mySet_New_DB = server.createobject("adodb.recordset")
mySet_New_DB.open mySQL_New_DB, myNew_DB_Connection, 3,3
i = i -1
j = 1

 Do While j <= i 
 if myOpen_Agenda(i) = True OR myOpen_Agenda(i) = False Then
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Agenda"
	mySet_New_DB.fields("Application_ID") = 1
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Admninistrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
	If myOpen_Agenda(i) Then
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 
	else 	  
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1
	end if 
  mySet_New_DB.Update
  end if
	j=j+1
 loop
 
j = 1

 Do While j <= i 
 if myOpen_Projects(i) = True OR myOpen_Projects(i) = False Then
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Projects"
	mySet_New_DB.fields("Application_ID") = 2
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Admninistrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
	If myOpen_Projects(i) Then
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 
	else 	  
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1
	end if 
  mySet_New_DB.Update
  end if
	j=j+1
 loop
 
j = 1

 Do While j <= i 
 if myOpen_Members(i) = True OR myOpen_Members(i) = False Then
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Members"
	mySet_New_DB.fields("Application_ID") = 3
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Admninistrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
	If myOpen_Members(i) Then
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 
	else 	  
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1
	end if 
  mySet_New_DB.Update
  end if
	j=j+1
 loop
 
j = 1

 Do While j <= i 
 if myOpen_Contacts(i) = True OR myOpen_Contacts(i) = False Then
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Contacts"
	mySet_New_DB.fields("Application_ID") = 4
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Admninistrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
	If myOpen_Contacts(i) Then
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 
	else 	  
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1
	end if 
  mySet_New_DB.Update
  end if
	j=j+1
 loop 
 
 j = 1

 Do While j <= i 
 if myOpen_Webs(i) = True OR myOpen_Webs(i) = False Then
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Webs"
	mySet_New_DB.fields("Application_ID") = 5
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Admninistrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
	If myOpen_Webs(i) Then
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 
	else 	  
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1
	end if 
  mySet_New_DB.Update
  end if
	j=j+1
 loop
  
 j = 1

 Do While j <= i 
 if myOpen_News(i) = True OR myOpen_News(i) = False Then
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "News"
	mySet_New_DB.fields("Application_ID") = 6
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Admninistrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
	If myOpen_News(i) Then
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 
	else 	  
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1
	end if 
  mySet_New_DB.Update
  end if
	j=j+1
 loop
 
 
 j = 1

 Do While j <= i 
 if myOpen_Events(i) = True OR myOpen_Events(i) = False Then
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Events"
	mySet_New_DB.fields("Application_ID") = 7
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Admninistrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
	If myOpen_Events(i) Then
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 
	else 	  
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1
	end if 
  mySet_New_DB.Update
  end if
	j=j+1
 loop
 
j = 1

 Do While j <= i 
 if myOpen_NewsGroups(i) = True OR myOpen_NewsGroups(i) = False Then
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Forum"
	mySet_New_DB.fields("Application_ID") = 8
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Administrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
	If myOpen_Members(i) Then
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3
	else 	  
	 mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1
	end if 
  mySet_New_DB.Update
  end if
	j=j+1
 loop

j = 1
 
Do While j <= i 
 
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Styles"
	mySet_New_DB.fields("Application_ID") = 9
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Administrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
  mySet_New_DB.fields("Site_Application_Public_Type_ID") = 1 

  mySet_New_DB.Update
 
	j=j+1
 loop

j = 1
 
Do While j <= i 
 
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Quit"
	mySet_New_DB.fields("Application_ID") = 12
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Administrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
  mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 

   
  mySet_New_DB.Update
 
	j=j+1
 loop
  
  j = 1
 
Do While j <= i 
 
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "About"
	mySet_New_DB.fields("Application_ID") = 11
  mySet_New_DB.fields("Site_ID") = i  
	mySet_New_DB.fields("Site_Application_Author_Update") = "Administrator"
	mySet_New_DB.fields("Site_Application_Date_Update") = Now()
	mySet_New_DB.fields("Site_Application_Opened") = 1
  mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 

   
  mySet_New_DB.Update
 
	j=j+1
 loop
  
 j = 1
 
Do While j <= i 
 
  mySet_New_DB.AddNew
  mySet_New_DB.fields("Site_Application_Title") = "Files"
  mySet_New_DB.fields("Application_ID") = 10
  mySet_New_DB.fields("Site_ID") = i  
  mySet_New_DB.fields("Site_Application_Author_Update") = "Administrator"
  mySet_New_DB.fields("Site_Application_Date_Update") = Now()
  mySet_New_DB.fields("Site_Application_Opened") = 1
  mySet_New_DB.fields("Site_Application_Public_Type_ID") = 3 

   
  mySet_New_DB.Update
 
	j=j+1
 loop 


 set	mySet_New_DB=nothing 
 Response.Write "...." & myDB_Message_Copied & "<br>"	 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

end if





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SPECIAL ADAPTATION FOR 1.13.X and 1.14.X

if myVersion = 4 OR myVersion = 5 or myVersion = 6 or myVersion = 7 or myVersion = 14 or myVersion = 15 or myVersion = 17 or myVersion = 16  or myVersion = 8 or myVersion = 18  Then


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
			mySet_New_DB.fields("Site_Extranet_Code") = 0
    		mySet_New_DB.fields("Site_Presentation") = mySet_Old_DB.fields("Site_Presentation")
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
			mySet_New_DB.fields("Site_Date_Update") = myDate_Now()
			mySet_New_DB.fields("Site_Maximum_Files_Size") = mySet_Old_DB.fields("Site_Maximum_Files_Size")
			mySet_New_DB.fields("Site_Date_Format")       = mySet_Old_DB.fields("Site_Date_format")
			mySet_New_DB.fields("Site_Hour_Format")       = mySet_Old_DB.fields("Site_Hour_Format")
			mySet_New_DB.fields("Site_Language_ID") = mySet_Old_DB.fields("Site_language_ID")
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



if myVersion > 10 Then 
 Response.Redirect "__Administration_Site_SQL4.asp?action=import&version=" & myVersion
end if 



'Replacing Old Files
set myFile_System_Object=server.createobject("scripting.FileSystemObject")
myFile_System_Object.DeleteFile myDatabase_Path
myFile_System_Object.CopyFile  myNew_Database_Path , myDatabase_Path
myFile_System_Object.DeleteFile myNew_Database_Path

set myFile_System_Object=nothing


'CLOSE FRAME AND BACK TO ADMINISTRATION PAGE
%>
<script language="javascript">
<!--
top.location.href="__Administration_Site.asp"	
-->
</script>

<html><script language="JavaScript"></script></html>