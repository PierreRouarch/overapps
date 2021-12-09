<%
' ------------------------------------------------------------------------------------------------
' Copyright (C) 2001  + Ov-erA-pps - http://www.overapps.com
'
' This program "DB_Environment.asp" is free software; you can redistribute it 
' and/or modify
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
'-------------------------------------------------------------------------------------------------

%>

<%
' ------------------------------------------------------------------------------------------------
' Name : DB_Environment.asp
' Path : /_Include
' Version : 1.15.0
' Description : Read Environment in DB
' By : Pierre Rouarch
' Company : OverApps
' Update : October, 20 2001
'
' Contributors : Stéphane Chova 
' 
' 
' Modify by : 	
' Company : 
' Date : 
' 
' Last Modifications : 	
' 
' -------------------------------------------------------------------------------------------------


' DB Connection and SQL 
Dim myConnection, mySQL_Select_tb_Sites, mySet_tb_Sites,  mySQL_Select_tb_Sites_Members, mySet_tb_Sites_Members, mySQL_Select_tb_Styles, mySet_tb_Styles

Dim mySQL_Select_tb_Sites_Applications
Dim mySet_tb_Sites_Applications


' DB Site Variables
Dim myEngine_ID, myDistributor_ID, myNetWork_ID 

Dim mySite_ID, mySite_URL,  mySite_Name, mySite_Presentation , mySite_Public_Type_ID, mySite_Extranet_Code, mySite_Use_ID, mySite_Company, mySite_Address, mySite_Zip, mySite_City, mySite_State, mySite_Country_ID,  mySite_Phone, mySite_Fax, mySite_Web, mySite_Email

Dim mySite_Language_ID

Dim  mySite_Style_ID, mySite_Agenda_Open,  mySite_Projects_Open,  mySite_tasks_Open, mySite_Contacts_Open, mySite_Members_Open, mySite_Webs_Open, mySite_News_Open,  mySite_NewsLists_Open, mySite_Events_Open,  mySite_Files_Open,   mySite_NewsGroups_Open, mySite_Chatrooms_Open, mySite_Public_Directory, mySite_User_Personalization, mySite_Author_Update, mySite_Date_Update 

' DB Members Variables
Dim  myUser_ID, myUser_Login, myUser_Password, myUser_Title_ID, myUser_Name, myUser_Firstname, myUser_Pseudo, myUser_Email, myUser_Type_ID


' DB User Applications Variables
Dim myMax_User_Applications
Dim myNumber_User_Applications

Dim myUser_Application_ID()
Dim myUser_Application_Name()
Dim myUser_Application_Title()
Dim myUser_Application_Entry_Page()


Dim myUser_Application_Public_Type_ID()
Dim myUser_Application_Opened()

Dim myBox_title
Dim myApplication_Title
Dim myApplication_Public_Type_ID
Dim myMaximum_File_Size
Dim myHour_Format, myDate_Format
Dim myAgenda_Start, myAgenda_End, myAgenda_Week_Start



' Style Variables
DIM  myStyle_ID,  myStyle_Name, myGlobal_Width, myLeft_Width, myApplication_Width, myRight_Width,  myBGColor, myBGImage, myBGTextColor, myBorderColor,  myBorderImage, myBorderTextColor,  myApplicationColor, myApplicationImage, myApplicationTextColor,  myStyle_Author_Update, myStyle_Date_Update


Dim myBanner_Width, myBanner_Height

' Session's  Variables


mySite_ID = session("Site_ID") 
' Forced to site 1 
if len(mySite_ID)=0 then 
	mySite_ID=1
	session("Site_ID")=mySite_ID
end if
mySite_ID = CInt(mySite_ID)

' User ID
myUser_ID = session("User_ID")
if len(myUser_ID)=0 then
	myUser_ID = 0
else
	myUser_ID=CInt(myUser_ID)
end if

' USER Type ID
myUser_Type_ID = session("User_Type_ID")
if len(myUser_Type_ID)=0 then
	myUser_Type_ID = 7 ' Public
else
	myUser_Type_ID=CInt(myUser_Type_ID)
end if

' Standard' style (OverApps English)
myStyle_ID=1



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Read tb_Sites					 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Connection 
set myConnection = Server.CreateObject("ADODB.Connection")

'TEST IF SQL SERVER CONNECTION WORKS IF NOT , WE CONNECT TO ACCES DB

on error resume next
myConnection.Open myConnection_String



If myConnection.state <> 1 Then
 myConnection.Close
 set myConnection = Nothing
 Response.redirect("__Connection_Error.asp?error=1")
end if 
	



' Select 
mySQL_Select_tb_Sites = "SELECT * FROM tb_Sites WHERE Site_ID = "&mySite_ID
myConnection.Execute(mySQL_select_tb_Sites)
set mySet_tb_Sites = myConnection.Execute(mySQL_select_tb_Sites)

' if it's OK 
if not mySet_tb_Sites.eof then

	' For Future multi-sites and ASP (Applications Services Providers) purpose
	myEngine_ID = mySet_tb_Sites("Engine_ID")
	myDistributor_ID = mySet_tb_Sites("Distributor_ID")
	' For Web Services Purpose
	myNetwork_ID = mySet_tb_Sites("Network_ID")
	
	' Site information
	mySite_URL = mySet_tb_Sites("Site_URL")
	mySite_Name = mySet_tb_Sites("Site_Name")
	mySite_Presentation = mySet_tb_Sites("Site_Presentation")



	mySite_Public_Type_ID = mySet_tb_Sites("Site_Public_Type_ID")

	' For Extranet 
	mySite_Extranet_Code = mySet_tb_Sites("Site_Extranet_Code")

	' for Future Multi-Sites purpose
	mySite_Use_ID = mySet_tb_Sites("Site_Use_ID")


	' Language
	mySite_Language_ID=mySet_tb_Sites("Site_Language_ID")

	myCurrent_Language=mySite_Language_ID

	' Site Style
	mySite_Style_ID = mySet_tb_Sites("Site_Style_ID")
		
	' Company information
	mySite_Company = mySet_tb_Sites("Site_Company")
	mySite_Address = mySet_tb_Sites("Site_Address")	
	mySite_Zip = mySet_tb_Sites("Site_Zip")
	mySite_City = mySet_tb_Sites("Site_City")
	mySite_State = mySet_tb_Sites("Site_State")
	mySite_Country_ID = mySet_tb_Sites("Site_Country_ID")
	mySite_Phone = mySet_tb_Sites("Site_Phone")
	mySite_Fax = mySet_tb_Sites("Site_Fax")
	mySite_Web = mySet_tb_Sites("Site_Web")
	mySite_Email = mySet_tb_Sites("Site_Email")

	' For User Personalization Future Extension
	mySite_User_Personalization = mySet_tb_Sites("Site_User_Personalization")

	' For multi-sites extensions purpose
	mySite_Public_Directory= mySet_tb_Sites("Site_Public_Directory")

	' UpDate information
	mySite_Author_Update= mySet_tb_Sites("Site_Author_Update")
	mySite_Date_Update= mySet_tb_Sites("Site_Date_Update")
	
	'Files Information
	myMaximum_File_Size = mySet_tb_Sites("Site_Maximum_Files_Size")
	
	'Date Format
	myHour_Format = mySet_tb_Sites("Site_Hour_Format")
	myDate_Format = mySet_tb_Sites("Site_Date_Format")

	'Agenda Start&End
	myAgenda_Start = mySet_tb_Sites("Site_Agenda_Start")
	myAgenda_End  = mySet_tb_Sites("Site_Agenda_End")
	
	'Agenda Week Start
	myAgenda_Week_Start = mySet_tb_Sites("Site_Agenda_Week_Start")
	
	

else
	' Close Recordset and Connection 
	mySet_tb_Sites.close
	Set mySet_tb_Sites = Nothing
	myConnection.Close
	set myConnection = Nothing
	Response.redirect("__Connection_Error.asp?error=2")
end if


' Close Recordset and  Connection 
mySet_tb_Sites.close
Set mySet_tb_Sites = Nothing
myConnection.Close
set myConnection = Nothing

' Style 

if len(mySite_Style_ID)>0 then 
	myStyle_ID=mySite_Style_ID
end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Read tb_sites_members								 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''
if  myUser_ID<>0 and mySite_ID<>0 then 
	' Connection 
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String
	' SELECT
	mySQL_Select_tb_Sites_Members = "SELECT * FROM  tb_Sites_Members WHERE Site_ID="&mySite_ID&" AND Member_ID = "&myUser_ID

	set mySet_tb_Sites_Members = myConnection.Execute(mySQL_select_tb_Sites_Members)
	' if it's ok
	if not mySet_tb_Sites_Members.eof then
		myUser_Login = mySet_tb_Sites_Members("Member_Login")
		myUser_Title_ID = mySet_tb_sites_members("Member_Title_ID")	
		myUser_FirstName = mySet_tb_sites_members("Member_FirstName")
		myUser_Name = mySet_tb_sites_members("Member_Name")
		myUser_Pseudo = mySet_tb_sites_members("Member_Pseudo")		
		myUser_Email = mySet_tb_Sites_Members("Member_Email")
		myUser_Type_ID = mySet_tb_Sites_Members("Member_Type_ID")
	end if
' Close Recordset and connection
mySet_tb_Sites_Members.close
Set mySet_tb_Sites_Members = Nothing
myConnection.Close
set myConnection = Nothing

End if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Read tb_sites_applications to get user Applications opened.						 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if mySite_ID<>0 then 
	' Connection 
	set myConnection = Server.CreateObject("ADODB.Connection")
	myConnection.Open myConnection_String
	' SELECT COUNT
	mySQL_Select_tb_Sites_Applications = "SELECT COUNT (*) AS Max_User_Applications FROM tb_Sites_Applications WHERE Site_ID="&mySite_ID&" AND Site_Application_Opened=1 AND Site_Application_Public_type_ID >= "&myUser_type_ID
	set mySet_tb_Sites_Applications = myConnection.Execute(mySQL_select_tb_Sites_Applications)
	' if it's ok
	if not mySet_tb_Sites_Applications.eof then
		myMax_User_Applications=mySet_tb_Sites_Applications("Max_User_Applications")	
		ReDim myUser_Application_ID(myMax_User_Applications)
		ReDim myUser_Application_Name(myMax_User_Applications)
		ReDim myUser_Application_Title(myMax_User_Applications)
		ReDim myUser_Application_Public_Type_ID(myMax_User_Applications)
		ReDim myUser_Application_Opened(myMax_User_Applications)
		ReDim myUser_Application_Entry_Page(myMax_User_Applications)
	end if 
	' Close Recordset 
	mySet_tb_Sites_Applications.close
	Set mySet_tb_Sites_Applications = Nothing
	if myMax_User_Applications>=1 then

	' SELECT
	mySQL_Select_tb_Sites_Applications = "SELECT * FROM tb_Sites_Applications INNER JOIN tb_Applications ON tb_Applications.Application_ID=tb_Sites_Applications.Application_ID WHERE tb_Sites_Applications.Site_ID="&mySite_ID&" AND tb_Sites_Applications.Site_Application_Opened=1 AND tb_Sites_Applications.Site_Application_Public_type_ID >= "&myUser_type_ID&" Order by tb_Sites_Applications.Application_ID"
	


	set mySet_tb_Sites_Applications = myConnection.Execute(mySQL_select_tb_Sites_Applications)
		mySet_tb_Sites_Applications.movefirst
		myNumber_User_Applications=0
		do while not mySet_tb_Sites_Applications.eof and myNumber_User_Applications<=myMax_User_Applications
		
			myUser_Application_ID(myNumber_User_Applications)= mySet_tb_Sites_Applications("Application_ID")
			
			myUser_Application_Name(myNumber_User_Applications)= mySet_tb_Sites_Applications("Application_Name")
			myUser_Application_Title(myNumber_User_Applications)= mySet_tb_Sites_Applications("Site_Application_Title")
			myUser_Application_Entry_Page(myNumber_User_Applications)= mySet_tb_Sites_Applications("Application_Entry_Page")
			myUser_Application_Public_Type_ID(myNumber_User_Applications)= mySet_tb_Sites_Applications("Application_Public_Type_ID")
	
			myNumber_User_Applications=myNumber_User_Applications+1
			mySet_tb_Sites_Applications.movenext
		loop
	' Close Recordset
		mySet_tb_Sites_Applications.close
		Set mySet_tb_Sites_Applications = Nothing
	end if
' CLose Connection
myConnection.Close
set myConnection = Nothing

End if


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Read tb_Styles     					  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

' Select
mySQL_Select_tb_Styles = "SELECT * FROM tb_Styles WHERE Style_ID = "&myStyle_ID


set mySet_tb_Styles = myConnection.Execute(mySQL_select_tb_Styles)

' if it's  ok
if not mySet_tb_Styles.eof then
	myStyle_Name = mySet_tb_Styles("Style_Name")

	' Widthes
 	myGlobal_Width = mySet_tb_Styles("Style_Global_Width")
 	myLeft_Width = mySet_tb_Styles("Style_Left_Width")
	myApplication_Width = mySet_tb_Styles("Style_Application_Width")
	' Not Used
 	myRight_Width = mySet_tb_Styles("Style_Right_Width")

	' Colors, Images and Text Colors
 	myBGColor = mySet_tb_Styles("Style_BGColor")
 	myBGImage = mySet_tb_Styles("Style_BGImage")
 	myBGTextColor = mySet_tb_Styles("Style_BGTextColor")
	myBorderColor = mySet_tb_Styles("Style_BorderColor")
 	myBorderImage = mySet_tb_Styles("Style_BorderImage")
 	myBorderTextColor = mySet_tb_Styles("Style_BorderTextColor")
 	myApplicationColor = mySet_tb_Styles("Style_ApplicationColor")
 	myApplicationImage = mySet_tb_Styles("Style_ApplicationImage")
 	myApplicationTextColor = mySet_tb_Styles("Style_ApplicationTextColor")

else 
	 myStyle_ID=0
end if
' Close Recordset and Connection
mySet_tb_Styles.close
Set mySet_tb_Styles = Nothing
myConnection.Close
set myConnection = Nothing

'


''''''''''''''''''''''''''''''''''''''''''''''''''''
' IF STYLE NOT FOUND
''''''''''''''''''''''''''''''''''''''''''''''''''''

if myStyle_ID=0 then 

	' Widthes
	myGlobal_Width = 768
	myLeft_Width = 150
	myApplication_Width = 618
	' Not Used
	myRight_Width = 0

	' Colors, Images, Text Colors
	myBGColor = "#FFFFFF"
	myBGImage = ""
	myBGTextColor = "#000000"
	myBorderColor = "#cccc00"
	myBorderImage = ""
	myBorderTextColor = "#000000"
	myApplicationColor = "#663399"
	myApplicationImage = ""
	myApplicationTextColor = "#FFFFFF"


end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Other Style Constants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

myBanner_Width=468
myBanner_height=60

%>

<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Language or Vocabularry Selection system
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<!-- #include file="Global_Languages.asp" -->

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SOFTWARE VERSION 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim mySoftware_Version
 mySoftware_Version ="Version 1.18.0"
%>


<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>