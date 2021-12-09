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

<% 	Option Explicit 
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
myPage = "__Administration_Site.asp"

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

<%
Dim myConnection_SQL,myString_Temp,myString_temp2


Dim mySQL_Create(58)


'Declaration of all CREATE TABLE SQL QUERIES 
mySQL_Create(1) = "CREATE TABLE [tb_NewsWires_Members] (	[NewsWire_ID] [int] NULL ,	[Member_ID] [int] NULL,	[Site_ID] [int] NULL ,	[NewsWire_Member_Title] [nvarchar] (255) NULL ,	[NewsWire_Member_Name] [nvarchar](255) NULL ,	[NewsWire_Member_Firstname] [nvarchar] (255) NULL ,	[NewsWire_Member_Pseudo] [nvarchar](255) NULL ,	[NewsWire_Member_Email] [nvarchar] (255) NULL ,	[NewsWire_Member_Top] [int] NOT NULL)"

mySQL_Create(2) = "CREATE TABLE [tb_Applications] (	[Application_ID] [int] ,	[Application_Name][nvarchar] (255) NULL ,	[Application_Presentation] [nvarchar] (255) NULL ,[Application_Title][nvarchar] (255) NULL ,	[Application_Entry_Page] [nvarchar] (255) NULL ,[Application_Include_Box] [nvarchar] (255) NULL ,	[Application_Public_Type_ID] [int] NULL ,[Application_Opened][int] NULL ,	[Application_Author_Creation] [nvarchar] (255) NULL ,[Application_Date_Creation] [varchar](20) NULL ,	[Application_Author_Update] [nvarchar] (255) NULL ,[Application_Date_Update] [varchar](20) NULL)"

mySQL_Create(3) = "CREATE TABLE [tb_Calendars] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,[Calendar_ID][int] IDENTITY (1, 1) NOT NULL ,	[Calendar_Type_ID] [int] NULL ,	[Calendar_theme_ID] [int] NULL ,[Calendar_Name] [nvarchar] (255) NULL ,	[Calendar_Presentation] [nvarchar] (255) NULL ,	[Calendar_Public][int]NOT NULL ,	[Calendar_Author_Update] [nvarchar] (255) NULL ,	[Calendar_Date_Update] [varchar](20) NULL)"

mySQL_Create(4) = "CREATE TABLE [tb_Calendars_Members] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,[Calendar_ID] [int] NULL ,	[Calendar_Member_Title] [nvarchar] (255) NULL ,	[Calendar_Member_Name] [nvarchar](255)NULL ,	[Calendar_Member_Firstname] [nvarchar] (255) NULL ,	[Calendar_Member_Pseudo] [nvarchar](255) NULL ,	[Calendar_Member_Email] [nvarchar] (255) NULL ,	[Calendar_Member_Top] [int] NOT NULL)"

mySQL_Create(5) = "CREATE TABLE [tb_Calendars_Sites] (	[Calendar_ID] [int] NULL ,	[Site_ID] [int] NULL ,[Calendar_Site_Top] [int] NOT NULL)"

mySQL_Create(6) = "CREATE TABLE [tb_Calendars_Themes] (	[Calendar_theme_ID] [int] IDENTITY (1, 1) NOT NULL ,[Calendar_theme_Path] [nvarchar] (255) NULL ,	[Calendar_theme_Name] [nvarchar] (255) NULL ,[Calendar_theme_Presentation] [nvarchar] (255) NULL)"

mySQL_Create(7) = "CREATE TABLE [tb_Calendars_Types] (	[Calendar_type_ID] [int] IDENTITY (1, 1) NOT NULL ,[Calendar_type_Name] [nvarchar] (255) NULL ,	[Calendar_type_Share] [int] NOT NULL)"

mySQL_Create(8) = "CREATE TABLE [tb_contacts] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Directory_ID][int] NULL ,	[Contact_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Contact_Title_ID] [int] NULL ,[Contact_Name] [nvarchar] (255) NULL ,	[Contact_Firstname] [nvarchar] (255) NULL ,	[Contact_Company_Type] [int] NULL ,	[Contact_Company] [nvarchar] (255) NULL ,	[Contact_Company_Activity_ID] [int] NULL ,[Contact_Company_Address] [nvarchar] (255) NULL ,	[Contact_Company_Zip] [nvarchar] (255) NULL ,[Contact_Company_City] [nvarchar] (255) NULL ,	[Contact_Company_State] [nvarchar] (255) NULL ,[Contact_Company_Country_ID] [int] NULL ,	[Contact_Company_Phone] [nvarchar] (255) NULL ,[Contact_Company_Mobile] [nvarchar] (255) NULL ,	[Contact_Company_Fax] [nvarchar] (255) NULL ,[Contact_Company_Email] [nvarchar] (255) NULL ,	[Contact_Company_Web] [nvarchar] (255) NULL ,[Contact_Company_Fonction] [nvarchar] (255) NULL ,	[Contact_Home_Type] [int]  NULL ,	[Contact_Home_Address][nvarchar] (255) NULL ,	[Contact_Home_Zip] [nvarchar] (255) NULL ,	[Contact_Home_City][nvarchar] (255) NULL ,	[Contact_Home_State] [nvarchar] (255) NULL ,[Contact_Home_Country_ID][int] NULL ,	[Contact_Home_Phone] [nvarchar] (255) NULL ,	[Contact_Home_Mobile][nvarchar] (255) NULL ,	[Contact_Home_Fax] [nvarchar] (255) NULL ,	[Contact_Home_Email][nvarchar] (255) NULL ,	[Contact_Home_Web] [nvarchar] (255) NULL ,	[Contact_Comments] [ntext]NULL ,	[Contact_Author_Update] [nvarchar] (255) NULL ,	[Contact_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(9) = "CREATE TABLE [tb_Contacts_Activities] (	[Directory_ID] [int] NULL ,	[Contact_Activity_ID][int]IDENTITY (1, 1) NOT NULL ,	[Contact_Activity_Path] [nvarchar] (255) NULL ,[Contact_Activity_NAF700] [nvarchar] (255) NULL ,	[Contact_Activity_Name] [nvarchar] (255) NULL ,[Contact_Activity_Presentation] [nvarchar] (255) NULL)" 

mySQL_Create(10)= "CREATE TABLE [tb_Countries] (	[Country_ID] [int] NULL ,	[Country] [nvarchar] (255) NULL ,	[Country_Domain] [nvarchar] (255) NULL)"

mySQL_Create(11)= "CREATE TABLE [tb_Directories] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Directory_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Directory_Type_ID] [int] NULL ,	[Directory_theme_ID] [int] NULL ,	[Directory_Name] [nvarchar] (255) NULL ,	[Directory_Presentation] [nvarchar] (255) NULL ,	[Directory_Public] [int] NOT NULL ,	[Directory_Author_Update] [nvarchar] (255) NULL ,	[Directory_Date_Update] [varchar](20) NULL)"

mySQL_Create(12)= "CREATE TABLE [tb_Directories_Members] (	[Directory_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Directory_Member_Title] [nvarchar] (255) NULL ,	[Directory_Member_Name] [nvarchar] (255) NULL ,	[Directory_Member_Firstname] [nvarchar] (255) NULL ,	[Directory_Member_Pseudo] [nvarchar] (255) NULL ,	[Directory_Member_Email] [nvarchar] (255) NULL ,	[Directory_Member_Top] [int] NOT NULL)"

mySQL_Create(13)= "CREATE TABLE [tb_Directories_Sites] (	[Directory_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Directory_Site_Top] [int] NOT NULL)"

mySQL_Create(14)= "CREATE TABLE [tb_Directories_themes] (	[Directory_theme_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Directory_theme_Path] [nvarchar] (255) NULL ,	[Directory_theme_Name] [nvarchar] (255) NULL ,	[Directory_theme_Presentation] [nvarchar] (255) NULL)"

mySQL_Create(15)= "CREATE TABLE [tb_Directories_Types] (	[Directory_type_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Directory_type_Name] [nvarchar] (255) NULL ,	[Directory_type_Share] [int] NOT NULL)"

mySQL_Create(16)= "CREATE TABLE [tb_Events] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Calendar_ID] [int] NULL ,	[Event_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Event_Name] [nvarchar] (255) NULL ,	[Event_Presentation] [nvarchar] (255) NULL ,	[Event_Date_Beginning] [nvarchar] (50) NULL ,	[Event_Date_End] [nvarchar] (50) NULL ,	[Event_Author_Update] [nvarchar] (255) NULL ,	[Event_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(17)= "CREATE TABLE [tb_Files] (	[File_ID] [int] IDENTITY (1, 1) NOT NULL ,	[File_Creator_ID] [int] NULL ,	[File_Responsible_ID] [int] NULL ,	[File_Folder_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[File_Modificator_ID] [int] NULL ,	[File_Name] [nvarchar] (255) NULL ,	[File_Short_Description] [nvarchar] (255) NULL ,	[File_Long_Description] [ntext] NULL ,	[File_Type] [nvarchar] (255) NULL ,	[File_Size] [int] NULL ,	[File_Modification_Date] [nvarchar] (50) NULL)"

mySQL_Create(18)= "CREATE TABLE [tb_Files_Extensions] (	[File_Extension_ID] [int] IDENTITY (1, 1) NOT NULL ,	[File_Extension] [nvarchar] (50) NULL ,	[File_Extension_Autorised] [int] NULL ,	[Site_ID] [int] NULL)"

mySQL_Create(19)= "CREATE TABLE [tb_Folders] (	[Folder_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Folder_Creator_ID] [int] NULL ,	[Folder_Responsible_ID] [int] NULL ,	[Folder_Modificator_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Folder_Name] [nvarchar] (50) NULL ,	[Folder_Short_Description] [nvarchar] (50) NULL ,	[Folder_Long_Description] [ntext] NULL ,	[Folder_Public] [int] NULL ,	[Folder_Modification_Date] [nvarchar] (50) NULL)"

mySQL_Create(20)= "CREATE TABLE [tb_Folders_Access] (	[Folder_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Site_ID] [int] NULL)"

mySQL_Create(21)= "CREATE TABLE [tb_Meetings] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Project_ID] [int] NULL ,	[Phase_ID] [int] NULL ,	[Meeting_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Meeting_Title] [nvarchar] (255) NULL ,	[Meeting_Date_Beginning] [nvarchar] (50) NULL ,	[Meeting_Hour] [int] NULL ,	[Meeting_Minute] [int] NULL ,	[Meeting_Length] [int] NULL ,	[Meeting_Length_In_Minutes] [int] NULL ,	[Meeting_Place] [nvarchar] (255) NULL ,	[Meeting_Agenda] [ntext] NULL ,	[Meeting_Comments] [ntext] NULL ,	[Meeting_Author_Update] [nvarchar] (255) NULL ,	[Meeting_Date_Update] [nvarchar] (50) NULL ,	[Meeting_Public] [int] NULL)"

mySQL_Create(21)= "CREATE TABLE [tb_Meetings_Members] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Meeting_ID] [int] NULL ,	[Meeting_Role_ID] [int] NULL)"

mySQL_Create(22)= "CREATE TABLE [tb_News] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[NewsWire_ID] [int] NULL ,	[New_ID] [int] IDENTITY (1, 1) NOT NULL ,	[New_Category_ID] [int] NULL ,	[New_Date] [nvarchar] (50) NULL ,	[New_Title] [nvarchar] (255) NULL ,	[New_Description_Short] [nvarchar] (255) NULL ,	[New_Description_Long] [ntext] NULL ,	[New_Top] [int] NULL ,	[New_Public] [int] NULL ,	[New_Author_Update] [nvarchar] (255) NULL ,	[New_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(23)= "CREATE TABLE [tb_NewsGroups] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[NewsGroup_theme_ID] [int] NULL ,	[NewsGroup_ID] [int] IDENTITY (1, 1) NOT NULL ,	[NewsGroup_Name] [nvarchar] (255) NULL ,	[NewsGroup_Presentation] [nvarchar] (255) NULL ,	[NewsGroup_Public] [int] NOT NULL ,	[NewsGroup_Author_Update] [nvarchar] (255) NULL ,	[NewsGroup_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(24)= "CREATE TABLE [tb_NewsGroups_messages] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[NewsGroup_ID] [int] NULL ,	[NewsGroup_Message_ID] [int] IDENTITY (1, 1) NOT NULL ,	[NewsGroup_Message_Date] [nvarchar] (50) NULL ,	[NewsGroup_Message_Author] [nvarchar] (255) NULL ,	[NewsGroup_Message_title] [nvarchar] (255) NULL ,	[NewsGroup_Message] [ntext] NULL ,	[NewsGroup_Message_Thread] [nvarchar] (255) NULL ,	[NewsGroup_Message_Thread_Date] [nvarchar] (50) NULL)"

mySQL_Create(25)= "CREATE TABLE [tb_NewsGroups_Sites] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[NewsGroup_ID] [int] NULL ,	[NewsGroup_Site_Top] [int]  NULL ,	[NewsGroupSite_Author_UpDate] [nvarchar] (255) NULL ,	[NewsGroup_Site_Date_Update] [varchar](20) NULL)"

mySQL_Create(26)= "CREATE TABLE [tb_NewsGroups_themes] (	[NewsGroup_theme_ID] [int] IDENTITY (1, 1) NOT NULL ,	[NewsGroup_theme_Path] [nvarchar] (255) NULL ,	[NewsGroup_theme_Name] [nvarchar] (255) NULL ,	[NewsGroup_theme_Presentation] [nvarchar] (255) NULL)"

mySQL_Create(27)= "CREATE TABLE [tb_NewsWires] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[NewsWire_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[NewsWire_Type_ID] [int] NULL ,	[NewsWire_theme_ID] [int] NULL ,	[NewsWire_Name] [nvarchar] (255) NULL ,	[NewsWire_Presentation] [nvarchar] (255) NULL ,	[NewsWire_Public] [int]  NULL ,	[NewsWire_Author_Update] [nvarchar] (255) NULL ,	[NewsWire_Date_Update] [varchar](20) NULL)"

mySQL_Create(28)= "CREATE TABLE [tb_NewsWires_Sites] (	[NewsWire_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[NewsWire_Site_Top] [int]  NULL)"

mySQL_Create(29)= "CREATE TABLE [tb_NewsWires_Themes] (	[NewsWire_theme_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[NewsWire_theme_Path] [nvarchar] (255) NULL ,	[NewsWire_theme_Name] [nvarchar] (255) NULL ,	[NewsWire_theme_Presentation] [nvarchar] (255) NULL)"

mySQL_Create(30)= "CREATE TABLE [tb_NewsWires_Types] (	[NewsWire_type_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[NewsWire_type_Name] [nvarchar] (255) NULL ,	[NewsWire_type_Share] [int]  NULL)"

mySQL_Create(31)= "CREATE TABLE [tb_phases] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Project_ID] [int] NULL ,	[Phase_Parent_ID] [int] NULL ,	[Phase_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Phase_Path] [nvarchar] (255) NULL ,	[Phase_Name] [nvarchar] (255) NULL ,	[Phase_Presentation] [nvarchar] (255) NULL ,	[Phase_Date_Beginning] [nvarchar] (50) NULL ,	[Phase_Date_End] [nvarchar] (50) NULL ,	[Phase_Date_Beginning2] [nvarchar] (50) NULL ,	[Phase_Date_End2] [nvarchar] (50) NULL ,	[Phase_Status_ID] [int] NULL ,	[Phase_Leader_ID] [int] NULL ,	[Phase_Priority_ID] [int] NULL ,	[Phase_Progress] [int] NULL ,	[Phase_Personnal] [int]  NULL ,	[Phase_Author_Update] [nvarchar] (255) NULL ,	[Phase_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(32)= "CREATE TABLE [tb_Phases_Members] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Phase_ID] [int] NULL)"

mySQL_Create(33)= "CREATE TABLE [tb_Phases_Priorities] (	[Engine_ID] [int] NULL ,	[Distributor_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Phase_Priority_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Phase_Priority_Name] [nvarchar] (255) NULL)"

mySQL_Create(34)= "CREATE TABLE [tb_phases_status] (	[Engine_ID] [int] NULL ,	[Distributor_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Phase_Status_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Phase_Status_Name] [nvarchar] (255) NULL)"

mySQL_Create(35)= "CREATE TABLE [tb_projects] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Project_Parent_ID] [int] NULL ,	[Project_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Project_Path] [nvarchar] (255) NULL ,	[Project_Name] [nvarchar] (255) NULL ,	[Project_Presentation] [nvarchar] (255) NULL ,	[Project_Date_Beginning] [nvarchar] (50) NULL ,	[Project_Date_End] [nvarchar] (50) NULL ,	[Project_Date_Beginning2] [nvarchar] (50) NULL ,	[Project_Date_End2] [nvarchar] (50) NULL ,	[Project_Status_ID] [int] NULL ,	[Project_Leader_ID] [int] NULL ,	[Project_Priority_ID] [int] NULL ,	[Project_Progress] [int] NULL ,	[Project_Personnal] [int]  NULL ,	[Project_Author_Update] [nvarchar] (255) NULL ,	[Project_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(36)= "CREATE TABLE [tb_Projects_Members] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Project_ID] [int] NULL)"

mySQL_Create(37)= "CREATE TABLE [tb_Projects_Priorities] (	[Engine_ID] [int] NULL ,	[Distributor_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Project_Priority_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Project_Priority_Name] [nvarchar] (255) NULL)"

mySQL_Create(38)= "CREATE TABLE [tb_Projects_Status] (	[Engine_ID] [int] NULL ,	[Distributor_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Project_Status_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Project_Status_Name] [nvarchar] (255) NULL)"

mySQL_Create(39)= "CREATE TABLE [tb_Sites] (	[Engine_ID] [int] NULL , [Distributor_ID] [int] NULL ,	[NetWork_ID] [int] NULL ,	[Site_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Site_URL] [nvarchar] (255) NULL ,	[Site_Name] [nvarchar] (255) NULL ,	[Site_Presentation] [nvarchar] (255) NULL ,	[Site_Public_Type_ID] [int] NULL ,	[Site_Extranet_Code] [nvarchar] (255) NULL ,	[Site_Use_ID] [int] NULL ,	[Site_Company] [nvarchar] (255) NULL ,	[Site_Address] [nvarchar] (255) NULL ,	[Site_zip] [nvarchar] (255) NULL ,	[Site_City] [nvarchar] (255) NULL ,	[Site_State] [nvarchar] (255) NULL ,	[Site_Country_ID] [int] NULL ,	[Site_Phone] [nvarchar] (255) NULL ,	[Site_Fax] [nvarchar] (255) NULL ,	[Site_Web] [nvarchar] (255) NULL ,	[Site_Email] [nvarchar] (255) NULL ,[Site_Language_ID] [int] NULL ,	[Site_Style_ID] [int] NULL ,	[Site_User_Personalization] [int] NULL ,[Site_Public_Directory] [int]  NULL ,	[Site_Author_Update] [nvarchar] (255) NULL ,	[Site_Date_Update][varchar](20) NULL ,	[Site_Maximum_Files_Size] [int] NULL ,	[Site_Date_Format] [int] NULL,[Site_Hour_Format] [int] NULL ,	[Site_Agenda_Start] [int] NULL ,[Site_Agenda_End] [int] NULL ,[Site_Agenda_Week_Start] [int] NULL)"

mySQL_Create(40)= "CREATE TABLE [tb_Sites_Activities] (	[Engine_ID] [int] NULL ,	[Site_Activity_ID][int] IDENTITY (1, 1)  NOT NULL ,	[Site_Activity_Path] [nvarchar] (255) NULL ,[Site_Activity_NAF700] [nvarchar] (255) NULL ,	[Site_Activity_Name] [nvarchar] (255) NULL ,[Site_Activity_Presentation] [nvarchar] (255) NULL)"

mySQL_Create(41)= "CREATE TABLE [tb_Sites_Applications] (	[Site_ID] [int] NULL ,	[Application_ID] [int] NULL ,	[Site_Application_Title] [nvarchar] (255) NULL ,	[Site_Application_Public_Type_ID] [int] NULL ,	[Site_Application_Opened] [int] NULL ,	[Site_Application_Author_Update] [nvarchar] (255) NULL ,	[Site_Application_Date_Update] [varchar](20) NULL)"

mySQL_Create(42)= "CREATE TABLE [tb_Sites_members] (	[Site_ID] [int] NULL ,	[Member_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Member_Login] [nvarchar] (255) NULL ,	[Member_Password] [nvarchar] (255) NULL ,	[Member_Title_ID] [int] NULL ,	[Member_Name] [nvarchar] (255) NULL ,	[Member_Firstname] [nvarchar] (255) NULL ,	[Member_Pseudo] [nvarchar] (255) NULL ,	[Member_Email] [nvarchar] (255) NULL ,	[Member_Company_Type] [int]  NULL ,	[Member_Company] [nvarchar] (255) NULL ,	[Member_Company_Activity_ID] [int] NULL ,	[Member_Company_Address] [nvarchar] (255) NULL ,	[Member_Company_Zip] [nvarchar] (255) NULL ,	[Member_Company_City] [nvarchar] (255) NULL ,	[Member_Company_State] [nvarchar] (255) NULL ,	[Member_Company_Country_ID] [int] NULL ,	[Member_Company_Phone] [nvarchar] (255) NULL ,	[Member_Company_Mobile] [nvarchar] (255) NULL ,	[Member_Company_Fax] [nvarchar] (255) NULL ,	[Member_Company_Email] [nvarchar] (255) NULL ,	[Member_Company_Web] [nvarchar] (255) NULL ,	[Member_Company_Fonction] [nvarchar] (255) NULL ,	[Member_Home_Type] [int]  NULL ,	[Member_Home_Address] [nvarchar] (255) NULL ,	[Member_Home_Zip] [nvarchar] (255) NULL ,	[Member_Home_City] [nvarchar] (255) NULL ,	[Member_Home_State] [nvarchar] (255) NULL ,	[Member_Home_Country_ID] [int] NULL ,	[Member_Home_Phone] [nvarchar] (255) NULL ,	[Member_Home_Mobile] [nvarchar] (255) NULL ,	[Member_Home_Fax] [nvarchar] (255) NULL ,	[Member_Home_Email] [nvarchar] (255) NULL ,	[Member_Home_Web] [nvarchar] (255) NULL ,	[Member_Comments] [ntext] NULL ,	[Member_Type_ID] [int] NULL ,	[Member_Public_Directory] [int]  NULL ,	[Member_Mailing_List] [int]  NULL ,	[Member_Promotion_List] [int]  NULL ,	[Member_Author_Update] [nvarchar] (255) NULL ,	[Member_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(43)= "CREATE TABLE [tb_Sites_Members_Activities] (	[Site_ID] [int] NULL ,	[Member_Activity_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Member_Activity_Path] [nvarchar] (255) NULL ,	[Member_Activity_NAF700] [nvarchar] (255) NULL ,	[Member_Activity_Name] [nvarchar] (255) NULL ,	[Member_Activity_Presentation] [nvarchar] (255) NULL)"

mySQL_Create(44)= "CREATE TABLE [tb_Styles] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Style_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Style_Name] [nvarchar] (255) NULL ,	[Style_Global_Width] [int] NULL ,	[Style_Left_Width] [int] NULL ,	[Style_Application_Width] [int] NULL ,	[Style_Right_Width] [int] NULL ,	[Style_BGColor] [nvarchar] (255) NULL ,	[Style_BGImage] [nvarchar] (255) NULL , [Style_BGTextColor] [nvarchar] (255) NULL ,	[Style_BorderColor] [nvarchar] (255) NULL ,	[Style_BorderImage] [nvarchar] (255) NULL ,	[Style_BorderTextColor] [nvarchar] (255) NULL ,	[Style_ApplicationColor] [nvarchar] (255) NULL ,	[Style_ApplicationImage] [nvarchar] (255) NULL ,	[Style_ApplicationTextColor] [nvarchar] (255) NULL ,	[Style_Author_Update] [nvarchar] (255) NULL ,	[Style_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(45)= "CREATE TABLE [tb_Styles_Color] (	[Color_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Color] [nvarchar] (255) NULL ,	[Color_Code] [nvarchar] (50) NULL)"

mySQL_Create(46)= "CREATE TABLE [tb_tasks] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Project_ID] [int] NULL ,	[Phase_ID] [int] NULL ,	[Task_Parent_ID] [int] NULL ,	[Task_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Task_Path] [nvarchar] (255) NULL ,	[Task_Name] [nvarchar] (255) NULL ,	[Task_Presentation] [nvarchar] (255) NULL ,	[Task_date_Beginning][varchar](20) NULL ,	[Task_Date_End] [varchar](20) NULL ,	[Task_Date_Beginning2] [varchar](20) NULL ,	[Task_Date_End2] [varchar](20) NULL ,	[Task_Status_ID] [int] NULL ,	[Task_Leader_ID] [int] NULL ,[Task_Priority_ID] [int] NULL ,	[Task_Progress] [int] NULL ,	[Task_Personnal] [int]  NULL ,	[Task_Author_Update] [nvarchar] (255) NULL ,	[Task_Date_Update] [varchar](20) NULL)"

mySQL_Create(47)= "CREATE TABLE [tb_Tasks_Members] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Task_ID] [int] NULL)"

mySQL_Create(48)= "CREATE TABLE [tb_Tasks_Priorities] (	[Engine_ID] [int] NULL ,	[Distributor_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Task_Priority_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Task_Priority_Name] [nvarchar] (255) NULL)"

mySQL_Create(49)= "CREATE TABLE [tb_tasks_Status] (	[Engine_ID] [int] NULL ,	[Distributor_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Task_Status_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Task_Status_Name] [nvarchar] (255) NULL)"

mySQL_Create(50)= "CREATE TABLE [tb_WebDirectories] (	[WebDirectory_ID] [int] IDENTITY (1, 1) NOT NULL ,	[WebDirectory_Type_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[WebDirectory_theme_ID] [int] NULL ,	[WebDirectory_Name] [varchar] (255) NULL ,	[WebDirectory_Presentation] [varchar] (255) NULL ,	[WebDirectory_Public] [int]  NULL ,	[WebDirectrory_Author_Update] [nvarchar] (255) NULL ,	[WebDirectrory_Date_Update] [varchar](20) NULL)"

mySQL_Create(51)= "CREATE TABLE [tb_WebDirectories_Members] (	[WebDirectory_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[WebDirectory_Member_Title] [nvarchar] (255) NULL ,	[WebDirectory_Member_Name] [nvarchar] (255) NULL ,	[WebDirectory_Member_Firstname] [nvarchar] (255) NULL ,	[WebDirectory_Member_Pseudo] [nvarchar] (255) NULL ,	[WebDirectory_Member_Email] [nvarchar] (255) NULL ,	[WebDirectory_Member_Top] [int]  NULL)"

mySQL_Create(52)= "CREATE TABLE [tb_WebDirectories_Sites] (	[WebDirectory_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[WebDirectory_Site_Top] [int]  NULL)"

mySQL_Create(53)= "CREATE TABLE [tb_WebDirectories_themes] (	[WebDirectory_theme_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[WebDirectory_theme_Path] [nvarchar] (255) NULL ,	[WebDirectory_theme_Name] [nvarchar] (255) NULL ,	[WebDirectory_theme_Presentation] [nvarchar] (255) NULL)"

mySQL_Create(54)= "CREATE TABLE [tb_WebDirectories_types] (	[WebDirectory_type_ID] [int] IDENTITY (1, 1) NOT NULL ,	[WebDirectory_type_Name] [nvarchar] (255) NULL ,	[WebDirectory_type_Share] [int]  NULL)"

mySQL_Create(55)= "CREATE TABLE [tb_Webs] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[WebDirectory_ID] [int] NULL ,	[Category_ID] [int] NULL ,	[Web_ID] [int] IDENTITY (1, 1) NOT NULL ,	[Web_URL] [nvarchar] (50) NULL ,	[Web_Name] [nvarchar] (50) NULL ,	[Web_Description_Short] [nvarchar] (50) NULL ,	[Web_Description_Long] [ntext] NULL ,	[Web_Top] [int]  NULL ,	[Web_Public] [int]  NULL ,	[Web_Author_Update] [nvarchar] (50) NULL ,	[Web_Date_Update] [nvarchar] (50) NULL)"

mySQL_Create(56)= "CREATE TABLE [tb_Webs_Categories] (	[Category_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Category_Parent_ID] [int] NULL ,	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Category_URL] [nvarchar] (255) NULL ,	[Category_Name] [nvarchar] (255) NULL ,	[Category_Top] [int]  NULL)"

mySQL_Create(57)= "CREATE TABLE [tb_Meetings] (	[Site_ID] [int] NULL ,	[Member_ID] [int] NULL ,	[Project_ID] [int] NULL ,	[Phase_ID] [int] NULL ,	[Meeting_ID] [int] IDENTITY (1, 1)  NOT NULL ,	[Meeting_Title] [nvarchar] (255) NULL ,	[Meeting_Date_Beginning] [nvarchar] (50) NULL , [Meeting_Hour] [int] NULL ,	[Meeting_Minute] [int] NULL ,[Meeting_Length] [int] NULL ,	[Meeting_Length_In_Minutes] [int] NULL ,	[Meeting_Place] [nvarchar] (255) NULL ,	[Meeting_Agenda] [ntext] NULL ,	[Meeting_Comments] [ntext] NULL ,	[Meeting_Author_Update] [nvarchar] (255) NULL ,	[Meeting_Date_Update] [nvarchar] (50) NULL ,	[Meeting_Public] [int] NULL)" 
%>

<%
Dim i,myTable_Name
i = 1
'NEEDED FOR SQL ERROR WHEN DELETING TABLES THAT MAY NOT EXISTS
on error resume next



'Connect to DB

set myConnection_SQL = Server.CreateObject("ADODB.Connection")
myConnection_SQL.Open myConnection_String_SQL


'DELETE ALL OLD TABLES BEFORE CREATING NEWS
i = 1
do while i < ubound(mySQL_Create)
'GET THE NAME OF THE TABLE
 myString_Temp = split(mySQL_Create(i),"]")
 myString_Temp2 = split(myString_Temp(0),"[")
 myTable_Name = myString_Temp2(1)
 myConnection_SQL.Execute("DROP TABLE "  & myTable_Name )
 i = i + 1
loop 

'CREATION OF TABLES
i = 1
do while i < ubound(mySQL_Create)
 if len(mySQL_Create(i)) > 0 Then myConnection_SQL.Execute(mySQL_Create(i))
 i = i + 1
 
loop

myConnection_SQL.Close()
set myConnection_SQL = NOTHING
response.redirect "__Administration_Site_SQL4.asp"
%>

<%''''''''''''''''''''''''''''''''''END OF SCRIPT''''''''''''''''''''%>
</TD></TR>
<%
'CENTER APPLICATION
%>
</TABLE>
<%
'CENTER
%>
<%
'DOWN
%> 
<!-- #include file="_borders/Down.asp" --> 
<% 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do not remove this Copyright Notice if you want to stay under this programm	'
' license's compliances.		                             					'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright			                                	'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 
</BODY>
</HTML>

<html><script language="JavaScript"></script></html>