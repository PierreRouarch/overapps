<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001  Pierre Rouarch
' This program "Global_Languages.asp" is free software; you can redistribute it 
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
' 	" Copyright (C) 2001 OverApps and contributors"
' at the bottom of the page with an active link from the name "OverApps" to 
' the Address http://www.overapps.com where the user (netsurfer) could find the 
' appropriate information about this license. 
'
'-----------------------------------------------------------------------------
%>

<%
' -------------------------------------------------------------------------------------
' Name 			:	Global_Languages.asp
' Path 			: /_Include
' Version 		: 1.18.0
' Description 	: General Languages File configuration
' By			: Pierre Rouarch
' Company 		: 
' Update		: April, 11 2002
' Comments 		: Italian Language
'
' Contributions : Nicolas Sanchez, Stéphane Chova, Alfredo Salvador, Dania Tcherkezoff
' 
' --------------------------------------------------------------------------------------
%>

<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Messages 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim myMessage_A_Professionnal
Dim myMessage_Activity
Dim myMessage_Add


' ADDRESS
Dim myMessage_Address
	Dim myMessage_City
	Dim myMessage_Zip_Code
	Dim myMessage_State
	Dim myMessage_Country
	Dim myMessage_Phone
	Dim myMessage_Mobile
	Dim myMessage_Fax
	Dim myMessage_Email
	Dim myMessage_Web


Dim myMessage_Administration
Dim myMessage_All
Dim myMessage_An_Individual


' APPLICATIONS
Dim myMessage_Application
Dim myMessage_Applications

	Dim myMessage_Agenda
	Dim myMessage_Projects
	Dim myMessage_Tasks
	Dim myMessage_Members
	Dim myMessage_Contacts
	Dim myMessage_Webs
	Dim myMessage_News
	Dim myMessage_Events
	Dim myMessage_NewsGroups


Dim myMessage_Applications_Administration
Dim myMessage_Article
Dim myMessage_At
Dim myMessage_Author
Dim myMessage_Beginning
Dim myMessage_Blank_By_The_Moderator
Dim myMessage_Both
Dim myMessage_Catch_Line
Dim myMessage_Check_if_you_are_a_professional
Dim myMessage_Check_if_you_are_a_particular
Dim myMessage_Clean
Dim myMessage_Comments
Dim myMessage_Company
Dim myMessage_Confirmation
Dim myMessage_Contact
Dim myMessage_Correction
Dim myMessage_Create_Your_Account
Dim myMessage_Date



' DAYS
Dim myMessage_Day
	Dim myMessage_Sunday
	Dim myMessage_Monday
	Dim myMessage_Tuesday
	Dim myMessage_Wednesday
	Dim myMessage_Thursday
	Dim myMessage_Friday
	Dim myMessage_Saturday
' /DAYS

Dim myMessage_Delete
Dim myMessage_During
Dim myMessage_Directory
Dim myMessage_End
Dim myMessage_Event
Dim myMessage_Fonction
Dim myMessage_General_Parameters
Dim myMessage_Go
Dim myMessage_Cancel
Dim myMessage_Hello
Dim myMessage_Hour
Dim myMessage_Home
Dim myMessage_Identification

' IDENTITY
Dim myMessage_Identity

	' IDENTITY TITLE
	Dim myMessage_Title
		Dim myMessage_Mister
		Dim myMessage_Miss
		Dim myMessage_Misses
		' For a religious Community		
		Dim myMessage_Father 
		Dim myMessage_Brother 
		Dim myMessage_Sister 
		Dim myMessage_Mother
	' / IDENTITY TITLE

	Dim myMessage_FirstName
	Dim myMessage_Name
	Dim myMessage_Pseudo

	' / IDENTITY 

Dim myMessage_Information
Dim myMessage_Inscription


Dim myMessage_Language
Dim myMessage_Leader
Dim myMessage_Length
Dim myMessage_Login
Dim myMessage_Meeting_Agenda
Dim myMessage_Member

' Member Type 
Dim myMessage_Member_Type
	Dim myMessage_Administrator	
	Dim myMessage_Moderator
	Dim myMessage_Intranet_Member
	Dim myMessage_Extranet_Member
	Dim myMessage_Web_Member
	Dim myMessage_Identified_Email
	


Dim myMessage_Members_Administration
Dim myMessage_Message
Dim myMessage_Min

Dim myMessage_Modification

' MONTHS
Dim myMessage_Month
	Dim myMessage_January
	Dim myMessage_February
	Dim myMessage_March
	Dim myMessage_April
	Dim myMessage_May
	Dim myMessage_June
	Dim myMessage_July
	Dim myMessage_August
	Dim myMessage_September
	Dim myMessage_October
	Dim myMessage_November
	Dim myMessage_December

Dim myMessage_Modify
Dim myMessage_More

Dim myMessage_New
Dim myMessage_No_Event
Dim myMessage_No_News

Dim myMessage_No_Project
Dim myMessage_No_Task

Dim myMessage_Office
Dim myMessage_Opened
Dim myMessage_Other_Web_Site

Dim myMessage_Page
Dim myMessage_Parent
Dim myMessage_Participants
Dim myMessage_Participate
Dim myMessage_Password
Dim myMessage_Phase
Dim myMessage_Place
Dim myMessage_Presentation
Dim myMessage_Priority
Dim myMessage_Project
Dim myMessage_Progress
Dim myMessage_Public

Dim myMessage_Required
Dim myMessage_Response
Dim myMessage_Revised

Dim myMessage_Search
Dim myMessage_Scheduled
Dim myMessage_Select
Dim myMessage_Site
Dim myMessage_Site_Public
Dim myMessage_Status

Dim myMessage_Style


Dim myMessage_To_Come
Dim myMessage_Today
Dim myMessage_URL_Address
Dim myMessage_Week
Dim myMessage_You_Search_For

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Error_Messages 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim myError_Message_Date_End_Before_Date_Beginning
Dim myError_Message_Enter_Date_Beginning
Dim myError_Message_Enter_Date_End
Dim myError_Message_Invalid_Date
Dim myError_Message_Login
Dim myError_Message_Login_Password
Dim myError_Message_Login_Already_used
Dim myError_Message_Password
Dim myError_Message_Password_Confirmation
Dim myError_Message_Password_and_confirmation_Do_not_match

' Form_Validation.asp error Messages
Dim myError_Message_Not_a_valid_name
Dim myError_Message_Not_a_valid_Address
Dim myError_Message_not_a_valid_Date
Dim  myError_Message_Required
Dim myError_Message_Not_Valid_alpha ' Not valid (caracters and/or blanks)
Dim myError_Message_Not_Valid_alpha_without_blank 'Not Valid (caracters no blank)
Dim myError_Message_Not_Valid_alphanumeric ' Not valid (caracters and/or digits)
Dim myError_Message_Not_Valid_alphanumeric_without_blank ' Non valide (caracters, digits, no Blank)
Dim myError_Message_Not_a_Valid_Directory_name
Dim myError_Message_Not_a_Valid_numerical 'Not valid (digits)
Dim myError_Message_Not_a_Valid_Email 'Not a valid Email
Dim myError_Message_Not_a_Valid_French_zip_Code 'Not a valid French zip code
Dim myError_Message_Not_Valid_Phone_Number 'Not a valid phone number
Dim myError_Message_Not_a_Valid_City_State_name 
Dim myError_Message_Not_a_Valid_File_Name 
Dim myError_Message_Too_Small      
Dim myError_Message_Too_Big
Dim myError_Message_Not_enough_caracters
Dim myError_Message_Too_much_caracters
Dim myError_Message_color_code

%>
<%
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' NewsGroups2 Messages
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim myNewsGroups2_Message_Forum
 Dim myNewsGroups2_Message_Forum_Members
 Dim myNewsGroups2_Message_Forum_No
 Dim myNewsGroups2_Message_Forum_Open
 Dim myNewsGroups2_Message_Forum_Close
 Dim myNewsGroups2_Message_Forum_Description
 Dim myNewsGroups2_Message_Subject
 Dim myNewsGroups2_Message_Topics
 Dim myNewsGroups2_Message_Post
 Dim myNewsGroups2_Message_Post_No
 Dim myNewsGroups2_Message_Post_by
 Dim myNewsGroups2_Message_Post_Posted
 Dim myNewsGroups2_Message_Post_Updated
 Dim myNewsGroups2_Message_Last_Post
 Dim myNewsGroups2_Message_Moderator
 Dim myNewsGroups2_Message_Moderator_All
 Dim myNewsGroups2_Message_Forum_Moderator_All
 Dim myNewsGroups2_Message_Moderator_Advert1
 Dim myNewsGroups2_Message_Moderator_Advert2 
 Dim myNewsGroups2_Message_Forum_User_Level
 Dim myNewsGroups2_Message_Forum_Public
 Dim myNewsGroups2_Message_Forum_Private
 Dim myNewsGroups2_Message_Forum_Edit
 Dim myNewsGroups2_Message_Forum_Add
 Dim myNewsGroups2_Message_Forum_Delete
 Dim myNewsGroups2_Message_Forum_Delete_Advert1
 Dim myNewsGroups2_Message_Forum_Delete_Advert2
 Dim myNewsGroups2_Message_Forum_Delete_Advert3
 Dim myNewsGroups2_Message_Forum_Delete_Advert4
 Dim myNewsGroups2_Message_Forum_Closed
 Dim myNewsGroups2_Message_Moderator_No
 Dim myNewsGroups2_Message_Moderator_Add
 Dim myNewsGroups2_Message_Moderator_Delete
 Dim myNewsGroups2_Message_Replies
 Dim myNewsGroups2_Message_Count
 Dim myNewsGroups2_Message_Author
 Dim myNewsGroups2_Message_New_Post
 Dim myNewsGroups2_Message_Reply_Post
 Dim myNewsGroups2_Message_Post_Topic_New_Title
 Dim myNewsGroups2_Message_Post_Topic_Update_Title
 Dim myNewsGroups2_Message_Post_Reply_New_Title
 Dim myNewsGroups2_Message_Post_Reply_Update_Title
 Dim myNewsGroups2_Message_Post_Delete_Title
 
%>
<%
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' Styles Messages
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim myStyles_Message_Name
 Dim myStyles_Message_Global_Width
 Dim myStyles_Message_Left_Width
 Dim myStyles_Message_Right_Width
 Dim myStyles_Message_Application_Width
 Dim myStyles_Message_BGColor
 Dim myStyles_Message_BGImage
 Dim myStyles_Message_BGTextColor
 Dim myStyles_Message_ApplicationColor
 Dim myStyles_Message_ApplicationImage
 Dim myStyles_Message_ApplicationTextColor
 Dim myStyles_Message_BorderColor
 Dim myStyles_Message_BorderImage
 Dim myStyles_Message_BorderTextColor
 Dim myStyles_Message_Date_Update
 Dim myStyles_Message_Author_Update
 Dim myStyles_Message_Modify
 Dim myStyles_Message_Border
 Dim myStyles_Message_Back
 Dim myStyles_Message_Add
 Dim myStyles_Message_Delete
 Dim myStyles_Message_ChooseColor
 Dim myStyles_Message_Color_Panel
 Dim myStyles_Message_BG
 Dim myStyles_Message_Application
 Dim myStyles_Message_WidthErrorNoChange
 Dim myStyles_Message_Panel
 Dim myStyles_Message_Name_Error
 Dim myStyles_Message_Coherence_problem
 Dim myStyles_Administration
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DB UPDATE MESSAGE 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim myDB_Message_Copied
 Dim myDB_Message_End
 Dim myDB_Message_Importation
 Dim myDB_Message_Import
 Dim myDB_Message_Question
 Dim myDB_Message_Question2
 Dim myDB_Message_Copy
 Dim myDB_Message_Patient
 Dim myDB_Message_CheckFile
 Dim myDB_Message_Cant_Proceed
 Dim myDB_Message_Description
 Dim myDB_Message_Adv
 Dim myDB_Message_Done
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
'FILES MESSAGE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
 
 Dim my_File_Message_Folder_Name
 Dim my_File_Message_Folder_Short_Description
 Dim my_File_Message_Public
 Dim my_File_Message_Check_Members
 Dim my_File_Message_Check_Public
 Dim my_File_Message_Description
 Dim my_File_Message_Folder
 Dim my_File_Message_Acces
 Dim my_File_Message_Last_Upload
 Dim my_File_Message_Pulic_list
 Dim my_File_Message_Private
 Dim my_File_Message_None
 Dim my_File_Message_Folder_Creator
 Dim my_File_Message_File_Number
 Dim my_File_Message
  Dim my_File_Message_Files
 Dim my_File_Message_Type
 Dim my_File_Message_Size 
 Dim myFile_Message_Empty
 Dim myFile_Message_File_Creator
 Dim myFile_Message_Folder_Name_Invalid
 Dim myFile_Message_Extension
 Dim myFile_Message_Maximum_Size
 Dim myFile_Message_Maximum_Size2
 Dim myFile_Message_Add_Extension
 Dim myFile_Message_File_Too_Big
 Dim myFile_Message_Extension_Not_Allowed
 Dim myFile_Message_Delete_Folder
 Dim myFile_Message_Files_Administration
 Dim myFile_Message_Files_Exists

Dim myMessage_Date_Format
Dim myMessage_Hour_Format

Dim myMessage_Personnal_Agenda
Dim myMessage_Global_Agenda
Dim myMessage_Global_Agenda_Members
Dim myMessage_Meeting_Confidentiality
Dim myMessage_Agenda_Start
Dim myMessage_Agenda_End
Dim myMessage_Date_Europe
Dim myMessage_Date_US
Dim myMessage_Hour_Europ
Dim myMessage_Hour_US
Dim myMessage_Agenda_Week_start



''''''''''''''''''''''''''''''''''''''''''''''''''''
'GROUPS MESSAGES
''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim myMessage_Group
Dim myMessage_Group_Administrator

''''''''''''''''''''''''''''''''''''''''''''''''''''
'SQL SERVER MESSAGES
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim myMessage_SQL_Administration  
Dim myMessage_Parameters
Dim myMessage_Server 
Dim myMessage_Database
Dim myMessage_Connection_established 
Dim myMessage_error_parameters
Dim myMessage_error
Dim myMessage_Choose
Dim myMessage_Create_Table
Dim myMessage_skip_creation
Dim myMessage_current_version
Dim myMessage_Warning1
Dim myMessage_Warning2
Dim myMessage_Warning3
Dim myMessage_Leave
Dim myMessage_create_new_site
Dim myMessage_import_from_base
Dim myMessage_sql_mode
Dim myMessage_Connection


%>
<%
  select Case myCurrent_Language
	Case myFrench_Language %>
	  <!-- #include file="Language_French.asp" -->
	<% Case mySpanish_Language %>
	  <!-- #include file="Language_Spanish.asp" -->
	<% Case myPortuguese_Language %>
	  <!--#include file="Language_Portuguese.asp" -->
	<% Case myGerman_Language %>                   
	  <!--#include file="Language_German.asp" --> 
	<% Case myItalian_Language %>                   
	  <!--#include file="Language_Italian.asp" -->  
<% Case else %>
     <!-- #include file="Language_English.asp" -->
<% end select%>





