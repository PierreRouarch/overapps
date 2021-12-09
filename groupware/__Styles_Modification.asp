<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Styles_Modification.asp" is free software; 
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
' Name 		       	: __Styles_Modification.asp
' Path   	      	: /
' Version 	    	: 1.15.0
' Description   	: Styles Modification
' By		        	: Dania TCHERKEZOFF
' Company	      	: OverApps
' Date			      : October 22, 2001
' Modify by		    :
' Company		      :
' Date			      :
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
if myApplication_Public_type_ID < myUser_type_ID then
	Response.redirect("__Quit.asp")
end if
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET APPLICATION TITLE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myApplication_Title = Get_Application_Title(myPage_Application)

Dim mySearch, myMaxRspByPage, mySortStyleDirectory_ID, myStyleDirectory_ID, myStyleDirectory_Name, mySortStyle_Name, mySortStyle_Description_Short, myOrder, myRs,  myNumPage, myNbrPage, indice, myStyle_URL, myInfo, myModif
Dim mySQL_Update_Style,myAction,Style_id,myUpdate_Style_id,mySet_tb_Styles_update,mySet_tb_Style_update,myUpdate_Style_ApplicationTextColor,mySQL_Select_ID_Styles

Dim myList_Styles_Name, myList_Styles_Global_Width , myList_Styles_Left_Width , myList_Styles_Right_Width, myList_Styles_Application_Width, myList_Styles_BGColor, myList_Styles_BGImage, myList_Styles_BGTextColor, myList_Styles_BorderColor, myList_Styles_BorderImage, myList_Styles_BorderTextColor, myList_Styles_ApplicationColor, myList_Styles_ApplicationImage, myList_Styles_ApplicationTextColor, myList_Styles_Date_Update, myList_Styles_Author_Update,myList_Styles_ID,mySet_ID_Styles
Dim myUpdate_Style_Name
Dim myUpdate_Style_GlobalWidth
Dim myUpdate_Style_LeftWidth
Dim myUpdate_Style_ApplicationWidth
Dim myUpdate_Style_BGColor
Dim myUpdate_Style_BGTextColor
Dim myUpdate_Style_ApplicationColor
Dim myUpdate_Style_AppliationTextColor
Dim myUpdate_Style_BorderColor
Dim myUpdate_Style_BorderTextColor
Dim mySQL_Delete_Styles
Dim myStyleError,myParameter
Dim myColor
Dim Style_Update_Parameter
Dim myRedirection
Dim myKnownColor(120)


'Known Colors for verification
'Those colors doesn't require a Hexa Decimal Code

myKnownColor(0)="Gainsboro"
myKnownColor(1)="OldLace"
myKnownColor(2)="Snow"
myKnownColor(3)="WhiteSmoke"
myKnownColor(4)="FloralWhite"
myKnownColor(5)="Linen"
myKnownColor(6)="AntiqueWhite"
myKnownColor(7)="PapayaWhip"
myKnownColor(8)="BlanchedAlmond"
myKnownColor(9)="Bisque"
myKnownColor(10)="PeachPuff"
myKnownColor(11)="NavajoWhite"
myKnownColor(12)="Mocassin"
myKnownColor(13)="CornSilk"
myKnownColor(14)="LemonChiffon"
myKnownColor(15)="Ivory"
myKnownColor(16)="HoneyDew"
myKnownColor(17)="MintCream"
myKnownColor(18)="Azure"
myKnownColor(19)="AliceBlue"
myKnownColor(20)="Lavender"
myKnownColor(21)="LavenderBlush"
myKnownColor(22)="MistyRose"
myKnownColor(23)="White"
myKnownColor(24)="Black"
myKnownColor(25)="MidnightBlue"
myKnownColor(26)="NavyBlue"
myKnownColor(27)="CornFlowerBlue"
myKnownColor(28)="DarkSlateBlue"
myKnownColor(29)="MediumSlateBlue"
myKnownColor(30)="LightSlateBlue"
myKnownColor(31)="MediumBlue"
myKnownColor(32)="RoyalBlue"
myKnownColor(33)="Blue"
myKnownColor(34)="DodgerBlue"
myKnownColor(35)="DeepSkyBlue"
myKnownColor(36)="MediumSkyBlue"
myKnownColor(37)="SkyBlue"
myKnownColor(38)="LightSkyBlue"
myKnownColor(39)="SteelBlue"
myKnownColor(40)="LightSteelBlue"
myKnownColor(41)="LightBlue"
myKnownColor(42)="PowderBlue"
myKnownColor(43)="PaleTurquoise"
myKnownColor(44)="DarkTurquoise"
myKnownColor(45)="MediumTurquoise"
myKnownColor(46)="Turqoise"
myKnownColor(47)="Cyan"
myKnownColor(48)="LightCyan"
myKnownColor(49)="CadetBlue"
myKnownColor(50)="MediumAquamarine"
myKnownColor(51)="Aquamarine"
myKnownColor(52)="DarkGreen"
myKnownColor(53)="DarkOliveGreen"
myKnownColor(54)="DarkSeaGreen"
myKnownColor(55)="SeaGreen"
myKnownColor(56)="MediumSeaGreen"
myKnownColor(57)="LightSeaGreen"
myKnownColor(58)="PaleGreen"
myKnownColor(59)="SpringGreen"
myKnownColor(60)="LawnGreen"
myKnownColor(70)="Green"
myKnownColor(71)="Chartreuse"
myKnownColor(72)="GreenYellow"
myKnownColor(73)="LimeGreen"
myKnownColor(74)="YellowGreen"
myKnownColor(75)="ForestGreen"
myKnownColor(76)="OliveDrab"  
myKnownColor(78)="DarkKhaki"  
myKnownColor(79)="PaleGoldenrod"
myKnownColor(80)="Yellow"     
myKnownColor(81)="Gold"                              
myKnownColor(81)="Goldenrod"                         
myKnownColor(82)="DarkGoldenrod"                     
myKnownColor(83)="RosyBrown"                         
myKnownColor(84)="IndianRed"                         
myKnownColor(85)="SaddleBrown"                       
myKnownColor(86)="Sienna"
myKnownColor(87)="Peru"
myKnownColor(88)="BurlyWood"
myKnownColor(89)="Wheat"
myKnownColor(90)="SandyBrown"
myKnownColor(91)="Tan"
myKnownColor(92)="Chocolate"
myKnownColor(93)="FireBrick"
myKnownColor(94)="Brown"
myKnownColor(95)="Salmon"
myKnownColor(96)="LightSalmon"
myKnownColor(97)="DarkSalmon"
myKnownColor(98)="Orange"
myKnownColor(99)="DarkOrange"
myKnownColor(100)="Coral"
myKnownColor(101)="LightCoral"
myKnownColor(102)="Tomato"
myKnownColor(103)="OrangeRed"
myKnownColor(104)="Red"
myKnownColor(105)="HotPink"
myKnownColor(106)="DeepPink"
myKnownColor(107)="LightPink"
myKnownColor(108)="Pink"
myKnownColor(108)="PaleVioletRed"
myKnownColor(108)="Maroon"
myKnownColor(109)="MediumVioletRed"
myKnownColor(110)="Magenta"
myKnownColor(111)="Violet"
myKnownColor(112)="Plum"
myKnownColor(113)="Orchid"
myKnownColor(114)="MediumOrchid"
myKnownColor(115)="DarkOrchid"
myKnownColor(116)="DarkViolet"
myKnownColor(117)="BlueViolet"
myKnownColor(118)="Purple"
myKnownColor(119)="MediumPurple"
myKnownColor(120)="Thistle"
'''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Parameters from form or URL string
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
myAction = Request.form("Action")

myStyle_id = request.form("Style_id")

if len(myStyle_id)=0 then
	myStyle_id=request.querystring("Style_id")
end if	
	
if len(myStyle_id)=0 then
	myStyle_id=request.querystring("myStyle_id")	
end if	

if len(myAction)=0 then
	myAction=request.querystring("action")
end if	

if len(myStyle_id) = 0 then myStyle_id=request.querystring("myStyle_id_selected")
if len(myAction)=0 Then myAction="Modify"



if len(request.form("ModifBGColor")) <> 0 Then myParameter="BGColor"
if len(request.form("ModifBGTextColor")) <> 0 Then myParameter="BGTextColor"
if len(request.form("ModifBorderColor")) <> 0 Then myParameter="BorderColor"
if len(request.form("ModifBorderTextColor")) <> 0 Then myParameter="BorderTextColor"
if len(request.form("ModifApplicationColor")) <>0 Then myParameter="ApplicationColor"
if len(request.form("ModifApplicationTextColor")) <> 0 Then myParameter = "ApplicationTextColor"

myStyleError = Request.QueryString("myStyleError")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UPDATING  A COLOR  SENT BY THE PANEL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if myAction = "changecolor"  Then



myParameter = Request.QueryString("myColorParameter")
myColor=Request.QueryString("myColor")

Style_Update_Parameter = "Style_" & myParameter


'Open Connection
 set myConnection = Server.CreateObject("ADODB.Connection")
 myConnection.Open myConnection_String


'Selection of the style to be updated
 mySQL_Select_tb_Styles = "SELECT * FROM tb_Styles WHERE Style_ID=" & myStyle_id 
 Set mySet_tb_Styles_update = server.createobject("adodb.recordset")
 mySet_tb_Styles_update.open mySQL_Select_tb_Styles, myConnection, 3,3
 
 response.write Style_Update_Parameter
 mySet_tb_Styles_update.fields(Style_Update_Parameter) = "#"& MyColor

'UPDATE IN DB  and close conection
 mySet_tb_Styles_update.Update
 mySet_tb_Styles_update.close
 set mySet_tb_Styles_update = Nothing


 myRedirection= "__Styles_Modification.asp?myStyle_id_selected=" & myStyle_ID & "&myStyleError=" & myStyleError

 response.redirect(myRedirection)
 
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UPDATE 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


if (Request.form("Validation")=myMessage_Go or len(myParameter) <> 0 )  and myAction <> "Add" then
myUpdate_Style_Name                 =request.form("Name") 
myUpdate_Style_BGColor              =request.form("BGColor") 
myUpdate_Style_BGTextColor          =request.form("BGTextColor")
myUpdate_Style_ApplicationColor     =request.form("ApplicationColor")
myUpdate_Style_ApplicationTextColor =request.form("ApplicationTextColor") 
myUpdate_Style_BorderColor          =request.form("BorderColor") 
myUpdate_Style_BorderTextColor      =request.form("BorderTextColor")
myUpdate_Style_ApplicationWidth     =request.form("Application_Width")
myUpdate_Style_LeftWidth            =request.form("Left_Width") 
myUpdate_Style_GlobalWidth          =request.form("Global_Width")

'Open Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String

'Selection of the style to be updated
 mySQL_Select_tb_Styles = "SELECT * FROM tb_Styles WHERE Style_ID=" & myStyle_id 
 Set mySet_tb_Styles_update = server.createobject("adodb.recordset")
 mySet_tb_Styles_update.open mySQL_Select_tb_Styles, myConnection, 3,3


'TEST OF NUMERICAL VALUE FOR WIDTH / IF NOT A NUMERICAL, VALUE IS NOT CHANGED

if myValueIsGood (myerr_numerical, myUpdate_Style_GlobalWidth)= True And len(myUpdate_Style_GlobalWidth) > 0 Then 
	mySet_tb_Styles_update.fields("Style_Global_Width") = myUpdate_Style_GlobalWidth
else myStyleError="A"
end if
	
if myValueIsGood (myerr_numerical, myUpdate_Style_LeftWidth)= True And len(myUpdate_Style_LeftWidth) > 0 then
	mySet_tb_Styles_update.fields("Style_Left_Width") = myUpdate_Style_LeftWidth 	
else myStyleError=myStyleError & "B"
end if

if myValueIsGood (myerr_numerical, myUpdate_Style_ApplicationWidth)= True AND len(myUpdate_Style_ApplicationWidth) > 0 Then
	mySet_tb_Styles_update.fields("Style_Application_Width") = myUpdate_Style_ApplicationWidth
else myStyleError=myStyleError & "C"
end if



'TEST FOR COHERENCE OF WIDTH VALUES 

if(mySet_tb_Styles_update.fields("Style_Left_Width") + mySet_tb_Styles_update.fields("Style_Application_Width")) > mySet_tb_Styles_update.fields("Style_Global_Width") Then myStyleError=myStyleError & "K"


'TEST FOR COLOR CODE

if (myValueIsGood (myerr_color_code, myUpdate_Style_BGColor)= True) OR (Ubound(	filter(myKnownColor, myUpdate_Style_BGColor	,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_BGColor") = myUpdate_Style_BGColor
else myStyleError=myStyleError&"D"
end if


if (myValueIsGood (myerr_color_code, myUpdate_Style_BGTextColor)= True) OR (Ubound(	filter(myKnownColor, myUpdate_Style_BGTextColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_BGTextColor") = myUpdate_Style_BGTextColor
else myStyleError=myStyleError&"E"
end if

if (myValueIsGood (myerr_color_code, myUpdate_Style_BorderColor)= True) OR (Ubound(	filter(myKnownColor, myUpdate_Style_BorderColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_BorderColor") = myUpdate_Style_BorderColor
else myStyleError=myStyleError&"F"
end if

if (myValueIsGood (myerr_color_code, myUpdate_Style_BorderTextColor)= True) OR (Ubound(	filter(myKnownColor, myUpdate_Style_BorderTextColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_BorderTextColor") = myUpdate_Style_BorderTextColor
else myStyleError=myStyleError&"G"
end if

if (myValueIsGood (myerr_color_code, myUpdate_Style_ApplicationColor)= True) OR (Ubound(filter(myKnownColor, myUpdate_Style_ApplicationColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_ApplicationColor") = myUpdate_Style_ApplicationColor
else myStyleError=myStyleError&"H"
end if

if (myValueIsGood (myerr_color_code, myUpdate_Style_ApplicationTextColor)= True) OR (Ubound(filter(myKnownColor, myUpdate_Style_ApplicationTextColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_ApplicationTextColor") = myUpdate_Style_ApplicationTextColor
else myStyleError=myStyleError&"I"
end if

'TEST FOR STYLE NAME

if len(myUpdate_Style_Name) >0 Then
 mySet_tb_Styles_update.fields("Style_Name") = myUpdate_Style_Name
else  
  myStyleError=myStyleError & "J"
end if


'END OF TEST

mySet_tb_Styles_update.fields("Member_ID") = myUser_ID
mySet_tb_Styles_update.fields("Style_Date_Update")=myDate_Now()
mySet_tb_Styles_update.fields("Style_Author_Update")= myUser_Login

'UPDATE IN DB  and close conection
mySet_tb_Styles_update.Update
mySet_tb_Styles_update.close
set mySet_tb_Styles_update = Nothing


'REDIRECTION TO THE PANEL COLOR IF USER WANT TO CHANGE A COLOR PARAMETER

if len(myParameter) <>0 Then
Response.redirect("__Styles_Color.asp?myStyle_ID=" & myStyle_id &"&myParameter="& myParameter & "&myStyleError=" & myStyleError )
end if

'Redirection to this page for refresh 	
Response.redirect("__Styles_Modification.asp?MyStyle_id_selected=" & myStyle_id & "&myStyleError=" & myStyleError)

end if 'END OF UPDATE


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ADD A NEW STYLE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


if (Request.form("Validation")=myMessage_Go or  len(myParameter) <>0 )and myAction = "Add" then


myUpdate_Style_Name=request.form("Name") 
myUpdate_Style_GlobalWidth=request.form("Global_Width")
myUpdate_Style_LeftWidth=request.form("Left_Width") 
myUpdate_Style_ApplicationWidth=request.form("Application_Width")
myUpdate_Style_BGColor=request.form("BGColor") 
myUpdate_Style_BGTextColor=request.form("BGTextColor")
myUpdate_Style_ApplicationColor=request.form("ApplicationColor")
myUpdate_Style_ApplicationTextColor=request.form("ApplicationTextColor") 
myUpdate_Style_BorderColor=request.form("BorderColor") 
myUpdate_Style_BorderTextColor=request.form("BorderTextColor")



'Open Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


mySQL_Select_tb_Styles = "SELECT * FROM tb_Styles "
Set mySet_tb_Styles_update = server.createobject("adodb.recordset")
mySet_tb_Styles_update.open mySQL_Select_tb_Styles, myConnection, 3,3
mySet_tb_Styles_update.AddNew


'TEST OF NUMERICAL VALUE FOR WIDTH / IF NOT A NUMERICAL VALUE IS NOT ADDED
if myValueIsGood (myerr_numerical, myUpdate_Style_GlobalWidth)= True and len(myUpdate_Style_GlobalWidth)>0  then 
	mySet_tb_Styles_update.fields("Style_Global_Width") = myUpdate_Style_GlobalWidth	

else myStyleError="A"

end if
	
if myValueIsGood (myerr_numerical, myUpdate_Style_LeftWidth)= True and len(myUpdate_Style_LeftWidth) > 0 then
	mySet_tb_Styles_update.fields("Style_Left_Width") = myUpdate_Style_LeftWidth 	

else myStyleError=myStyleError & "B"

end if

if myValueIsGood (myerr_numerical, myUpdate_Style_ApplicationWidth)= True AND len(myUpdate_Style_ApplicationWidth) > 0 then
	mySet_tb_Styles_update.fields("Style_Application_Width") = myUpdate_Style_ApplicationWidth

else myStyleError=myStyleError & "C"		

end if


'TEST FOR COLOR CODE

if (myValueIsGood (myerr_color_code, myUpdate_Style_BGColor)= True) OR (Ubound(	filter(myKnownColor, myUpdate_Style_BGColor	,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_BGColor") = myUpdate_Style_BGColor
else myStyleError=myStyleError & "D"	
end if


if (myValueIsGood (myerr_color_code, myUpdate_Style_BGTextColor)= True) OR (Ubound(	filter(myKnownColor, myUpdate_Style_BGTextColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_BGTextColor") = myUpdate_Style_BGTextColor
else myStyleError=myStyleError & "E"	
end if

if (myValueIsGood (myerr_color_code, myUpdate_Style_BorderColor)= True) OR (Ubound(	filter(myKnownColor, myUpdate_Style_BorderColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_BorderColor") = myUpdate_Style_BorderColor
else myStyleError=myStyleError & "F"	
end if

if (myValueIsGood (myerr_color_code, myUpdate_Style_BorderTextColor)= True) OR (Ubound(filter(myKnownColor, myUpdate_Style_BorderTextColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_BorderTextColor") = myUpdate_Style_BorderTextColor
else myStyleError=myStyleError & "G"	
end if

if (myValueIsGood (myerr_color_code, myUpdate_Style_ApplicationColor)= True) OR (Ubound(filter(myKnownColor, myUpdate_Style_ApplicationColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_ApplicationColor") = myUpdate_Style_ApplicationColor
else myStyleError=myStyleError & "H"	
end if

if (myValueIsGood (myerr_color_code, myUpdate_Style_ApplicationTextColor)= True) OR (Ubound(filter(myKnownColor, myUpdate_Style_ApplicationTextColor,true)) > 0 )then
	mySet_tb_Styles_update.fields("Style_ApplicationTextColor") = myUpdate_Style_ApplicationTextColor
else myStyleError=myStyleError & "I"	
end if


'TEST FOR COHERENCE OF WIDTH VALUES 

if ( mySet_tb_Styles_update.fields("Style_Left_Width") + mySet_tb_Styles_update.fields("Style_Application_Width") ) > mySet_tb_Styles_update.fields("Style_Global_Width") Then 
myStyleError=myStyleError & "K"
end if


'TEST FOR STYLE NAME
if len(myUpdate_Style_Name) >0 Then
 mySet_tb_Styles_update.fields("Style_Name") = myUpdate_Style_Name
else  
  myStyleError=myStyleError & "J"
end if



'END OF TEST


mySet_tb_Styles_update.fields("Style_Name") = myUpdate_Style_Name

mySet_tb_Styles_update.fields("Member_ID") = myUser_ID
mySet_tb_Styles_update.fields("Style_Date_Update")=myDate_Now()
mySet_tb_Styles_update.fields("Style_Author_Update")= myUser_Login
mySet_tb_Styles_update.fields("Site_ID")=mySite_ID

mySet_tb_Styles_update.Update
mySet_tb_Styles_update.close
set mySet_tb_Styles_update = Nothing




'GET THE ID OF THE NEW STYLE

mySQL_Select_ID_Styles = "SELECT max(tb_Styles.Style_ID) as ID FROM tb_Styles"
set mySet_ID_Styles = myConnection.Execute(mySQL_Select_ID_Styles)

 if not mySet_ID_Styles.bof then 
	 mySet_ID_Styles.MoveFirst
 end if


myStyle_id =  mySet_ID_Styles("ID")

' Close Connection 
		myConnection.close
		set myConnection = nothing


'REDIRECTION TO THE PANEL COLOR IF USER WANT TO CHANGE A PARAMETER

if len(myParameter) <>0 Then
Response.redirect("__Styles_Color.asp?myStyle_ID=" & myStyle_id &"&myParameter="& myParameter & "&myStyleError=" & myStyleError)
end if


myAction="Modify"

end if 


''''''''''''''''''''''''''''''''''''''''''''''
' DELETE                '							
''''''''''''''''''''''''''''''''''''''''''''''

if myAction = "delete" and len(myStyle_id) > 0 Then

'Open Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


mySQL_Delete_Styles = "Delete From tb_Styles where Style_ID ="& myStyle_id
myConnection.Execute(mySQL_Delete_Styles)



' Close Connection 
		myConnection.close
		set myConnection = nothing

response.redirect("__Styles_List.asp")


end if 'end of delete



%>
<html>

<head>
<title><%=mySite_Name%>  Styles - Modification -</title>
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

<TD width="<%=myApplication_Width%>" BGCOLOR="<%=myBGColor%>" VALIGN="top" ALIGN="left"> 

<%
' APPLICATION TITLE
%>
<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>"><font face="Arial, Helvetica, sans-serif" size="4" color="<%=myApplicationTextColor%>"><b><%=myApplication_Title%></b></font></TD></TR> 
</table>

<%
' CENTER APPLICATION

%> 

<%
'Open Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String


'''''''''''''''''''''''
'MODIFICATION FORM    '
'''''''''''''''''''''''
if myAction = "Modify" then

'GET STYLE INFORMATION FOR AN UPDATE

if len(myStyle_id)<> 0 then
' Execute SQL QUERY
 mySQL_Select_tb_Styles = "SELECT * FROM tb_Styles WHERE tb_Styles.Style_ID=" & myStyle_id

 set mySet_tb_Styles = myConnection.Execute(mySQL_Select_tb_Styles)

 if not mySet_tb_Styles.bof then 
	 mySet_tb_Styles.MoveFirst
 end if
' Get Styles Informations

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

 ' Close Connection 
		myConnection.close
		set myConnection = nothing
 end if
%>
<%
' FORM BOXS
%>
<form action="__Styles_Modification.asp?Style_id=<%=myStyle_id%>" method=post encrypt=multipart/data>
<input type=hidden name=Style_id value=<%=myStyle_id%>>
<%
if len(myStyle_id) = 0 then
%>
<input type=hidden name=Action value=Add>
<%
end if
%>

<%
 'DISPLAY OF THE CURRENT STYLE               
%> 
            <table border="0"  cellpadding="1" cellspacing="1" align=center width="20%">
              <tr> 
                <td bgcolor="<%=myList_Styles_BGColor%>" height="10"> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myList_Styles_BGTextColor%>"><strong> 
                  <%=myStyles_Message_BG%> </strong></font> </td>
              </tr>
              <tr> 
                <td valign="top" align="left" bgcolor="<%=myList_Styles_BorderColor%>" height="10"> 
                  <b><font face="Arial, Helvetica, sans-serif" size="2"  color="<%=myList_Styles_BorderTextColor%>"><%=myStyles_Message_Border%></font></b> 
                </td>
              </tr>
              <tr> 
                <td bgcolor="<%=myList_Styles_ApplicationColor%>" height="10"><font face="Arial, Helvetica, sans-serif" size="4"  color="<%=myList_Styles_ApplicationTextColor%>"><b><%=myStyles_Message_Application%></b></font></td>
              </tr>
            </table>
<br>
        <table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="1" >
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_Name%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <input type=text value="<%=myList_Styles_Name%>" name=Name>
<%
'DISPLAY ERROR FOR BAD SYLE NAME

if len(myStyleError) <= 0 Then myStyleError=Request.QueryString("myStyleError")

if (InStr(myStyleError,"J") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myStyles_Message_Name_Error%>
              </font> <%
end if
'END ERROR
%>              
            </td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_Global_Width%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <input type=text value="<%=myList_Styles_Global_Width%>" size="4" name=Global_Width>
<%
'DISPLAY ERROR FOR BAD GLOBAL WIDTH



if (InStr(myStyleError,"A") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_Not_a_Valid_numerical%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%> 
<%
'DISPLAY ERROR FOR  WIDTH inconsistency


if (InStr(myStyleError,"K") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%= myStyles_Message_Coherence_problem%></font> <%
end if
'END ERROR
%>

</td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_Left_Width%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <input type=text value="<%=myList_Styles_Left_Width%>" size="4" name=Left_Width>
              <%
'DISPLAY ERROR FOR BAD LEFT WIDTH


if (InStr(myStyleError,"B") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_Not_a_Valid_numerical%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%> </td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" rowspan="2" width="45%" height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_Application_Width%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left rowspan="2" colspan="2" height="10"> 
              <input type=text value="<%=myList_Styles_Application_Width%>" size="4" name=Application_Width>
              <%
'DISPLAY ERROR FOR BAD Application WIDTH


if (InStr(myStyleError,"C") > 0) Then 
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_Not_a_Valid_numerical%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%> </td>
          </tr>
          <tr> </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_BGColor%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2" height="10"> 
              <input type=text value="<%=myList_Styles_BGColor%>" size="7" name=BGColor>
              <input type=submit name=ModifBGColor value=<%=myStyles_Message_Panel%>>
<%
'DISPLAY ERROR FOR BAD BGColor

if (InStr(myStyleError,"D") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_color_code%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%>
 </td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="2"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_BGTextColor%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2" valign="middle" height="2"> 
              <input type=text value="<%=myList_Styles_BGTextColor%>" size="7" name=BGTextColor>
              <input type=submit name=ModifBGTextColor value=<%=myStyles_Message_Panel%>>
<%
'DISPLAY ERROR FOR BAD BGTextColor

if (InStr(myStyleError,"E") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_color_code%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%>              
            </td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_BorderColor%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <input type=text value="<%=myList_Styles_BorderColor%>" size="7" name=BorderColor>
              <input type=submit name=ModifBorderColor value=<%=myStyles_Message_Panel%>>
<%
'DISPLAY ERROR FOR BAD BorderColor

if (InStr(myStyleError,"F") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_color_code%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%>              
            </td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_BorderTextColor%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <input type=text value="<%=myList_Styles_BorderTextColor%>" size="7" name=BorderTextColor>
              <input type=submit name=ModifBorderTextColor value=<%=myStyles_Message_Panel%>>
<%
'DISPLAY ERROR FOR BAD BoderTextColor

if (InStr(myStyleError,"G") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_color_code%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%>              
            </td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_ApplicationColor%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <input type=text value="<%=myList_Styles_ApplicationColor%>" size="7" name=ApplicationColor>
              <input type=submit name=ModifApplicationColor value=<%=myStyles_Message_Panel%>>
<%
'DISPLAY ERROR FOR BAD Applicationcolor

if (InStr(myStyleError,"H") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_color_code%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%>              
            </td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBorderTextColor%>"> 
              <b><%=myStyles_Message_ApplicationTextColor%> : </b> </font></td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <input type=text value="<%=myList_Styles_ApplicationTextColor%>" size="7" name=ApplicationTextColor>
              <input type=submit name=ModifApplicationTextColor value=<%=myStyles_Message_Panel%>>
<%
'DISPLAY ERROR FOR BAD ApplicationTextColor

if (InStr(myStyleError,"I") > 0) Then
%> <font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><%=myError_Message_color_code%> 
              : <%=myStyles_Message_WidthErrorNoChange%></font> <%
end if
'END ERROR
%>              
            </td>
          </tr>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width=45% height="20"> 
              &nbsp;&nbsp;</td>
            <td bgcolor="<%=myBGColor%>" align=left colspan="2"> 
              <input type=submit value="<%=myMessage_GO%>" name=Validation>
            </td>
          </tr>
        </table>          
</form>


<table border="0" width="<%=myApplication_Width%>" cellpadding="3" cellspacing="0" > 
<TR><TD bgcolor="<%=myApplicationColor%>" align=middle><font face="Arial, Helvetica, sans-serif" size="1" color="<%=myApplicationTextColor%>">
<% if len(myList_Styles_Date_Update) > 0 then %> 
<% = myDate_Display(myList_Styles_Date_Update,2) %> -- <% = myList_Styles_Author_Update %>
<% end if %></font></TD></TR> 
</table>

<font face="Arial, Helvetica, sans-serif" size="2" color="<%=myBGTextColor%>"><a href=__Styles_List.asp><%=myStyles_Message_Back%></a> 
  , <a href=__Styles_Modification.asp?action=delete&Style_id=<%=myStyle_id%>><%=myStyles_Message_Delete%></a></font> 

<%
end if 'END OF FORM 
'End of Application
%> 

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
' End Copyright									'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%> 

</body>
</html>

<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>