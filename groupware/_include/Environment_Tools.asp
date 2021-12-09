<%
' ------------------------------------------------------------------------------------------------
' Copyright (C) 2001  + Ov-erA-pps - http://www.overapps.com
'
' This program "Environment_tools.asp" is free software; you can redistribute it 
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
' Name : Environment_tools.asp
' Path : /_Include
' Description : Functions for environment purpose
' By : Pierre Rouarch
' Company : OverApps
' Update : June, 26 2001
' 
' Contributors : 
' 
' Last Modification :
' Modify by : 
' Company : 
' Date :
' 
' Last Modifications : 	
'
' This is the OverApps official Source Version N° 1.18.0
' -------------------------------------------------------------------------------------------------



	Function Get_Application_Title(myCurrent_Application_Name)
	Dim i		
	i=0
	Get_Application_Title=mySite_Name
	Do While i<=myMax_User_Applications
		if myCurrent_Application_Name=myUser_Application_Name(i) then
			Get_Application_Title=myUser_Application_Title(i)
			i=myMax_User_Applications+1
		else 
			i=i+1
		end if
		Loop 
	End Function
	

	Function Get_Application_Public_Type_ID(myCurrent_Application_Name)
	Dim i		
	i=0
	Get_Application_Public_Type_ID=0
	Do While i<=myMax_User_Applications
		if myCurrent_Application_Name=myUser_Application_Name(i) then
			Get_Application_Public_Type_ID=myUser_Application_Public_Type_ID(i)
			i=myMax_User_Applications+1
		else 
			i=i+1
		end if
		Loop 
	End Function

' DATE FUNCTION

'Returns a string formated like this yyyy/mm/dd hh:mm:ss

function myDate_Now(  ) 
 Dim myYear, myMonth, myDay, myHour, myMinute, mySecond
 myYear =  Year(Now)
 myMonth  = Month(Now)
 if len(myMonth) = 1 Then myMonth = "0" & myMonth
 myDay  =   Day(Now)
 if len(myDay) = 1 Then myDay = "0" & myDay
 myHour = Hour(Now)
 if len(myHour) = 1 Then myHour = "0" & myHour
 myMinute = Minute(Now)
 if len(myMinute) = 1 Then myMinute = "0" & myMinute
 mySecond = Second(Now)
 if len(mySecond) = 1 Then mySecond = "0" & mySecond
 myDate_Now = myYear & "/" & myMonth & "/" & myDay & " " & myHour & ":" & myMinute & ":" & mySecond  
end function

'returns a string Formated Date to format choosen, need a string formated like this yyyy/mm/dd hh:mm:ss
'myDisplay : 1 For display only Date, 2 for date and time , 3 for only time 

function myDate_Display(myDate, myDisplay)
 Dim myYear, myMonth, myDay
 Dim myHour, myMinute, mySecond, myIndicator, myTime
if len(myDate) > 0 Then
 myYear = mid(myDate,1,4)
 myMonth = mid(myDate,6,2)

 myDay = mid(myDate,9,2)

 If myDate_Format = 1 Then
  myDate_Display = myMonth & "/" & myDay &"/" & myYear
 end if

 If myDate_Format = 2 Then
  myDate_Display = myDay & "/" & myMonth & "/" & myYear 
 end if

 If myDisplay  > 1 Then
 myHour = mid(myDate,12,2)

 if len(myHour) >0 Then
 If myHour > 12 AND myHour_Format = 1 Then
  myHour= myHour - 12
  myIndicator = "PM"
 else
  myIndicator = "AM" 
 end if
 
 myMinute = mid(myDate,15,2)
 mySecond= mid(myDate,18,2)

 myTime = myHour & ":" & myMinute &":" & mySecond 
 
 If myHour_Format = 1 Then
  myTime = myTime & " " & myIndicator
 end if
end if 
end if
 
 if myDisplay = 2 Then
  myDate_Display = myDate_Display & "  " & myTime
 end if
 
 if myDisplay = 3 Then
  myDate_Display = myTime
 end if  
end if 
end function


function myDate_Construct(myYear,myMonth,myDay,myHour,myMinute,mySecond)


 if len(myMonth) = 1 Then myMonth = "0" & myMonth
 if len(myDay) = 1 Then myDay = "0" & myDay
 if len(myHour) = 1 Then myHour = "0" & myHour
 if len(myMinute) = 1 Then myMinute = "0" & myMinute
 if len(mySecond) = 1 Then mySecond = "0" & mySecond

 myDate_Construct = myYear & "/" & myMonth & "/" & myDay & " " & myHour & ":" & myMinute & ":" & mySecond  

end function

function myYear(myDate)
 if len(myDate) > 9 Then myYear = mid(myDate,1,4)
end function 

function  myDay(myDate) 
 if len(myDate) > 9 Then myDay = mid(myDate,9,2)
end function

function myMonth(myDate)
 if len(myDate) > 9 Then myMonth = mid(myDate,6,2)
end function
%>



	