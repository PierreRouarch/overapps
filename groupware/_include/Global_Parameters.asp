<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001  - Pierre ROuarch
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
' 	" Copyright (C) 2001 OverApps & contributors "
' at the bottom of the page with an active link from the name "OverApps" to 
' the Address http://www.overapps.com where the user (netsurfer) could find the 
' appropriate information about this license. 
'
'-----------------------------------------------------------------------------
%>
<!--#include file="config.asp"-->
<%
' ------------------------------------------------------------
' Name : Global_Parameters.asp
' Path : /_INCLUDE
' Description : General Variables and Parameters - Must be changed 
' to fit to your environment
' by : Pierre Rouarch
' Last Update : November, 21, 2001
' Version : 1.18.0
'
' Contibutions : Dania Tcherkezoff
' ------------------------------------------------------------
function decrypt(myString)
 Dim mycode(27),myEntry(27)
 Dim  myChar,i,j
 mycode(1) = "b"
 mycode(2) = "c"
 mycode(3) = "d"
 mycode(4) = "e"
 mycode(5) = "g"
 mycode(6) = "h"
 mycode(7) = "i"
 mycode(8) = "j"
 mycode(9) = "k"
 mycode(10) = "u"
 mycode(11) = "a"
 mycode(12) = "v"
 mycode(13) = "w"
 mycode(14) = "x"
 mycode(15) = "y"
 mycode(16) = "z"
 mycode(17) = "p"
 mycode(18) = "q"
 mycode(19) = "t"
 mycode(20) = "r"
 mycode(21) = "s"
 mycode(22) = "o"
 mycode(23) = "n"
 mycode(24) = "m"
 mycode(25) = "l"
 mycode(26) = "f"
 
 myEntry(1)  = "a"
 myEntry(2)  = "b"
 myEntry(3)  = "c"
 myEntry(4)  = "d"
 myEntry(5)  = "e"
 myEntry(6)  = "f"
 myEntry(7)  = "g"
 myEntry(8)  = "h"
 myEntry(9)  = "i"
 myEntry(10) = "j"
 myEntry(11) = "k"
 myEntry(12) = "l"
 myEntry(13) = "m"
 myEntry(14) = "n"
 myEntry(15) = "o"
 myEntry(16) = "p"
 myEntry(17) = "q"
 myEntry(18) = "r"
 myEntry(19) = "s"
 myEntry(20) = "t"
 myEntry(21) = "u"
 myEntry(22) = "v"
 myEntry(23) = "w"
 myEntry(24) = "x"
 myEntry(25) = "y"
 myEntry(26) = "z"

 
 i = 1
 do while i < len(myString) + 1
 
 j = 1
 do while myCode(j) <> lcase(mid(myString,i,1)) 
  j = j +1 
  If j > 25 Then exit do  
 loop 
 
  myChar=  replace( lcase(mid(myString,i,1)), myCode(j),myEntry(j))
  decrypt = decrypt & myChar 

  i = i + 1
 loop
 
end function

'''''''''''''''''''''''''''''''''''''''''''''''
' Root Server					
'''''''''''''''''''''''''''''''''''''''''''''''
Dim myRoot
myRoot = server.mapPath("\")

''''''''''''''''''''''''''''''''''''''''''''''''
' Current Directory Path 
''''''''''''''''''''''''''''''''''''''''''''''''
Dim myCurrent_Directory_Path 
myCurrent_Directory_Path=server.mapPath(".")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data base Path and upload files path				
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim  myDatabase_Path,mySQL_Dont_Connect,myConnection_Access
Dim myConnection_String, myConnection_String_SQL,myConnection_String_Access

Dim myShared_Files_Path
Dim myShared_Files_Folder
Dim myShared_Files_Download_Path


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If you host your site at brinkster.com :
' for a free account you must put the database Overapps_V115X.mdb in the directory "\db"
' and take the following path
' myDatabase_Path=myRoot&"\DB\Overapps_V115X.mdb"
' Or try also
' myDatabase_Path= myCurrent_Directory_Path &"..\..\DB\Overapps_V115X.mdb"
' For a Premium Account you must put the database Overapps_V115X.mdb in 
' the directory "\YourAccount\database" and take the  following Path
' myDatabase_Path = "c:\sites\premium\YourAccount\Database\Overapps_V115X.mdb"
' Where YourAccount is Your Account Name given by Brinkster.com
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' if you host your site at Phidji.com the Path is 
' myDatabase_Path=myRoot&"\DB\Overapps_V115X.mdb"
' Or try also
' myDatabase_Path= myCurrent_Directory_Path &"..\..\DB\Overapps_V115X.mdb"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' if you host your site at efrance.com
' myDatabase_Path=myRoot&"\YourAccount\BaseDonnee\Overapps_V115X.mdb"
' Where YourAccount is Your Account Name given by efrance.com
' Or try also
' myDatabase_Path= myCurrent_Directory_Path &"..\..\BaseDonnee\Overapps_V115X.mdb"
' Or try also
' MyDatabase_Path= Server.MapPath("..") & "\BaseDonnee\Overapps_V115X.mdb"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' if you host your site at WebSamba.com
' myDatabase_Path=myRoot&"\YourAccount\DB\Overapps_V115X.mdb"
' Where YourAccount is Your Account Name given by WebSamba.com
' Or try also
' myDatabase_Path= myCurrent_Directory_Path &"..\..\DB\Overapps_V115X.mdb"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Standard Path (on your computer)
myDatabase_Path=Server.MapPath("..") & "\DB\Overapps_V118X.mdb"

myServer    = decrypt(myServer)
myDatabase  = decrypt(myDatabase)
myLogin     = decrypt(myLogin)
myPassword  = decrypt(myPassword)


myConnection_String_SQL = "DRIVER={SQL Server}; Server=" & myServer & "; Database=" & myDatabase & "; UID=" & myLogin 

myConnection_String_Access =  "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & myDatabase_Path



If mySQL_Enabled = 0 or mySQL_Dont_Connect = 1  Then 
  myConnection_String = myConnection_String_Access
 else
  myConnection_String = myConnection_String_SQL

end if  



' PATH TO OLD AND NEW DATABASES USED FOR DATA IMPORTATION 
Dim myNew_Database_Path,myOld_Database_Path,myOld_Database_Path_OS,myOld_Database_Path_Gw

myNew_Database_Path=server.mapPath("..") & "\DB\Overapps_V118X_empty.mdb"

'1.11.1 AND PREVIOUS DB PATH
myOld_Database_Path_OS=server.mapPath("\Overapps-software\database\")& "\"
'1.13X AND GREATER DB PATH
myOld_Database_Path_GW=server.mapPath("\Overapps\DB\")& "\"


'PATH FOR UPLOAD FILES

'Enter The Folder Name where file will be stored
myShared_Files_Folder = "SharedFiles"

myShared_Files_Path =server.mapPath("..") & "\" & myShared_Files_Folder& "\" 
myShared_Files_Download_Path = "..\" & myShared_Files_Folder& "\"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Languages
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const myEnglish_Language= 1
Const myFrench_Language = 2
Const mySpanish_Language = 3
Const myPortuguese_Language = 4
Const myGerman_Language = 5                            
Const myItalian_Language = 6

Dim myCurrent_Language
' Force Current language to French
myCurrent_Language= myEnglish_Language


' see in DB_Environment.asp Messages depending on  language


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Other Constants							   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Site's User's types

Const mySite_Adminitrator=1
Const mySite_Moderator = 2
Const mySite_Intranet_Member=3
Const mySite_Extranet_Member=4
Const mySite_Web_Member=5
Const mySite_Email_Member=6

Const mySite_User=2 ' For compatibility with old versions
%>
