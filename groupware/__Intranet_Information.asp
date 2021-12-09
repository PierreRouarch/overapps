<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001-2002   - OverApps - http://www.overapps.com
'
' This program "__Administration_Site_General.asp" is free software; you can redistribute it and/or modify
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
'-----------------------------------------------------------------------------
%>
<% 	Option Explicit 
	Response.Buffer = true	
	Response.ExpiresAbsolute = Now () - 1
	Response.Expires = 0
'	Response.CacheControl = "no-cache" 
' Does n't Work with PWS ???
%>
<!-- #include file="_INCLUDE/Global_Parameters.asp" -->
<!-- #include file="_INCLUDE/Form_validation.asp" -->
<%
' ------------------------------------------------------------
' 
' Name		: __Intranet_Informationl.asp
' Path    	: /
' Description 	: Site Global Parameter
' By		:  Dania Tcherkezoff
' Company	: OverApps
' Date		: December, 10, 2001
' Versions : 1.15.0
'
'
' Modify by	:
' Company	:
' Date		:
' ------------------------------------------------------------

Dim myPage
myPage = "__Intranet_Information.asp"
Dim myPage_Application
myPage_Application="About"







%>

<!-- #include file="_INCLUDE/DB_Environment.asp" -->
<!-- #include file="_INCLUDE/Environment_Tools.asp" -->

<%


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TEST IF USER AS ACCESS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


myApplication_Public_Type_ID = Get_Application_Public_Type_ID(myPage_Application)
if myApplication_Public_type_ID<myUser_type_ID then
	Response.redirect("__Quit.asp")
end if

	
Dim myAuthor_Update, myDate_Update

Dim  mySQL_Select_tb_Sites_Activities, mySet_tb_Sites_Activities
Dim mySQL_Select_tb_Countries, mySet_tb_Countries
Dim myCountry_ID, myCountry

Dim mySQL_Select_tb_Applications 					
Dim mySet_tb_Applications

Dim mySQL_Delete_tb_Sites_Applications 					

Dim i
Dim myMax_Applications

Dim myApplication_ID()
Dim myApplication_Name()



Dim mySite_Application_Title()
Dim mySite_Application_Public_Type_ID()
Dim mySite_Application_Opened()

Dim myApplication_Field_Name


%>
<html>

<head>
<title><%=mySite_Name%> </title>

</head>

<BODY BackGround="<%=myBGImage%>" bgColor="<%=myBGColor%>"  Text="<%=myBGTextColor%>" link="<%=myBGTextColor%>" vLink="<%=myBGTextColor%>"marginwidth="0" marginborder="0" leftmargin="0" topmargin="0" marginheight="0">

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

<TD WIDTH="<%=myLeft_Width%>"> 
<!-- #include file="_borders/Left.asp" --> 
</td>

<%
' CENTER APPLICATION
%> 

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate Form	- Information coming from  DB_Environment.asp		'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>


<td WIDTH="<%=myApplication_Width%>" BGCOLOR="<%=myBGCOLOR%>" valign="top"> 
<form method="POST" action="<%=myPage%>" name="myForm"> 

        <table border="0" Width="<%=myApplication_Width%>" bgcolor="<%=myBGColor%>" cellpadding="5" cellspacing="1">
          <%
' TITLE
%>
          <tr> 
            <td align="center" colspan="2" bgcolor="<%=myApplicationColor%>"> 
              <font color="<%=myApplicationTextColor%>" face="Arial, Helvetica, sans-serif" size="4"><b><%=mySite_Name%></b></font></td>
          </tr>
<%
' URL ADDRESS
%>

<% If len(mySite_URL) >0 Then %>

          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_URL_Address%></FONT></B> 
            </td>
            <td align="left" width="72%"> 
              <a href="<%=mySite_URL%>"><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_URL%></font></a>
              <INPUT TYPE="hidden" NAME="Site_ID" VALUE="<%=mySite_ID%>">
            </td>
          </tr>
<% End If %>

<%
' NAME
%>
          <TR> 
            <TD ALIGN="right" BGCOLOR="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Name%></Font></B> </TD>
            <TD ALIGN="left" width="72%" ><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_Name%>
              </FONT> </TD>
          </TR>
<%
' Presentation
%>


<% If len(mySite_Presentation)> 0 Then %>
          <TR> 
            <TD ALIGN="right" BGCOLOR="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Presentation%></font></b>
	         </TD>
            <TD ALIGN="left" width="72%"><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2">
<%=mySite_Presentation%> 

              </FONT> </TD>
          </TR>
<% End If %>		  
		  
		  
<%
' Company
%>

<% If len(mySite_Company) >0 Then %>


          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Company%></font></b> 
            </td>
            <td align="left" width="72%" ><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_Company%>
              </font> </td>
          </tr>
<% End If %>

<%
' Address
%>


<% If len(mySite_Address) > 0 Then %>

          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Address%></font></b> 
            </td>
            <td align="left" width="72%" ><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_Address%></font>
            </td>
          </tr>
<% End If %>		  
		  
<%
' Zip Code
%>

<% If len(mySite_Zip) > 0 Then %>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"> <b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Zip_Code%></font></b> 
            </td>
            <td align="left" width="72%" ><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_Zip%>
              </font> </td>
          </tr>
<% End If %>

<%
' City
%>

<%If len(mySite_City)> 0 Then%>

          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_City%></FONT></B> 
            </td>
            <td align="left" width="72%"><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_City%>
              </font> </td>
          </tr>
<% End If %>

<%
' State
%>


<%If len(mySite_State) Then%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_State%></FONT></B> 
            </td>
            <td align="left" width="72%" > <font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_State%>
              </font> </td>
          </tr>

<% End If %>

<%
' Country
%>

<%If mySite_Country_ID <> 0 Then %>

          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <b><font face="Arial, Helvetica, sans-serif" size="2" Color="<%=myBorderTextColor%>"><%=myMessage_Country%></font></b> 
            </td>
            <td align="left" width="72%"> 
              <% 
''''''''''''''''''''''''''''''''''
' Get Country			 '	
''''''''''''''''''''''''''''''''''
' DB Connection
set myConnection = Server.CreateObject("ADODB.Connection")
myConnection.Open myConnection_String
mySQL_Select_tb_Countries = "SELECT * FROM tb_Countries order by Country"
set mySet_Tb_Countries = 	myConnection.Execute(mySQL_Select_tb_Countries) %>
            <%if mySite_Country_ID = 0 then%> <%end if %> 
              
                  <%do while not mySet_Tb_Countries.eof
	myCountry_ID = mySet_Tb_Countries("Country_ID")
	myCountry = mySet_Tb_Countries("Country")
%>
                   <%if mySite_Country_ID = myCountry_ID then%><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=myCountry%></font> <%end if%>
                  <%
	mySet_Tb_Countries.MoveNext
	loop
%>
                
              <%
' Close Recordset
mySet_Tb_Countries.close
Set mySet_Tb_Countries = Nothing
' Close Connection
myConnection.Close
set myConnection = Nothing
%>
            </td>
          </tr>

<%end if%>

<%
' Phone
%>

<%If len(mySite_Phone) > 0 Then%>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Phone%></FONT></B> 
            </td>
            <td align="left" width="72%" ><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_Phone%>
              </font> </td>
          </tr>
<%end if%>


<%
' Fax
%>

<%If len(mySite_Fax) > 0 Then %>
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Fax%></FONT></B> 
            </td>
            <td align="left" width="72%"><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_Fax%>
              </font> </td>
          </tr>
<%end if%>



<%
' Other Web
%>

<% If len(mySite_Web) > 0 Then%>
          <tr> 
            <td align="right" bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Other_Web_Site%></FONT></B> 
            </td>
<td align="left" width="72%"> 
            <a href="<%=mySite_Web%>"><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_Web%></font></a>
               </td>
          </tr>
<%end if%>		  

<%
' Email
%>

<% If len(mySite_Email) >0 Then %>
 
          <tr> 
            <td align="right"  bgcolor="<%=myBorderColor%>" width="28%"> <B><FONT FACE="Arial, Helvetica, sans-serif" SIZE="2" Color="<%=myBorderTextColor%>"><%=myMessage_Email%><BR>
              </FONT></B> </td>
            <td align="left" width="72%">
			<a href="mailto:<%=mySite_Email%>"><font color="<%=myBGTextColor%>" face="Arial, Helvetica, sans-serif" size="2"><%=mySite_Email%></font></A>
              </td>
          </tr>
<%end if%>

          
     <tr> 
<td align="center"  colspan="2" bgcolor="<%=myApplicationColor%>">&nbsp; 
<b><font face="Arial, Helvetica, sans-serif" size="1" color="<%=myApplicationTextColor%>"> 
<% = myDate_Display(mySite_Date_Update,2) %>&nbsp;-&nbsp;<% = mySite_Author_Update %></font></b> 
</td>
</tr>
        </table>
</form>
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
%> <TABLE WIDTH="<%=myGlobal_Width%>" BORDER="0" BGCOLOR="<%=myBorderColor%>" CELLPADDING="0" CELLSPACING="0"><TR ALIGN="RIGHT"><TD><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">Copyright (C) 2001-2002  <A HREF="http://www.overapps.com"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1" COLOR="<%=myBorderTextColor%>">OverApps</Font></A> & contributors
</FONT></TD></TR></TABLE><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'				    End Copyright				'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %> 
</body>
</html>

