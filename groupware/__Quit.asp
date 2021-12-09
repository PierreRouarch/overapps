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
<% 
' ------------------------------------------------------------
' Name			: __quit.asp
' Path	    	: 
' Description 	: Session To nothing, go back to identification
' by			: Pierre Rouarch
' Company		: OverApps	
' Date			: September, 11, 2001 
' Version		: 1.15.0
' --------------------------------------------------------------

	Session.Abandon
	Response.Redirect("__Identification_Site.asp")

'	Response.Redirect("__Home.asp")
%>
<html>

<head>
<title>End session</title>
</head>

<body>
</body>
</html>

<html><script language="JavaScript"></script></html>
<html><script language="JavaScript"></script></html>