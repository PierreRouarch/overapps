<script language=JavaScript runat=Server> 
	function myValueIsGood (myerr_number, myfield) {
			var myvalidationReg;
			if (myerr_number==3) {
				myvalidationReg = "^[1-9]*(\d)(\.\d{0,2})|(\d)$"
			}
			if (myerr_number==4) {
				myvalidationReg = "^[a-zA-Z0-9_]([\-a-zA-Z0-9_]|\.[a-zA-Z0-9_])*\@[a-zA-Z0-9_]([\-a-zA-Z0-9_]|\.[a-zA-Z0-9_])*$"

			}
			if (myerr_number==5) {
					myvalidationReg = "^\\d{5}$"
			}
			if (myerr_number==6) {
				myvalidationReg = "^[0-9]*([,]|[0-9]){0,1}[0-9]*$"
			}
			if (myerr_number==7) {
				myvalidationReg = "^[A-Za-z 0-9\.\]{1,}$"
			}
			if (myerr_number==8) {
				myvalidationReg = "^[a-z A-Z\.\]+$"
			}
			if (myerr_number==9) {
				myvalidationReg = "^['a-zA-Z חאהגיטךכשüûןמצפÿ-]+$"
			}
			if (myerr_number==10) {
				myvalidationReg = "^[+]{0,1}[0-9 ]+$"
			}
			if (myerr_number==11) {
				myvalidationReg = "^[ ',-.0-9A-Z'a-zחאהגיטךכשüûןמצפÿ]{1,}$"
			}
			if (myerr_number==12) {
				myvalidationReg = "^[0-9_a-z/]{0,}[0-9_a-z]{1,}[.]{1,1}[a-z0-9]{3,3}$"
			}
			if (myerr_number==13) {
				myvalidationReg = "^[A-Za-zחאהגיטךכשüûןמצפÿ]{1,1}[ '-.A-Za-zחאהגיטךכשüûןמצפÿ]{0,}$"
			}
			if (myerr_number==17) {
				myvalidationReg = "^[A-Za-z0-9\.\]{1,}$"
			}
			if (myerr_number==18) {
				myvalidationReg = "^[a-zA-Z\.\]+$"
			}
			if (myerr_number==20) {
				myvalidationReg = "^[A-Za-z0-9\-\_]{1,}$"
			}        
			if (myerr_number==34) {
				myvalidationReg = "^#[0-9a-fA-F]{6}$"
			}    
			
			
			var regex = new RegExp(myvalidationReg);
			return regex.test(myfield);

	}
</script>




<%
' -----------------------------------------------------------------------------
' Copyright (C) 2001  - Over<>Apps - http://www.overapps.com
'
' This program "Form_Validation.asp" is free software; 
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
'-----------------------------------------------------------------------------
%>


<%
' ------------------------------------------------------------
' Name : Form_Validation.asp
' Path : /_Include
' Description : Validate Forms
' By : Jean Luc Lesueur
' Company : OverApps
' Update : February, 13, 2001
'
' Modify by : Pierre Rouarch
' Company : OverApps	
' Date : October 10, 2001
' ------------------------------------------------------------


' Validation Error Constants 
Const myerr_required = 2
Const myerr_currency = 3
Const myerr_email = 4
Const myerr_french_zip_code = 5
Const myerr_numerical = 6
Const myerr_alphanumeric = 7
Const myerr_alpha = 8
Const myerr_name = 9
Const myerr_phone = 10
Const myerr_address = 11
Const myerr_file = 12
Const myerr_city_state = 13
Const myerr_date = 14
Const myerr_alphanumeric_without_blank = 17
Const myerr_alpha_without_blank = 18
Const myerr_directory = 20
Const myerr_under_minimum_value = 30
Const myerr_upper_maximum_value = 31
Const myerr_under_minimum_lenght = 32
Const myerr_upper_maximum_lenght = 33
Const myerr_color_code =34

' Global Variables
Dim myform_entry_error, myform_entries_str

myform_entry_error=False


' Get form values in an array of array called  myform_entries_array
Sub myFormSetEntriesInString
	Dim Item
	For each Item in Request.form
		myform_entries_str=myform_entries_str & "||" & Item & "||0||"
	Next
End Sub



' Validate all values  
' (myerror_number and/or mymandatory and/or value or max lenght and/or  value or minimum lenght)
' Put an error code to each field
' (Global variable myform_entry_error=true and add error in the array de  myform_entries_array)


Sub myFormCheckEntry(myerror_number, myfield, mymandatory, myvalue_min, myvalue_max, mylen_min, mylen_max)
	Dim myerror_number_var, myindex, mystart, myend
	myerror_number_var = 0
	If (mymandatory and (Len(Request.form(myfield)) = 0)) Then
		myerror_number_var = myerr_required
	ElseIf Len(Request.form(myfield)) = 0 Then
		myerror_number_var = 0
	ElseIf myerror_number=14 then
		if Not IsDate(Request.form(myfield)) Then	myerror_number_var = myerror_number
	ElseIf Not(IsNull(myerror_number)) AND Not myValueIsGood (myerror_number, Request.form(myfield)) then
		myerror_number_var = myerror_number
	ElseIf Not(IsNull(myvalue_min)) AND (Request.form(myfield)-myvalue_min) < 0 Then
		myerror_number_var = myerr_under_minimum_value
	ElseIf Not(IsNull(myvalue_max)) And (Request.form(myfield)-myvalue_max) > 0 then
		myerror_number_var = myerr_upper_maximum_value
	ElseIf Not(IsNull(mylen_min)) And (Len(Request.form(myfield)) < mylen_min) then
		myerror_number_var = myerr_under_minimum_lenght
	ElseIf Not(IsNull(mylen_max)) And (Len(request.form(myfield)) > mylen_max) then
		myerror_number_var = myerr_upper_maximum_lenght
	End if

	if myerror_number_var > 0 Then
		mystart = Instr(1,myform_entries_str,"||" & myfield & "||") + Len(myfield) + 4
		myend = Instr(mystart,myform_entries_str,"||")
		myform_entries_str = Left(myform_entries_str,mystart-1) & myerror_number_var & Mid(myform_entries_str,myend)
		myform_entry_error = True
	end if
End Sub




' Get text error for each field or nothing if there is not error

Function myFormGetErrMsg(myfield)
	Dim myerror_number,mystart,myend

	if myform_entry_error Then
		' Get number error for the field
		mystart = Instr(1,myform_entries_str,"||" & myfield & "||") + Len(myfield) + 4
		myend = Instr(mystart,myform_entries_str,"||")
		myerror_number=Mid(myform_entries_str,mystart, myend-mystart)

		select case Int(myerror_number)

			case myerr_name
				myFormGetErrMsg = myError_Message_Not_a_valid_name
			case myerr_address
				myFormGetErrMsg = myError_Message_Not_a_valid_Address
			case myerr_date
				myFormGetErrMsg = myError_Message_not_a_valid_Date
			case myerr_required
				myFormGetErrMsg = myError_Message_Required
			case myerr_alpha
				myFormGetErrMsg = myError_Message_Not_Valid_alpha ' Not valid (caracters and/or blanks)
			case myerr_alpha_without_blank
				myFormGetErrMsg = myError_Message_Not_Valid_alpha_without_blank 'Not Valid (caracters no blank)
			case myerr_alphanumeric
				myFormGetErrMsg = myError_Message_Not_Valid_alphanumeric ' Not valid (caracters and/or digits)
			case myerr_alphanumeric_without_blank
				myFormGetErrMsg = myError_Message_Not_Valid_alphanumeric_without_blank ' Non valide (caracters, digits, no Blank)
			case myerr_Directory
				myFormGetErrMsg = myError_Message_Not_a_Valid_Directory_name
			case myerr_numerical
				myFormGetErrMsg = myError_Message_Not_a_Valid_numerical 'Not valid (digits)
			case myerr_email
				myFormGetErrMsg = myError_Message_Not_a_Valid_Email 
			case myerr_French_zip_code
				myFormGetErrMsg = myError_Message_Not_a_Valid_French_zip_Code 'Not a valid French zip code
			case myerr_phone
				myFormGetErrMsg = myError_Message_Not_Valid_Phone_Number 'Not a valid phone number
			case myerr_city_state
				myFormGetErrMsg = myError_Message_Not_a_Valid_City_State_name 
			case myerr_file
				myFormGetErrMsg = myError_Message_Not_a_Valid_File_Name 
			case myerr_under_minimum_value
				myFormGetErrMsg = myError_Message_Too_Small      
			case myerr_upper_maximum_value
				myFormGetErrMsg = myError_Message_Too_Big
			case myerr_under_minimum_lenght
				myFormGetErrMsg = myError_Message_Not_enough_caracters
			case myerr_upper_maximum_lenght
				myFormGetErrMsg = myError_Message_Too_much_caracters
			case myerr_color_code 	
				myFormGetErrMsg = myError_Message_color_code
			case else
				myFormGetErrMsg =""
		end select

	end if

End Function
%>
