<%
' ------------------------------------------------------------------------------------------------
' Copyright (C) 2001  Nicolas SOREL (www.ASPFr.com)
'
' This program "Files_Upload_Class.asp" is free software; you can redistribute it 
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
' 	" Copyright (C) 2001 & 2002 OverApps & contributors "
' at the bottom of the page with an active link from the name "OverApps" to 
' the Address http://www.overapps.com where the user (netsurfer) could find the 
' appropriate information about this license. 
'
'-------------------------------------------------------------------------------------------------
%>

<%
' ------------------------------------------------------------------------------------------------
' Name : File_Upload_Class.asp
' Path : /_Include
' Version : 1.18.0
' Description : Tools for uploading files
' By : Nicolas SOREL (www.ASPFr.com)
' Date : December 10 2001
'
' Modify by : Dania Tcherkezoff 	
' Company : Overapps
' Comments :  Done translation to english of variables and function
' -------------------------------------------------------------------------------------------------

Class UplFile

    Private AllSend

    Private VarBinFile
    Private VarFileSize
    Private VarBnFileSize
    
    Private FilesName()
    Private FilesSize()
    Private NbofFiles
    Private Files()
    Private FormsName()
    Private LocalPath
    Private DistantPath()
    Private LocalFileName
    Private TXTFieldName()
    Private TXTFieldsName()
    
    Private Property Let AddTXTField(TheTXT)
        Redim Preserve TXTFieldsName(Ubound(TXTFieldsName) + 1)
        TXTFieldsName(Ubound(TXTFieldsName)) = TheTXT
    End Property

    Private Property Let AddFieldName(TheName)
        Redim Preserve TXTFieldName(Ubound(TXTFieldName) + 1)
        TXTFieldName(Ubound(TXTFieldName)) = TheName
    End Property

    Private Property Let AddFileName(TheName)
        Redim Preserve FilesName(Ubound(FilesName) + 1)
        FilesName(Ubound(FilesName)) = TheName
    End Property

    Private Property Let AddFileSize(TheSize)
        Redim Preserve FilesSize(Ubound(FilesSize) + 1)
        FilesSize(Ubound(FilesSize)) = TheSize
    End Property

    Private Property Let AddDistantPath(TheDistantPath)
        Redim Preserve DistantPath(Ubound(DistantPath) + 1)
        DistantPath(Ubound(DistantPath)) = TheDistantPath
    End Property

    Private Property Let AddFile(TheFile)
        Redim Preserve Files(Ubound(Files) + 1)
        Files(Ubound(Files)) = TheFile
    End Property

    Private Property Let AddFormName(TheFormName)
        Redim Preserve FormsName(Ubound(FormsName) + 1)
        FormsName(Ubound(FormsName)) = TheFormName
    End Property

    Public Property Let Folder(TheFolder)
        LocalPath = TheFolder
    End Property

    Public Property Let NewName(NewFileName)
        LocalFileName = NewFileName
    End Property

    Public Function Save_File(Lequel)
        Dim fso, fs
        If LocalFileName = "" Then
            LocalFileName = FilesName(Lequel)
        End If
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fs = fso.OpenTextFile(LocalPath & LocalFileName, 2, True)
            If Err.Number <> 0 Then Response.Write "Error while writing the file : " & LocalPath & FilesName(Lequel) & vbCrLf & Err.Description & "<br>":LocalFileName = "":Exit Function
            fs.Write Files(LeQuel)
            If Err.Number <> 0 Then Response.Write "Error while writing the file : " & LocalPath & FilesName(Lequel) & vbCrLf & Err.Description & "<br>":LocalFileName = "":Exit Function
        Set fs = Nothing
        Set fso = Nothing
        LocalFileName = ""
    End Function

    Public Property Get Form_Field(Lequel)
        For i = 1 To UBound(TXTFieldName)
            If TXTFieldName(i) = Lequel Then
                Form_Field = TXTFieldsName(i)
                Exit For
            End If
        Next
    End Property

    Public Property Get File_Name(Lequel)
        File_Name = FilesName(Lequel)
    End Property

    Public Property Get DistantFilePath(Lequel)
        DistantFilePath = DistantPath(Lequel)
    End Property

    Public Property Get File_Size(Lequel)
        File_Size = FilesSize(Lequel)
    End Property

    Public Property Get FormName(Lequel)
        FormName = FormsName(Lequel)
    End Property

    Public Property Get Nb_Files()
        Nb_Files = NbofFiles
    End Property

    Private Property Get HttpContentType()
        HttpContentType = Request.ServerVariables ("HTTP_CONTENT_TYPE")
    End Property

    Public Property Get File_Type(Lequel)
       File_Type = Type_Of_File(FilesName(Lequel))
    End Property

    Public Property Get FileExtension(Lequel)
        FileExtension = Right(FilesName(Lequel), Len(FilesName(Lequel)) - InStrRev(FilesName(Lequel),"."))
    End Property

    Private Function Begining()
        VarBinFile = Request.BinaryRead(Request.TotalBytes)
        VarBnFileSize = LenB(VarBinFile)
    End Function

    Private Sub Class_Initialize()
        ReDim FilesName(0)
        ReDim Files(0)
        ReDim FilesSize(0)
        Redim FormsName(0)
        ReDim DistantPath(0)
        Redim TXTFieldsName(0)
        Redim TXTFieldName(0)
        LocalPath = myShared_Files_Path ' Upload Folder Name by default
        LocalFileName = "" 
        Call Begining
        Call LetsGOOOO
    End Sub

    Private Sub Class_Terminate()
        ' Problems with this line
        'Set FilesName = Nothing
        'Set Files = Nothing
        'Set FilesSize = Nothing
    End Sub

    Private Function Upl2ADO()
        Upl2ADO = False
        Dim MyObjRs
        Set MyObjRs = CreateObject("ADODB.Recordset")
            MyObjRs.Fields.Append "TmpBin", 201, VarBnFileSize
            MyObjRs.Open
            MyObjRs.AddNew
            MyObjRs("TmpBin").AppendChunk VarBinFile
            MyObjRs.Update
            AllSend = MyObjRs("TmpBin")
            MyObjRs.Close
        Set MyObjRs = Nothing
        If Err.Number <> 0 Then Response.Write "Error while uploading file(s) : " & vbCrLf & Err.Description & "<br>" : Exit Function
        Upl2ADO = True
    End Function

    Public Function LetsGOOOO()
        Dim Limits, PositionLimit
        Dim FileCount
        Dim LastFileStart, LastFileEnd, FileInUse
        Dim StartFileName, EndFileName, NameOfFile, LastFile
        Dim FileStart, FileEnd, DataOfFile
        Dim TheContentType, SizeofFile, InputName
        Dim IsFile
        
        If Not VarBnFileSize > 0 Then
                 Exit Function
        End If

        If Upl2ADO = True Then
            ' Get HTTP headers
            Limits = HttpContentType

            FileCount = 0

            ' Check for limits (Boundaries)
            PositionLimit = InStr(1, Limits, "boundary=") + 8
            Limits = "--" & Right(Limits, Len(Limits) - PositionLimit)

 
            ' Search first file
            LastFileStart = InStr(1, AllSend, Limits)
            LastFileEnd = InStr(InStr(1, AllSend ,Limits) + 1 , AllSend , Limits) - 1
            LastFile = False

            Do While LastFile = False
                FileInUse = Mid(AllSend, LastFileStart, LastFileEnd - LastFileStart)
                StartFileName = InStr(1, FileInUse, "filename=") + 10
                EndFileName = InStr(StartFileName, FileInUse, Chr(34))
                
                ' Check file is not empty
                If StartFileName <> EndFileName Then
                    FileCount = FileCount + 1
                    ' Get form fields if exist
                    InputName = InStr(1, FileInUse, "name=""")
                    If InputName > 0 Then
                        InputName = Mid(FileInUse, InputName + 6, InStr(InputName + 6, FileInUse, """") - InputName - 6)
                    End If
                    AddFormName = InputName
                    
                    ' Get the distant file path and extract the file name
                    NameOfFile = InStr(1, FileInUse, "filename=""")
                    IsFile = False
                    If NameOfFile > 0 Then
                        IsFile = True
                        NameOfFile = Mid(FileInUse, NameOfFile + 10, InStr(NameOfFile + 10, FileInUse, """") - NameOfFile - 10)
                    End If
                
                    ' Check if input contains a file
                    If IsFile = True Then
                        AddDistantPath = NameOfFile
                        NameOfFile = Right(NameOfFile, Len(NameOfFile) - InStrRev(NameOfFile,"\"))

                        ' Check for start of file just after  Content-Tpye
                        TheContentType = InStr(1, FileInUse, "Content-Type:")
                        If TheContentType > 0 Then
                            FileStart = InStr(TheContentType, FileInUse, vbCrLf) + 4
                        End If
                        FileEnd = Len(FileInUse)

                        ' Calculate size of file
                        SizeofFile = FileEnd - FileStart

                        ' Get file Date
                        DataOfFile = Mid(FileInUse, FileStart, SizeofFile)

                        AddFile = DataOfFile
                        AddFileName = NameOfFile
                        AddFileSize = Len(DataOfFile) 'The Size

                    Else
                        ' Get data from form input : text, textaera, radio button, checkbox etc...
                        FileCount = FileCount - 1
                        FileStart = InStr(InStr(1, FileInUse, "name=""") + 6, FileInUse, """") + 5
                        FileEnd = Len(FileInUse)

                        ' Calculate size of file
                        SizeofFile = FileEnd - FileStart

                        ' Get the File Data
                        DataOfFile = Mid(FileInUse, FileStart, SizeofFile)

                        AddFieldName = InputName
                        AddTXTField = DataOfFile
                    End If
                End If
                ' Go to the next file
                ' Check for start and end of the file
                LastFileStart = InStr(LastFileEnd + 1, AllSend, Limits)
                LastFileEnd = InStr(LastFileStart + 1 , AllSend, Limits) - 1
                If Not LastFileEnd > 0 Then LastFile = True
            Loop
            NbofFiles = FileCount
       End If

    End Function

	'Give Type of File depending on the file extension
	'You can change an add many as you want
	
    Private Function Type_Of_File(File)
        Dim TmpExt
            TmpExt = Right(File, Len(File) - InStrRev(File,"."))
        Select Case LCase(TmpExt)
            Case "jpg", "jpeg"
                Type_Of_File = "JPEG Image"
            Case "gif"
                Type_Of_File = "Gif Image"
            Case "png"
                Type_Of_File = "PNP Image"

            Case "txt"
                Type_Of_File = "Text"
            Case "asp"
                Type_Of_File = "Active Server Page"
            Case "html", "htm"
                Type_Of_File = "Web Document"
            Case "xml"
                Type_Of_File = "XML"
            Case "log"
                Type_Of_File = "Text"

            Case "doc"
                Type_Of_File = "MS Word Document"
            Case "xls"
                Type_Of_File = "MS Excel Document"
            Case "pdf"
                Type_Of_File = "Acrobat Reader Document"

            Case "exe"
                Type_Of_File = "Program"
            Case "zip"
                Type_Of_File = "Zip Archive"
            Case "rar"
                Type_Of_File = "RAR Archive"

            Case "mp3", "mp2"
                Type_Of_File = "Audio"
            Case "au"
                Type_Of_File = "Audio"
            Case "wav"
                Type_Of_File = "Audio"

            Case "mpg", "mpeg"
                Type_Of_File = "Video"
            Case "avi"
                Type_Of_File = "Video"
            'Non Exhaustive list, You can add more as you like

            Case Else
                Type_Of_File = ""
        End Select
    End Function

End Class
%>

