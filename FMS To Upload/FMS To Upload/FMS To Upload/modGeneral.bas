Attribute VB_Name = "modGeneral"
Option Explicit

Public boolConfirmDelete
Public strFileName, dbPathString, strServerPath, strTempServerPath As String
Public strUserName, strDept, strFilePath, strDocParentId, strDocUser As String
Public strDocID As String
Public strSortBy As String
Public strLastFlexiName As String
Public boolUserRoleSuccess As Boolean
Public strOldPass As String
Public strNewPass As String
Public strFormatString As String
Public strUserFileName, strUserDeptName As String
Public lngX, lngY As Long
Public lngMSFx, lngMSFy As Long
Public strFileId As String
Public boolFromFileUpload As Boolean
Public boolPrint As Boolean
Public boolMSFLastLineText As Boolean
Public strDeleteDate, strDeleteTime As String
Public strDocRevID, strDocRevUser, strDocRevParentID As String
Public boolFromRevision As Boolean
Public strFileLocation As String
Public boolRemoveFirst As Boolean
Public strSearchFileName, strSearchUserName As String
Public strNoOfLicenses, strNoOfUsers As String
Public chkLicenseLog As Boolean
Public retVal
Public fsFile

Public DBFileManager As Database
Public rsLogin As Recordset
Public rsUserRole As Recordset
Public rsLicenses As Recordset
Public rsUsers As Recordset
Public rsUserLogins As Recordset
Public rsDeptNames As Recordset
Public rsChkDept As Recordset
Public rsLoadFileActivity As Recordset
Public rsUploadNewFile As Recordset
Public rsSetUserRights As Recordset
Public rsLoadFiles As Recordset
Public rsLoadFilesDocId As Recordset
Public rsFileActivity As Recordset
Public rsDocID As Recordset
Public rsFileProperties As Recordset
Public rsUserFileRights As Recordset
Public rsUserForceLogout As Recordset
Public rsActivityLog As Recordset
Public rsDeleteFiles As Recordset
Public rsDocDetails As Recordset
Public rsDefaultUserRight As Recordset
Public rsPurgeRestore As Recordset
Public rsCheckOut As Recordset
Public rsCheckOutIDs As Recordset
Public rsChecOutDocs As Recordset
Public rsMaxRevNo As Recordset
Public rsPrevDoc As Recordset
Public rsOldRevisions As Recordset
Public rsSearch As Recordset
Public rsLocation As Recordset
Public rsLockDoc As Recordset
Public rsUnLockDoc As Recordset

Sub validateLoginPassword()
Set rsLogin = DBFileManager.OpenRecordset("select login,department from master_users where login='" & Replace(frmLogin.txtLoginName.Text, "'", "''") & "' and password='" & Replace(frmLogin.txtPassword.Text, "'", "''") & "'")
If rsLogin.EOF = True Then
    MsgBox "Invalid password !!!", vbExclamation
    frmLogin.txtPassword.SetFocus
    frmLogin.txtPassword.SelStart = 0
    frmLogin.txtPassword.SelLength = Len(frmLogin.txtPassword.Text)
ElseIf rsLogin.EOF = False Then
    If frmLogin.chkRememberUN.Value = 1 Then
        SaveSetting App.EXEName, "Auto", "Auto User Name", "True"
        SaveSetting App.EXEName, "Auto", "User Name", frmLogin.txtLoginName.Text
    Else
        SaveSetting App.EXEName, "Auto", "Auto User Name", "False"
        On Error Resume Next
        DeleteSetting "Inventor", "Auto", "User Name"
    End If
    If frmLogin.chkRememberPass.Value = 1 Then
        SaveSetting App.EXEName, "Auto", "Auto Password", "True"
        SaveSetting App.EXEName, "Auto", "Password", frmLogin.txtPassword.Text
    Else
        SaveSetting App.EXEName, "Auto", "Auto Password", "False"
        On Error Resume Next
        DeleteSetting App.EXEName, "Auto", "Password"
    End If
    
    rsLogin.MoveFirst
    strUserName = StrConv(frmLogin.txtLoginName.Text, vbProperCase)
    strDept = StrConv(rsLogin!Department, vbProperCase)
           
    If UCase(strUserName) = "ADMIN" Then
        frmMain.mnuAdmin.Visible = True
        Unload frmLogin
        frmMain.Show
    Else
        Set rsUserLogins = DBFileManager.OpenRecordset("select * from logins where Login='" & strUserName & "'")
        If rsUserLogins.EOF = True Then
            If chkNoOfLicenses = True Then
                rsUserLogins.AddNew
                    rsUserLogins!Login = strUserName
                    rsUserLogins!Department = strDept
                rsUserLogins.Update
                Unload frmLogin
                frmMain.Show
            Else
                Unload frmLogin
                Exit Sub
            End If
        Else
            MsgBox "You have already logged in from another system. Please logout to re-login.", vbExclamation
            Exit Sub
        End If
            
        frmMain.mnuAdmin.Visible = False
    End If
End If
End Sub
    
Sub changeUserRoleData()
With frmUserRoles
    Set rsUserRole = DBFileManager.OpenRecordset("select Login,Password,Department from Master_Users where Login='" & Replace(.cmbUserName.Text, "'", "''") & "'")
    If rsUserRole.EOF = False Then
        rsUserRole.MoveFirst
        .txtOldPassword.Text = rsUserRole!Password
        .cboDepartment.Text = rsUserRole!Department
        
        .cmdAdd.Enabled = False
        .cmdDelete.Enabled = True
        .cmdUpdate.Enabled = True
        
        .mnuAddUser.Enabled = False
        .mnuDeleteUser.Enabled = True
        .mnuUpdateUser.Enabled = True
    ElseIf rsUserRole.EOF = True Then
        .cmdAdd.Enabled = True
        .cmdDelete.Enabled = False
        .cmdUpdate.Enabled = False
        
        .mnuAddUser.Enabled = True
        .mnuDeleteUser.Enabled = False
        .mnuUpdateUser.Enabled = False
        
        .txtConfirmPassword.Text = ""
        .txtNewPassword.Text = ""
        .txtOldPassword.Text = ""
        .cboDepartment.Text = .cboDepartment.List(0)
    End If
End With
End Sub

Public Function userRoleValidateControls() As Boolean
userRoleValidateControls = True
With frmUserRoles
    If .txtConfirmPassword.Text <> .txtNewPassword.Text Then
        MsgBox "The New Password and Cofirmation Password does not match !!!", vbExclamation
        userRoleValidateControls = False
        .txtConfirmPassword.SetFocus
        Exit Function
    End If
    
    If Len(Trim(.cmbUserName.Text)) = 0 Then
        MsgBox "Enter a Valid user Name to add !!!", vbExclamation
        userRoleValidateControls = False
        .cmbUserName.SetFocus
        Exit Function
    End If
    
    If Len(Trim(.cmbUserName.Text)) > 100 Then
        MsgBox "User Name should be less than 100 characters!!!", vbExclamation
        userRoleValidateControls = False
        .cmbUserName.SetFocus
        Exit Function
    End If
    
End With
End Function

Sub addUserRoleData()
With frmUserRoles
    rsUserRole.AddNew
    rsUserRole!Login = .cmbUserName.Text
    rsUserRole!Password = .txtNewPassword.Text
    rsUserRole!Department = .cboDepartment.Text
    rsUserRole.Update
    
    .txtOldPassword.Text = .txtNewPassword.Text
    .txtNewPassword.Text = ""
    .txtConfirmPassword.Text = ""
    .StatusBar1.Panels(1).Text = "User Added !!!"
End With
End Sub

Sub loadUserRoleComboBoxes()
With frmUserRoles
    .cboDepartment.Clear
    Set rsUserRole = DBFileManager.OpenRecordset("select Department from Master_Departments where Department<>'admin'")
    If rsUserRole.EOF = False Then
        rsUserRole.MoveFirst
        While rsUserRole.EOF = False
            .cboDepartment.AddItem rsUserRole(0)
            rsUserRole.MoveNext
        Wend
    End If
    .cboDepartment.AddItem "Admin"
    .cboDepartment.Text = .cboDepartment.List(0)
    
    .cmbUserName.Clear
    Set rsUserRole = DBFileManager.OpenRecordset("select Login from Master_Users")
    If rsUserRole.EOF = False Then
        rsUserRole.MoveFirst
        While rsUserRole.EOF = False
            .cmbUserName.AddItem rsUserRole(0)
            rsUserRole.MoveNext
        Wend
    End If
    .cmbUserName.Text = .cmbUserName.List(0)
End With
End Sub

Sub deleteUser()
boolConfirmDelete = MsgBox("Do you want to delete this user ?", vbYesNo + vbDefaultButton2 + vbQuestion)
If boolConfirmDelete = vbYes And UCase(frmUserRoles.cmbUserName.Text) <> "ADMIN" Then
    rsUserRole.Delete
    frmUserRoles.StatusBar1.Panels(1).Text = "User Deleted !!!"
    Call loadUserRoleComboBoxes
ElseIf UCase(frmUserRoles.cmbUserName.Text) = "ADMIN" Then
    MsgBox "Cannot delete the user Admin !!!", vbExclamation
End If
End Sub

Sub updateUserRoleData()
With frmUserRoles
    If Len(Trim(.txtConfirmPassword.Text)) = 0 Then
        rsUserRole.Edit
            If UCase(.cmbUserName.Text) <> "ADMIN" Then
                rsUserRole!Department = .cboDepartment.Text
            End If
        rsUserRole.Update
        .StatusBar1.Panels(1).Text = "User Updated!!!"
    ElseIf Len(Trim(.txtConfirmPassword.Text)) > 0 Then
        rsUserRole.Edit
            rsUserRole!Password = .txtNewPassword.Text
            If UCase(.cmbUserName.Text) <> "ADMIN" Then
                rsUserRole!Department = .cboDepartment.Text
            End If
        rsUserRole.Update
        .txtOldPassword.Text = .txtNewPassword.Text
        .StatusBar1.Panels(1).Text = "User Updated!!!"
        .txtNewPassword.Text = ""
        .txtConfirmPassword.Text = ""
    End If
End With
End Sub

Sub addDeptName()
With frmDeptNames
    Set rsDeptNames = DBFileManager.OpenRecordset("select Department from Master_Departments where Department='" & Trim(Replace(.txtDept.Text, "'", "''")) & "'")
    If rsDeptNames.EOF = True Then
        rsDeptNames.AddNew
            rsDeptNames(0) = Trim(.txtDept.Text)
        rsDeptNames.Update
        .txtDept.Text = ""
        .txtDept.SetFocus
        .StatusBar1.Panels(1).Text = "Department Added!!!"
        Call loadDeptsList
    ElseIf rsDeptNames.EOF = False Then
        MsgBox "This department name already exist. Choose a new department name to update.", vbInformation
        .txtDept.SetFocus
    End If
End With
End Sub

Sub delDeptName()
With frmDeptNames
    boolConfirmDelete = MsgBox("Do you want to delete this department name ?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If boolConfirmDelete = vbYes Then
        Set rsDeptNames = DBFileManager.OpenRecordset("select * from Master_Departments where Department='" & Replace(.lstDepts.Text, "'", "''") & "'")
        If rsDeptNames.EOF = False Then
            Set rsChkDept = DBFileManager.OpenRecordset("Select * from Master_Users where Department='" & Replace(.lstDepts.Text, "'", "''") & "'")
            If rsChkDept.EOF = True Then
                rsDeptNames.MoveFirst
                rsDeptNames.Delete
                .StatusBar1.Panels(1).Text = "Department Deleted !!!"
                .txtDept.Text = ""
                Call loadDeptsList
                .lstDepts.SetFocus
                On Error Resume Next
                .lstDepts.Selected(0) = True
            Else
                MsgBox "There are active users in this department. Please delete or re-allocate the users before deleting the department.", vbExclamation
            End If
        ElseIf rsDeptNames.EOF = True Then
            MsgBox "Select a valid deparment to delete.", vbExclamation
            .txtDept.SetFocus
        End If
    End If
End With
End Sub

Sub loadDeptsList()
With frmDeptNames
    .lstDepts.Clear
    Set rsDeptNames = DBFileManager.OpenRecordset("select Department from Master_Departments")
    If rsDeptNames.EOF = False Then
        rsDeptNames.MoveFirst
        While rsDeptNames.EOF = False
            .lstDepts.AddItem rsDeptNames!Department
            rsDeptNames.MoveNext
        Wend
    End If
End With
End Sub

Sub updateDeptName()
With frmDeptNames
    Set rsDeptNames = DBFileManager.OpenRecordset("select Department from Master_Departments where Department='" & Trim(Replace(.txtDept.Text, "'", "''")) & "'")
    If rsDeptNames.EOF = True Then
        Set rsDeptNames = DBFileManager.OpenRecordset("select Department from Master_Departments where Department='" & Replace(.lstDepts.Text, "'", "''") & "'")
        If rsDeptNames.EOF = False Then
            rsDeptNames.MoveFirst
            rsDeptNames.Edit
            rsDeptNames(0) = Trim(.txtDept.Text)
            rsDeptNames.Update
            
            'update Master Users table also
            Set rsDeptNames = DBFileManager.OpenRecordset("select Department from Master_Users where Department='" & Replace(.lstDepts.Text, "'", "''") & "' and Login<>'Admin'")
            If rsDeptNames.EOF = False Then
                rsDeptNames.MoveFirst
                While rsDeptNames.EOF = False
                    rsDeptNames.Edit
                    rsDeptNames(0) = Trim(.txtDept.Text)
                    rsDeptNames.Update
                    rsDeptNames.MoveNext
                Wend
            End If
            
            Call loadDeptsList
            On Error Resume Next
            .lstDepts.Selected(0) = True
            
            .StatusBar1.Panels(1).Text = "Department Updated!!!"
        Else
            MsgBox "Select a valid department to update.", vbInformation
            .txtDept.SetFocus
        End If
    ElseIf rsDeptNames.EOF = False Then
        MsgBox "This department name already exist. Choose a new department name to update to.", vbInformation
        .txtDept.SetFocus
    End If
End With
End Sub

Sub setNewUserPassword()
strOldPass = "": strNewPass = ""

strOldPass = InputBox("Please enter your current password to continue.", "Enter your current password.")
Set rsUsers = DBFileManager.OpenRecordset("select password from Master_Users where Login='" & strUserName & "'")
If rsUsers.EOF = False Then
    rsUsers.MoveFirst
End If
If UCase(strOldPass) = UCase(rsUsers!Password) Then
    strNewPass = InputBox("Please enter your new password to.", "Enter your new password.")
    If Len(strNewPass) > 50 Then
        MsgBox "Your password cannot be more than 50 characters. Please retry", vbExclamation
    ElseIf Len(strNewPass) <= 0 Then
        MsgBox "Enter a valid new password and retry", vbExclamation
    Else
        rsUsers.Edit
        rsUsers!Password = strNewPass
        rsUsers.Update
        MsgBox "Password changed successfully", vbInformation
    End If
Else
    MsgBox "The current password entered by you is incorrect. Please re-try.", vbExclamation
End If
End Sub

Sub uploadNewFile()
With frmFileProperties
    Set rsUploadNewFile = DBFileManager.OpenRecordset("Select Revision,Name,Path,User,Department,[Created Date],[Created Time],Description,Status,Priority,Expiry,[Parent Id] from Documents where name='" & Replace(.lblFileName.Caption, "'", "''") & "' and user='" & strUserName & "' ")
    If rsUploadNewFile.EOF = True Then
        rsUploadNewFile.AddNew
            rsUploadNewFile!Revision = CLng(.lblRevision.Caption)
            rsUploadNewFile!Name = .lblFileName.Caption
            rsUploadNewFile!Path = strServerPath & strUserName & "\"
            rsUploadNewFile!User = strUserName
            rsUploadNewFile!Department = strDept
            rsUploadNewFile![Created Date] = Format(nowDate, "dd-MMM-yyyy")
            rsUploadNewFile![Created Time] = Format(nowTime, "hh:mm AMPM")
            rsUploadNewFile!Description = .txtDescription.Text
            rsUploadNewFile!Status = "None"
            rsUploadNewFile!Priority = .cboPriority.Text
            rsUploadNewFile!expiry = Format(.DTPInactive, "dd-MMM-yyyy")
'            'set parent id for revision
            If boolFromRevision = True Then
                rsUploadNewFile![Parent Id] = strDocRevParentID
            End If
        rsUploadNewFile.Update
            
        'Get the Document ID for added file
        Set rsUploadNewFile = DBFileManager.OpenRecordset("Select [Document Id],[Parent Id]  from Documents where name='" & Replace(.lblFileName.Caption, "'", "''") & "' and user='" & Replace(strUserName, "'", "''") & "' ")
        If rsUploadNewFile.EOF = False Then
            rsUploadNewFile.MoveFirst
            strFileId = rsUploadNewFile![Document Id]
            
            rsUploadNewFile.Edit
            If boolFromRevision = False Then
                rsUploadNewFile![Parent Id] = rsUploadNewFile![Document Id]
            End If
            rsUploadNewFile.Update
        End If
        
        'Update user rights
        DBFileManager.Execute ("Delete * from [Documents FilesUsers] where [Document Id]=" & strFileId)
        Dim intX As Integer
        For intX = 1 To .MSFActiveUsers.Rows - 2
            DBFileManager.Execute ("insert into [Documents FilesUsers] ([Document Id],[User Name],View,Print,Edit,Copy) values ('" & _
                strFileId & "','" & .MSFActiveUsers.TextMatrix(intX, 1) & "','" & .MSFActiveUsers.TextMatrix(intX, 2) & "','" & _
                .MSFActiveUsers.TextMatrix(intX, 3) & "','" & .MSFActiveUsers.TextMatrix(intX, 4) & "','" & _
                .MSFActiveUsers.TextMatrix(intX, 5) & "')")
        Next
        
        'Copy File to Server
        If boolFromRevision = False Then strFilePath = Replace(frmMain.CommonDialog1.FileName, frmMain.CommonDialog1.FileTitle, "", , , vbBinaryCompare)
        strTempServerPath = strServerPath & strUserName
        If chkFolderPath(strTempServerPath, True) = True Then
            Call CopyFile(strFilePath & .lblFileName.Caption, strTempServerPath & "\" & .lblFileName.Caption, 1)
        End If
        
        'Deleting the checkout details and updating the expiry to day before of existing old file
        If boolFromRevision = True Then
            DBFileManager.Execute ("Delete * from [Documents CheckOut] where [Document Id]=" & strDocID)
            DBFileManager.Execute ("update documents set expiry=dateadd('d',-1,now()) where  [document id]=" & strDocRevID & " or [parent id]=(select [parent id] from documents where [document id]=" & strDocRevID & ")" & _
            " and revision <>(select max(revision) from documents where [document id]=" & strDocRevID & " or [parent id]=(select [parent id] from documents where [document id]=" & strDocRevID & ")) " & _
            " and expiry>=now()")

            'DBFileManager.Execute ("Update Documents Set Expiry=#" & Format(DateAdd("d", -1, Format(nowDate, "dd-MMM-yyyy")), "dd-MMM-yyyy") & "# where [Document Id]=" & strDocRevID)
        End If
        
        Unload frmFileProperties
    Else
        MsgBox "You already have a document named " & .lblFileName.Caption & ". Please rename the file and upload.", vbExclamation
    End If
End With
End Sub

Sub uploadFileProperties()
With frmFileProperties
    Set rsUploadNewFile = DBFileManager.OpenRecordset("Select Revision,Name,Path,User,Department,[Created Date],[Created Time],Description,Status,Priority,Expiry from Documents where [Document Id]=" & strDocID)
    If rsUploadNewFile.EOF = False Then
        rsUploadNewFile.Edit
            rsUploadNewFile!Description = .txtDescription.Text
            rsUploadNewFile!Priority = .cboPriority.Text
            rsUploadNewFile!expiry = Format(.DTPInactive, "dd-MMM-yyyy")
        rsUploadNewFile.Update
        
        DBFileManager.Execute ("Delete * from [Documents FilesUsers] where [Document Id]=" & strDocID)
        Dim intX As Integer
        For intX = 1 To .MSFActiveUsers.Rows - 2
            DBFileManager.Execute ("insert into [Documents FilesUsers] ([Document Id],[User Name],View,Print,Edit,Copy) values ('" & _
                strDocID & "','" & .MSFActiveUsers.TextMatrix(intX, 1) & "','" & .MSFActiveUsers.TextMatrix(intX, 2) & "','" & _
                .MSFActiveUsers.TextMatrix(intX, 3) & "','" & .MSFActiveUsers.TextMatrix(intX, 4) & "','" & _
                .MSFActiveUsers.TextMatrix(intX, 5) & "')")
        Next
        Unload frmFileProperties
    Else
        MsgBox "You cannot update the properties of this file now.", vbExclamation
    End If
End With
End Sub

Sub loadOtherUsersFiles()
With frmMain
strFormatString = .MSFFiles.FormatString
.MSFFiles.Clear
.MSFFiles.FormatString = strFormatString
lngX = 0
Set rsLoadFilesDocId = DBFileManager.OpenRecordset("select [Document Id] from [Documents FilesUsers] where [User Name]='" & strUserName & "'")
If rsLoadFilesDocId.EOF = False Then
    rsLoadFilesDocId.MoveFirst
    While rsLoadFilesDocId.EOF = False
        Set rsLoadFiles = DBFileManager.OpenRecordset("select '' as Dummy,Name,Description,Revision,User,Department,[Created Date],Expiry,Priority from Documents where [Document Id]=" & rsLoadFilesDocId![Document Id] & " and Expiry >= #" & Format(Now, "dd-MMM-yyyy") & "# and [User] <> '" & strUserName & "' order by " & strSortBy)
        If rsLoadFiles.EOF = False Then
            rsLoadFiles.MoveFirst
            While rsLoadFiles.EOF = False
                .MSFFiles.Row = .MSFFiles.Rows - 1
                While lngX < .MSFFiles.Cols
                    .MSFFiles.Col = lngX
                    If UCase(rsLoadFiles!Priority) = "HIGH" Then
                        .MSFFiles.CellForeColor = &H800000
                    ElseIf UCase(rsLoadFiles!Priority) = "LOW" Then
                        .MSFFiles.CellForeColor = &H808080
                    Else
                        .MSFFiles.CellForeColor = vbBlack
                    End If
                    If IsDate(rsLoadFiles(lngX)) = True Then
                        .MSFFiles.Text = Format(rsLoadFiles(lngX), "dd-MMM-yyyy")
                    Else
                        .MSFFiles.Text = rsLoadFiles(lngX)
                    End If
                    lngX = lngX + 1
                Wend
                lngX = 0
                .MSFFiles.Rows = .MSFFiles.Rows + 1
                rsLoadFiles.MoveNext
            Wend
        End If
        rsLoadFilesDocId.MoveNext
    Wend
End If
End With
End Sub

Sub loadInactiveFiles()
With frmMain
strFormatString = .MSFFiles.FormatString
.MSFFiles.Clear
.MSFFiles.FormatString = strFormatString
lngX = 0
Set rsLoadFilesDocId = DBFileManager.OpenRecordset("select [Document Id] from [Documents FilesUsers] where [User Name]='" & strUserName & "'")
If rsLoadFilesDocId.EOF = False Then
    rsLoadFilesDocId.MoveFirst
    While rsLoadFilesDocId.EOF = False
        Set rsLoadFiles = DBFileManager.OpenRecordset("select '' as Dummy,Name,Description,Revision,User,Department,[Created Date],[Modified By],[Modified Date],Expiry,Priority  from Documents where [Document Id]=" & rsLoadFilesDocId![Document Id] & " and Expiry < #" & Format(nowDate, "dd-MMM-yyyy") & "# order by " & strSortBy)
        If rsLoadFiles.EOF = False Then
            rsLoadFiles.MoveFirst
            While rsLoadFiles.EOF = False
                .MSFFiles.Row = .MSFFiles.Rows - 1
                While lngX < .MSFFiles.Cols
                    .MSFFiles.Col = lngX
                    If UCase(rsLoadFiles!Priority) = "HIGH" Then
                        .MSFFiles.CellForeColor = &H800000
                    ElseIf UCase(rsLoadFiles!Priority) = "LOW" Then
                        .MSFFiles.CellForeColor = &H808080
                    Else
                        .MSFFiles.CellForeColor = vbBlack
                    End If
                    If IsDate(rsLoadFiles(lngX)) = True Then
                        .MSFFiles.Text = Format(rsLoadFiles(lngX), "dd-MMM-yyyy")
                    ElseIf Len(rsLoadFiles(lngX)) > 0 Then
                        .MSFFiles.Text = rsLoadFiles(lngX)
                    End If
                    lngX = lngX + 1
                Wend
                lngX = 0
                .MSFFiles.Rows = .MSFFiles.Rows + 1
                rsLoadFiles.MoveNext
            Wend
        End If
        rsLoadFilesDocId.MoveNext
    Wend
End If

lngX = 0
Set rsLoadFiles = DBFileManager.OpenRecordset("select '' as Dummy,Name,Description,Revision,User,Department,[Created Date],[Modified By],[Modified Date],Expiry,Priority from Documents where [User]='" & strUserName & "' and Expiry < #" & Format(nowDate, "dd-MMM-yyyy") & "# order by [Created Date],[Created Time] desc")
If rsLoadFiles.EOF = False Then
    rsLoadFiles.MoveFirst
    While rsLoadFiles.EOF = False
        .MSFFiles.Row = .MSFFiles.Rows - 1
        While lngX < .MSFFiles.Cols
            .MSFFiles.Col = lngX
            If UCase(rsLoadFiles!Priority) = "HIGH" Then
                .MSFFiles.CellForeColor = &H800000
            ElseIf UCase(rsLoadFiles!Priority) = "LOW" Then
                .MSFFiles.CellForeColor = &H808080
            Else
                .MSFFiles.CellForeColor = vbBlack
            End If
            If IsDate(rsLoadFiles(lngX)) = True Then
                .MSFFiles.Text = Format(rsLoadFiles(lngX), "dd-MMM-yyyy")
            ElseIf Len(rsLoadFiles(lngX)) > 0 Then
                If Len(rsLoadFiles(lngX)) > 0 Then
                    .MSFFiles.Text = rsLoadFiles(lngX)
                End If
            End If
            lngX = lngX + 1
        Wend
        lngX = 0
        .MSFFiles.Rows = .MSFFiles.Rows + 1
        rsLoadFiles.MoveNext
    Wend
End If

End With
End Sub

Sub loadNewFiles()
With frmMain
strFormatString = .MSFFiles.FormatString
.MSFFiles.Clear
.MSFFiles.FormatString = strFormatString
lngX = 1
Dim strSubQry1 As String, strSubQry2 As String
strSubQry1 = "select [Document Id] from [Documents FilesUsers] where [User Name]='" & strUserName & "'"
strSubQry2 = "Select [Document ID] from [Documents ActivityLog] where [User]='" & strUserName & "'"
Set rsLoadFiles = DBFileManager.OpenRecordset("select [Document Id] ,Name,Description,Revision,User,Department,[Created Date],Expiry,Priority " & _
" from Documents where [Document Id] in (" & strSubQry1 & ") and [Document ID] not in (" & strSubQry2 & ") " & _
" and Expiry >= #" & Format(Now, "dd-MMM-yyyy") & "# order by " & strSortBy)
If rsLoadFiles.EOF = False Then
    rsLoadFiles.MoveFirst
    Set rsFileActivity = DBFileManager.OpenRecordset("Select * from [Documents ActivityLog] where [Document Id]=" & rsLoadFiles![Document Id] & " and [User]='" & strUserName & "'")
    If rsFileActivity.EOF = True Then
        While rsLoadFiles.EOF = False
            .MSFFiles.Row = .MSFFiles.Rows - 1
            While lngX < .MSFFiles.Cols
                .MSFFiles.Col = lngX
                If UCase(rsLoadFiles!Priority) = "HIGH" Then
                    .MSFFiles.CellForeColor = &H800000
                ElseIf UCase(rsLoadFiles!Priority) = "LOW" Then
                    .MSFFiles.CellForeColor = &H808080
                Else
                    .MSFFiles.CellForeColor = vbBlack
                End If
                
                If IsDate(rsLoadFiles(lngX)) = True Then
                    .MSFFiles.Text = Format(rsLoadFiles(lngX), "dd-MMM-yyyy")
                Else
                    .MSFFiles.Text = rsLoadFiles(lngX)
                End If
                lngX = lngX + 1
            Wend
            lngX = 1
            .MSFFiles.Rows = .MSFFiles.Rows + 1
            rsLoadFiles.MoveNext
        Wend
    End If
End If
End With
End Sub

Sub setMenusButtons(mnuName As Menu, cmdButton As OptionButton)
With frmMain
    .MousePointer = vbHourglass
    
    .mnuMyFiles.Checked = False
    .mnuMyInactiveFiles.Checked = False
    .mnuOthersFiles.Checked = False
    .mnuNewFiles.Checked = False
    
    .cmdNewFiles.BackColor = &HC0FFFF
    .cmdOthersFiles.BackColor = &HC0FFFF
    .cmdMyInactiveFiles.BackColor = &HC0FFFF
    .cmdMyFiles.BackColor = &HC0FFFF
    
    mnuName.Checked = True
    cmdButton.BackColor = vbWhite
    
    .MSFFiles.Redraw = False
    .MSFFiles.Clear
End With
End Sub

Sub setDocumentID()
With frmMain
strDocID = ""
strDocUser = ""
If Len(.MSFFiles.TextMatrix(.MSFFiles.Row, 1)) > 0 And .MSFFiles.Row > 0 Then
    strDocID = .MSFFiles.TextMatrix(.MSFFiles.Row, 1)
    If .cmdMyFiles.Value = True Then
        strDocUser = strUserName
        Set rsDocID = DBFileManager.OpenRecordset("Select [Document Id],[Parent Id] from Documents where Name='" & strDocID & "' and User='" & strDocUser & "'")
        If rsDocID.EOF = False Then
            rsDocID.MoveFirst
            strDocID = rsDocID![Document Id]
            If Len(rsDocID![Parent Id]) > 0 Then
                strDocParentId = rsDocID![Parent Id]
            Else
                strDocParentId = strDocID
            End If
        Else
            strDocID = ""
            strDocParentId = ""
        End If
    Else
        strDocUser = .MSFFiles.TextMatrix(.MSFFiles.Row, 4)
        Set rsDocID = DBFileManager.OpenRecordset("Select [Document Id],[Parent Id] from Documents where Name='" & strDocID & "' and User='" & strDocUser & "'")
        If rsDocID.EOF = False Then
            rsDocID.MoveFirst
            strDocID = rsDocID![Document Id]
            If Len(rsDocID![Parent Id]) > 0 Then
                strDocParentId = rsDocID![Parent Id]
            Else
                strDocParentId = strDocID
            End If
        Else
            strDocID = ""
            strDocParentId = ""
        End If
    End If
End If
End With
End Sub

Sub setUserDocRights()
With frmMain
'.cmdDelete.Enabled = False
'.cmdCopyOut.Enabled = False
'.cmdDownload.Enabled = False
'.cmdOpen.Enabled = False
'.cmdDownload.Enabled = False

If Len(strDocID) > 0 Then
    If .lblFiles.Caption = "My Files" Then
        .cmdDelete.Enabled = True
        .cmdCopyOut.Enabled = True
        .cmdDownload.Enabled = True
        .cmdOpen.Enabled = True
        .cmdDownload.Enabled = True
        .cmdLock.Enabled = True
        GoTo ExitSub
    ElseIf .lblFiles.Caption = "My Inactive Files" Then
        If UCase(.MSFFiles.TextMatrix(.MSFFiles.Row, 4)) = UCase(strUserName) Then
            .cmdDelete.Enabled = True
            .cmdCopyOut.Enabled = True
            .cmdDownload.Enabled = True
            .cmdOpen.Enabled = True
            .cmdDownload.Enabled = True
            .cmdLock.Enabled = True
            GoTo ExitSub
        End If
    End If
    
    .cmdDelete.Enabled = False
    .cmdLock.Enabled = False
    Set rsUserFileRights = DBFileManager.OpenRecordset("Select View,Print,Copy,Edit from [Documents FilesUsers] where [Document Id]=" & strDocID & " and [User Name]='" & strUserName & "'")
    If rsUserFileRights.EOF = False Then
       rsUserFileRights.MoveFirst
       
        If UCase(rsUserFileRights!View) = "Y" Then
            .cmdOpen.Enabled = True
        Else
            .cmdOpen.Enabled = False
        End If
        
        If UCase(rsUserFileRights!Copy) = "Y" Then
            .cmdCopyOut.Enabled = True
        Else
            .cmdCopyOut.Enabled = False
        End If
    
        If UCase(rsUserFileRights!Edit) = "Y" Then
            .cmdDownload.Enabled = True
        Else
            .cmdDownload.Enabled = False
        End If
    Else
        .cmdDelete.Enabled = False
        .cmdCopyOut.Enabled = False
        .cmdDownload.Enabled = False
        .cmdOpen.Enabled = False
        .cmdDownload.Enabled = False
        .cmdLock.Enabled = False
        .cmdLock.Caption = "&Lock File"
    End If
Else
    .cmdDelete.Enabled = False
    .cmdCopyOut.Enabled = False
    .cmdDownload.Enabled = False
    .cmdOpen.Enabled = False
    .cmdDownload.Enabled = False
    .cmdLock.Enabled = False
    .cmdLock.Caption = "&Lock File"
End If

ExitSub:
If strDocID = "" Then Exit Sub
Dim rsChcout As Recordset
Set rsChcout = DBFileManager.OpenRecordset("Select User from [Documents CheckOut] where [Document Id]=" & strDocID)
.cmdDownload.Caption = "&Download for Editing"
If rsChcout.EOF = False Then
    If Len(.MSFFiles.TextMatrix(.MSFFiles.Row, 0)) > 0 And LCase(rsChcout(0)) = LCase(strUserName) Then
        .cmdDownload.Caption = "&Upload after Editing"
        .cmdDownload.Enabled = True
    Else
        If Len(.MSFFiles.TextMatrix(.MSFFiles.Row, 0)) > 0 Then
        .cmdDownload.Enabled = False
        End If
    End If
Else
    If Len(.MSFFiles.TextMatrix(.MSFFiles.Row, 0)) > 0 Then
        .cmdDownload.Enabled = False
    End If
End If
.mnuDownload.Caption = .cmdDownload.Caption: .mnuDownload.Enabled = .cmdDownload.Enabled

End With
End Sub

Sub setFileProperties()
With frmFileProperties
    Set rsFileProperties = DBFileManager.OpenRecordset("Select Name,Revision,User,[Created Date],[Created Time],[Modified By],[Modified Date],[Modified Time],Description,Priority,Expiry from Documents where [Document Id]=" & strDocID)
    .lblFileName.Caption = ""
    .lblRevision.Caption = ""
    .lblCreatedBy.Caption = ""
    .lblCreatedOn.Caption = ""
    .lblModifiedBy.Caption = ""
    .lblModifiedOn.Caption = ""
    .txtDescription.Text = ""
    .cboPriority.Text = .cboPriority.List(1)
    .DTPInactive.Value = DateAdd("d", -1, Now())
    
    If rsFileProperties.EOF = False Then
        rsFileProperties.MoveFirst
        
        If Len(rsFileProperties!Name) > 0 Then
            .lblFileName.Caption = rsFileProperties!Name
        End If
        
        If Len(rsFileProperties!Revision) > 0 Then
            .lblRevision.Caption = rsFileProperties!Revision
        End If
        
        If Len(rsFileProperties!User) > 0 Then
            .lblCreatedBy.Caption = rsFileProperties!User
        End If
        
        If Len(rsFileProperties![Created Date]) > 0 Then
            .lblCreatedOn.Caption = Format(rsFileProperties![Created Date], "dd-MMM-yyyy") & "  " & Format(rsFileProperties![Created Time], "hh:mm AMPM")
        End If
        
        If Len(rsFileProperties![Modified By]) > 0 Then
            .lblModifiedBy.Caption = rsFileProperties![Modified By]
        End If
        
        If Len(rsFileProperties![Modified Date]) > 0 Then
            .lblModifiedOn.Caption = Format(rsFileProperties![Modified Date], "dd-MMM-yyyy") & "  " & Format(rsFileProperties![Modified Time], "hh:mm AMPM")
        End If
        
        If Len(rsFileProperties!Description) > 0 Then
            .txtDescription.Text = rsFileProperties!Description
        End If
        
        If Len(rsFileProperties!Priority) > 0 Then
            .cboPriority.Text = rsFileProperties!Priority
        End If
        
        If Len(rsFileProperties!expiry) > 0 Then
            .DTPInactive.Value = Format(rsFileProperties!expiry, "dd-MMM-yyyy")
        End If
    End If
    
    .txtDescription.Locked = (UCase(strDocUser) <> UCase(strUserName))
    .cmdAddUsers.Enabled = (UCase(strDocUser) = UCase(strUserName))
    .cmdRemoveUsers.Enabled = (UCase(strDocUser) = UCase(strUserName))
    .cmdOK.Enabled = (UCase(strDocUser) = UCase(strUserName))
    .cboPriority.Enabled = (UCase(strDocUser) = UCase(strUserName))
    .DTPInactive.Enabled = (UCase(strDocUser) = UCase(strUserName))
    
    If UCase(frmMain.cmdLock.Caption) = UCase("&Unlock File") Then
        .cmdAddUsers.Enabled = False
        .cmdRemoveUsers.Enabled = False
        .cmdOK.Enabled = False
'    Else
'        .cmdAddUsers.Enabled = True
'        .cmdRemoveUsers.Enabled = True
'        .cmdOK.Enabled = True
    End If

End With
End Sub

Sub setActiveUsers()
With frmFileProperties.MSFActiveUsers
    strFormatString = .FormatString
    .Clear
    .FormatString = strFormatString
    lngX = 0
    .Rows = 1
    
    Set rsUsers = DBFileManager.OpenRecordset("Select [User Name],View,Print,Edit,Copy from [Documents FilesUsers] where [Document Id]=" & strDocID)
    If rsUsers.EOF = False Then
        rsUsers.MoveFirst
        While rsUsers.EOF = False
            Set rsChkDept = DBFileManager.OpenRecordset("Select Department from Master_Users where Login='" & rsUsers![User Name] & "'")
            .Rows = .Rows + 1
            .Row = .Rows - 1
            While lngX < .Cols
               .Col = lngX
                If lngX = 0 Then
                    If rsChkDept.EOF = True Then
                    
                    ElseIf Len(rsChkDept!Department) > 0 Then
                        rsChkDept.MoveFirst
                        .Text = rsChkDept!Department
                    End If
                Else
                    .Text = rsUsers(lngX - 1)
                End If
                lngX = lngX + 1
            Wend
            lngX = 0
            rsUsers.MoveNext
        Wend
    End If
End With
End Sub

Sub loadRSUserActivity(FlexiName As MSFlexGrid, rsFlexiValues As Recordset)
With FlexiName
    .Redraw = False
    strFormatString = .FormatString
    .Clear
    .FormatString = strFormatString
    
    If rsFlexiValues.EOF = False Then
        rsFlexiValues.MoveFirst
        .Rows = 2
        lngCol = 0
        While rsFlexiValues.EOF = False
            .Row = .Rows - 1
            While lngCol < .Cols
                .Col = lngCol
                If Len(rsFlexiValues(lngCol)) > 0 Then
                    If (lngCol) = 0 Then
                        .Text = rsFlexiValues(lngCol)
                    ElseIf (lngCol) = 1 Then
                        .Text = Format(rsFlexiValues(lngCol), "dd-MMM-yyyy")
                    ElseIf (lngCol) = 2 Then
                        .Text = Format(rsFlexiValues(lngCol), "hh:mm AMPM")
                    End If
                End If
                lngCol = lngCol + 1
            Wend
            lngCol = 0
            .Rows = .Rows + 1
            rsFlexiValues.MoveNext
        Wend
    End If
    .Redraw = True
End With
End Sub

Sub updateActivityTable(strActivityType As String)
If Len(strActivityType) > 0 Then
    Set rsActivityLog = DBFileManager.OpenRecordset("Select [Document Id],User,Type,Date,Time from [Documents ActivityLog]")
    rsActivityLog.AddNew
    rsActivityLog![Document Id] = strDocID
    rsActivityLog![User] = strUserName
    rsActivityLog![Type] = strActivityType
    rsActivityLog![Date] = Format(nowDate, "dd-MMM-yyyy")
    rsActivityLog![Time] = Format(nowTime, "hh:mm AMPM")
    rsActivityLog.Update
End If
End Sub

Sub chkDeleteFile(strPhysicalFile As String, strDocumentID As String)
    'Call DeleteFile(strPhysicalFile)
    DBFileManager.Execute ("Delete * from [Documents FilesUsers] where [Document Id]=" & strDocumentID)
    DBFileManager.Execute ("Delete * from [Documents ActivityLog] where [Document Id]=" & strDocumentID)
    Set rsDocDetails = DBFileManager.OpenRecordset("select [Document Id],Revision,Name,Path,User,[Created Date],[Created Time],Department,[Modified By],[Modified Date],[Modified Time],[Description],Status,Expiry,Priority,[Parent Id] from [Documents] where [Document Id]=" & strDocumentID)
    Set rsDeleteFiles = DBFileManager.OpenRecordset("select [Document Id],Revision,Name,Path,User,[Created Date],[Created Time],Department,[Modified By],[Modified Date],[Modified Time],[Description],Status,Expiry,Priority,[Parent Id],[Deleted Date],[Deleted Time] from [Documents Deleted] where [Document Id]=" & strDocumentID)
    If rsDocDetails.EOF = False Then
        rsDocDetails.MoveFirst
        rsDeleteFiles.AddNew
        rsDeleteFiles![Document Id] = rsDocDetails![Document Id]
        rsDeleteFiles![Revision] = rsDocDetails![Revision]
        rsDeleteFiles![Name] = rsDocDetails![Name]
        rsDeleteFiles![Path] = rsDocDetails![Path]
        rsDeleteFiles![User] = rsDocDetails![User]
        rsDeleteFiles![Created Date] = rsDocDetails![Created Date]
        rsDeleteFiles![Created Time] = rsDocDetails![Created Time]
        rsDeleteFiles![Department] = rsDocDetails![Department]
        rsDeleteFiles![Modified By] = rsDocDetails![Modified By]
        rsDeleteFiles![Modified Date] = rsDocDetails![Modified Date]
        rsDeleteFiles![Modified Time] = rsDocDetails![Modified Time]
        rsDeleteFiles![Description] = rsDocDetails![Description]
        rsDeleteFiles![Status] = rsDocDetails![Status]
        rsDeleteFiles![expiry] = rsDocDetails![expiry]
        rsDeleteFiles![Priority] = rsDocDetails![Priority]
        rsDeleteFiles![Parent Id] = rsDocDetails![Parent Id]
        rsDeleteFiles![Deleted Date] = Format(nowDate, "dd-MMM-yyyy")
        rsDeleteFiles![Deleted Time] = Format(nowTime, "hh:mm AMPM")
        rsDeleteFiles.Update
    End If
    rsDocDetails.Delete
    
    MsgBox "File deleted successfully.", vbInformation
    frmMain.cmdMyFiles.Value = True
End Sub

Sub setColRef()
With frmMain
    If .cmdMyFiles.Value = True Then
        If .cboSortBy.Text = "Created On" Then
            strSortBy = "[Created Date],[Created Time] desc"
            .MSFFiles.Col = 3
        ElseIf .cboSortBy.Text = "File Name" Then
            strSortBy = "[Name]"
            .MSFFiles.Col = 0
        ElseIf .cboSortBy.Text = "Expiry On" Then
            strSortBy = "[Expiry]"
            .MSFFiles.Col = 6
        ElseIf .cboSortBy.Text = "Modified On" Then
            strSortBy = "[Modified Date],[Modified Time] desc"
            .MSFFiles.Col = 5
        ElseIf .cboSortBy.Text = "User" Then
            strSortBy = "[User]"
        ElseIf .cboSortBy.Text = "Department" Then
            strSortBy = "[Department]"
        End If
    ElseIf .cmdOthersFiles.Value = True Then
        If .cboSortBy.Text = "Created On" Then
            strSortBy = "[Created Date],[Created Time] desc"
            .MSFFiles.Col = 5
        ElseIf .cboSortBy.Text = "File Name" Then
            strSortBy = "[Name]"
            .MSFFiles.Col = 0
        ElseIf .cboSortBy.Text = "Expiry On" Then
            strSortBy = "[Expiry]"
            .MSFFiles.Col = 6
        ElseIf .cboSortBy.Text = "Modified On" Then
            strSortBy = "[Modified Date],[Modified Time] desc"
        ElseIf .cboSortBy.Text = "User" Then
            strSortBy = "[User]"
            .MSFFiles.Col = 3
        ElseIf .cboSortBy.Text = "Department" Then
            strSortBy = "[Department]"
            .MSFFiles.Col = 4
        End If
    ElseIf .cmdMyInactiveFiles.Value = True Then
        If .cboSortBy.Text = "Created On" Then
            strSortBy = "[Created Date],[Created Time] desc"
            .MSFFiles.Col = 5
        ElseIf .cboSortBy.Text = "File Name" Then
            strSortBy = "[Name]"
            .MSFFiles.Col = 0
        ElseIf .cboSortBy.Text = "Expiry On" Then
            strSortBy = "[Expiry]"
            .MSFFiles.Col = 8
        ElseIf .cboSortBy.Text = "Modified On" Then
            strSortBy = "[Modified Date],[Modified Time] desc"
            .MSFFiles.Col = 7
        ElseIf .cboSortBy.Text = "User" Then
            strSortBy = "[User]"
            .MSFFiles.Col = 3
        ElseIf .cboSortBy.Text = "Department" Then
            strSortBy = "[Department]"
            .MSFFiles.Col = 4
        End If
    ElseIf .cmdNewFiles.Value = True Then
        If .cboSortBy.Text = "Created On" Then
            strSortBy = "[Created Date],[Created Time] desc"
            .MSFFiles.Col = 5
        ElseIf .cboSortBy.Text = "File Name" Then
            strSortBy = "[Name]"
            .MSFFiles.Col = 0
        ElseIf .cboSortBy.Text = "Expiry On" Then
            strSortBy = "[Expiry]"
            .MSFFiles.Col = 6
        ElseIf .cboSortBy.Text = "Modified On" Then
            strSortBy = "[Modified Date],[Modified Time] desc"
        ElseIf .cboSortBy.Text = "User" Then
            strSortBy = "[User]"
            .MSFFiles.Col = 3
        ElseIf .cboSortBy.Text = "Department" Then
            strSortBy = "[Department]"
            .MSFFiles.Col = 4
        End If
    End If
End With
End Sub

Sub setPurgeDocumentID()
With frmPurgeRestore
    strDocID = ""
    strDocUser = ""
    strDeleteDate = ""
    strDeleteTime = ""
    If Len(.MSFFiles.Text) > 0 And .MSFFiles.Row > 0 Then
        .MSFFiles.Col = 1
        strDocUser = .MSFFiles.Text
        .MSFFiles.Col = 2
        strDocID = .MSFFiles.Text
        .MSFFiles.Col = 4
        strDeleteDate = Format(.MSFFiles.Text, "dd-MMM-yyyy")
        .MSFFiles.Col = 5
        strDeleteTime = Format(.MSFFiles.Text, "hh:mm AMPM")
        
        Set rsDocID = DBFileManager.OpenRecordset("Select [Document Id] from [Documents Deleted] where Name='" & strDocID & "' and User='" & strDocUser & "' and [Deleted Date]=#" & strDeleteDate & "# and [Deleted Time]=#" & strDeleteTime & "#")
        If rsDocID.EOF = False Then
            rsDocID.MoveFirst
            strDocID = rsDocID![Document Id]
        Else
            strDocID = ""
        End If
    End If
End With
End Sub

Sub uploadSingleFile()
frmMain.Enabled = False
With frmFileProperties
    boolFromFileUpload = True
    .Show , frmMain
    .lblFileName.Caption = frmMain.CommonDialog1.FileTitle
    .lblCreatedBy.Caption = strUserName
    .lblCreatedOn.Caption = nowDate & "  " & nowTime
    .lblModifiedBy.Caption = ""
    .lblModifiedOn.Caption = ""
    .lblRevision.Caption = "0"
    .cboPriority.Text = .cboPriority.List(1)
    .DTPInactive.Value = Format(DateAdd("m", 1, Now()), "dd-MMM-yyyy")
    strFilePath = strServerPath & strUserName & "\"
End With
End Sub

Sub setDocumentRevID()
With frmUploadRevision
    strDocRevID = ""
    strDocRevUser = ""
    If Len(.MSFFiles.Text) > 0 And .MSFFiles.Row > 0 Then
        .MSFFiles.Col = 0
        strDocRevID = .MSFFiles.Text
        .MSFFiles.Col = 3
        strDocRevUser = .MSFFiles.Text
        Set rsDocID = DBFileManager.OpenRecordset("Select [Document Id],[Parent Id] from Documents where Name='" & strDocRevID & "' and User='" & strDocRevUser & "'")
        If rsDocID.EOF = False Then
            rsDocID.MoveFirst
            strDocRevID = rsDocID![Document Id]
            If Len(rsDocID![Parent Id]) > 0 Then
                strDocRevParentID = rsDocID![Parent Id]
            Else
                strDocRevParentID = strDocID
            End If
        Else
            strDocRevID = ""
            strDocRevParentID = strDocID
        End If
    End If
End With
End Sub

Function setMaxRevNo(strDocRevID As Long) As Long
setMaxRevNo = 1
Set rsMaxRevNo = DBFileManager.OpenRecordset("select max(revision) from Documents where [Parent Id]=" & strDocRevID & " Or [Document Id]=" & strDocRevID)
If rsMaxRevNo.EOF = False Then
    rsMaxRevNo.MoveFirst
    If Len(rsMaxRevNo(0)) > 0 Then
        setMaxRevNo = rsMaxRevNo(0) + 1
    Else
        setMaxRevNo = 1
    End If
Else
    setMaxRevNo = 1
End If
End Function

Sub setActiveUsersRev()
With frmFileProperties.MSFActiveUsers
    strFormatString = .FormatString
    .Clear
    .FormatString = strFormatString
    lngX = 0
    .Rows = 2
    
    Set rsUsers = DBFileManager.OpenRecordset("Select [User Name],View,Print,Edit,Copy from [Documents FilesUsers] where [Document Id]=" & strDocRevID & " and [User Name]<>'" & strUserName & "'")
    If rsUsers.EOF = False Then
        rsUsers.MoveFirst
        While rsUsers.EOF = False
            Set rsChkDept = DBFileManager.OpenRecordset("Select Department from Master_Users where Login='" & rsUsers![User Name] & "'")
            .Rows = .Rows + 1
            .Row = .Rows - 2
            While lngX < .Cols
               .Col = lngX
                If lngX = 0 Then
                    If rsChkDept.EOF = True Then
                    
                    ElseIf Len(rsChkDept!Department) > 0 Then
                        rsChkDept.MoveFirst
                        .Text = rsChkDept!Department
                    End If
                Else
                    .Text = rsUsers(lngX - 1)
                End If
                lngX = lngX + 1
            Wend
            lngX = 0
            rsUsers.MoveNext
        Wend
    End If
End With
End Sub

Sub setHighlightFileName()
With frmMain.MSFFiles
    .SetFocus
    lngX = 1
    If Len(strSearchUserName) <= 0 Then
        .Col = 1
        While lngX < .Rows
            .Row = lngX
            If UCase(.Text) = UCase(strSearchFileName) Then
                .Col = 0
                .ColSel = .Cols - 1
                
                Exit Sub
            End If
            lngX = lngX + 1
        Wend
    ElseIf Len(strSearchUserName) >= 0 Then
        .Col = 1
        While lngX < .Rows
            .Row = lngX
            If UCase(.Text) = UCase(strSearchFileName) Then
                .Col = 4
                If UCase(.Text) = UCase(strSearchUserName) Then
                    .Col = 0
                    .ColSel = .Cols - 1
                End If
                Exit Sub
            End If
            lngX = lngX + 1
        Wend
    End If
    strSearchFileName = ""
End With
End Sub

Sub unLockFile()
If Len(strDocID) > 0 Then
    DBFileManager.Execute ("Delete * from [Documents FilesUsers] where [Document Id]=" & strDocID)
    Set rsLockDoc = DBFileManager.OpenRecordset("Select [Document Id],[User Name],View,Print,Copy,Edit from [Documents FilesUsers]")
    
    Set rsUnLockDoc = DBFileManager.OpenRecordset("Select [Document Id],[User Name],View,Print,Copy,Edit from [Documents LockedUsers] where [Document Id]=" & strDocID)
    If rsUnLockDoc.EOF = False Then
        rsUnLockDoc.MoveFirst
        While rsUnLockDoc.EOF = False
            rsLockDoc.AddNew
                rsLockDoc![Document Id] = rsUnLockDoc![Document Id]
                rsLockDoc![User Name] = rsUnLockDoc![User Name]
                rsLockDoc![View] = rsUnLockDoc![View]
                rsLockDoc![Print] = rsUnLockDoc![Print]
                rsLockDoc![Copy] = rsUnLockDoc![Copy]
                rsLockDoc![Edit] = rsUnLockDoc![Edit]
            rsLockDoc.Update
            rsUnLockDoc.Delete
            rsUnLockDoc.MoveNext
        Wend
    End If
End If
End Sub

Sub LockFile()
If Len(strDocID) > 0 Then
    Set rsLockDoc = DBFileManager.OpenRecordset("Select [Document Id],[User Name],View,Print,Copy,Edit from [Documents FilesUsers] where [Document Id]=" & strDocID)
    DBFileManager.Execute ("Delete * from [Documents LockedUsers] where [Document Id]=" & strDocID)
    Set rsUnLockDoc = DBFileManager.OpenRecordset("Select [Document Id],[User Name],View,Print,Copy,Edit from [Documents LockedUsers]")
    If rsLockDoc.EOF = False Then
        rsLockDoc.MoveFirst
        While rsLockDoc.EOF = False
            rsUnLockDoc.AddNew
                rsUnLockDoc![Document Id] = rsLockDoc![Document Id]
                rsUnLockDoc![User Name] = rsLockDoc![User Name]
                rsUnLockDoc![View] = rsLockDoc![View]
                rsUnLockDoc![Print] = rsLockDoc![Print]
                rsUnLockDoc![Copy] = rsLockDoc![Copy]
                rsUnLockDoc![Edit] = rsLockDoc![Edit]
            rsUnLockDoc.Update
            rsLockDoc.Edit
                rsLockDoc![View] = "N"
                rsLockDoc![Print] = "N"
                rsLockDoc![Copy] = "N"
                rsLockDoc![Edit] = "N"
            rsLockDoc.Update
            rsLockDoc.MoveNext
        Wend
        MsgBox "File locked successfully.", vbInformation
    Else
        MsgBox "The file cannot be locked since no users are defined.", vbExclamation
    End If
End If
End Sub

Function GetNextRev(DocId As String) As Integer
Dim rsNextRev As Recordset
Set rsNextRev = DBFileManager.OpenRecordset("Select max(revision) from Documents where [Document Id]=" & DocId & " or [Parent Id]=(select [Parent Id] from Documents where [Document Id]=" & DocId & ")")
If rsNextRev.EOF = False Then
    rsNextRev.MoveFirst
    GetNextRev = CInt(rsNextRev(0)) + 1
End If
End Function
