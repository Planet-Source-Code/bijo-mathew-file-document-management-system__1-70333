Attribute VB_Name = "modServer"
Option Explicit

Public rsDateTime As Recordset
Public lngCol As Long

Public FSO

Public Const VK_ESCAPE = &H1B
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_CONTROL = &H11
Public Const VK_O = &H79

Declare Sub keybd_event Lib "USER32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function CloseClipboard Lib "USER32" () As Long

Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function gstrComputerName() As String
    gstrComputerName = String(200, Chr$(0))
    GetComputerName gstrComputerName, 200
    gstrComputerName = Left$(gstrComputerName, InStr(gstrComputerName, Chr$(0)) - 1)
End Function

Public Function nowDate() As String
Set rsDateTime = DBFileManager.OpenRecordset("SELECT format(now,'dd-MMM-yyyy')")
If rsDateTime.EOF = False Then
    rsDateTime.MoveFirst
    nowDate = Format(rsDateTime(0), "dd-mmm-yyyy")
Else
    nowDate = Format(Now, "dd-mmm-yyyy")
End If
End Function

Public Function nowTime() As String
Set rsDateTime = DBFileManager.OpenRecordset("SELECT format(now,'hh:mm:ss AMPM')")
If rsDateTime.EOF = False Then
    rsDateTime.MoveFirst
    nowTime = Format(rsDateTime(0), "hh:mm:ss AMPM")
Else
    nowTime = Format(Now, "hh:mm:ss AMPM")
End If
End Function

Public Function chkNoOfLicenses() As Boolean
chkNoOfLicenses = True
Set rsLicenses = DBFileManager.OpenRecordset("Select clients from [No Of Clients]")
Set rsUsers = DBFileManager.OpenRecordset("Select * from Master_Users")
If rsLicenses.EOF = True Then
    MsgBox "You have expired all your licenses. Please contact your vendor to update your licenses.", vbExclamation
    chkNoOfLicenses = False
Else
    If rsUsers.EOF = False Then
        rsUsers.MoveLast
        rsLicenses.MoveFirst
        If IsNumeric(Decrypt(rsLicenses!Clients, strPassword)) = True Then
            If rsUsers.RecordCount >= Decrypt(rsLicenses!Clients, strPassword) Then
                MsgBox "You have expired all your licenses. Please contact your vendor to update your licenses.", vbExclamation
                chkNoOfLicenses = False
            End If
        Else
            MsgBox "You have expired all your licenses. Please contact your vendor to update your licenses.", vbExclamation
             chkNoOfLicenses = False
        End If
    End If
End If
End Function

'read the value
Sub setDbPath()
dbPathString = GetSetting(App.EXEName, "DB Path", "DB Path")
If dbPathString = "" Then
    Unload frmLogin
    frmDBase.Show
End If
End Sub

Public Function chkFilePath(strFilePath As String) As Boolean
If (FSO.FileExists(strFilePath)) Then
    chkFilePath = True
Else
    chkFilePath = False
End If
End Function

Public Function chkFolderPath(strFilePath As String, boolCreateFolder As Boolean) As Boolean
If (FSO.FolderExists(strFilePath)) Then
    chkFolderPath = True
Else
    chkFolderPath = False
    If boolCreateFolder = True Then
        On Error GoTo err:
        FSO.CreateFolder (strFilePath)
        chkFolderPath = True
    End If
End If
Exit Function

err:
chkFolderPath = False
End Function

Sub loadRSFlexiValues(FlexiName As MSFlexGrid, rsFlexiValues As Recordset, Optional lngStartCol As Long = 0)
With FlexiName
Dim lngCol As Integer, lngDBCol As Integer, strFormatString As String
strFormatString = FlexiName.FormatString
.Redraw = False
.Rows = 1
.Rows = 2
FlexiName.FormatString = strFormatString
If rsFlexiValues.EOF = False Then
    rsFlexiValues.MoveFirst
    lngCol = lngStartCol: lngDBCol = 0
    While rsFlexiValues.EOF = False
        .Row = .Rows - 1
        While lngDBCol < rsFlexiValues.Fields.count
            .Col = lngCol
            .CellAlignment = 1
            If Len(rsFlexiValues(lngDBCol)) > 0 Then
                If IsDate(rsFlexiValues(lngDBCol)) = True And IsNumeric(rsFlexiValues(lngDBCol)) = False Then
                    If InStr(1, UCase(rsFlexiValues(lngDBCol)), "AM") <> 0 Or InStr(1, UCase(rsFlexiValues(lngDBCol)), "PM") <> 0 Then
                        .Text = Format(rsFlexiValues(lngDBCol), "dd-MMM-yyyy hh:mm AMPM")
                    Else
                        .Text = Format(rsFlexiValues(lngDBCol), "dd-MMM-yyyy")
                    End If
                Else
                    .Text = CStr(rsFlexiValues(lngDBCol))
                End If
            End If
            lngDBCol = lngDBCol + 1
            lngCol = lngCol + 1
        Wend
        lngCol = lngStartCol
        lngDBCol = 0
        .Rows = .Rows + 1
        rsFlexiValues.MoveNext
    Wend
End If
.Redraw = True

End With
End Sub

Sub cloneFlexi(MSFCopyFrom As MSFlexGrid, MSFCopyTo As MSFlexGrid)
MSFCopyTo.Redraw = False
strFormatString = MSFCopyTo.FormatString
MSFCopyTo.Clear
MSFCopyTo.FormatString = strFormatString
lngX = 1
lngY = 0
While lngX < MSFCopyFrom.Rows
    MSFCopyTo.Rows = lngX + 1
    MSFCopyTo.Row = lngX
    MSFCopyFrom.Row = lngX
    While lngY < MSFCopyFrom.Cols
        MSFCopyTo.Col = lngY
        MSFCopyFrom.Col = lngY
        MSFCopyTo.Text = MSFCopyFrom.Text
        lngY = lngY + 1
    Wend
    lngY = 0
    lngX = lngX + 1
Wend
MSFCopyTo.Redraw = True
End Sub

Sub saveFlexiToRecordset(MSFlexi As MSFlexGrid, RecSet As Recordset, Optional strAdditionalInfo As String, Optional intFieldNo As Integer)
lngX = 1
lngY = 0
If MSFlexi.Rows >= 2 Then
    MSFlexi.Row = MSFlexi.Rows - 1
    MSFlexi.Col = 0
    boolMSFLastLineText = False
    While MSFlexi.Col + 1 < MSFlexi.Cols
        If Len(Trim(MSFlexi.Text)) > 0 Then
            boolMSFLastLineText = True
        End If
        MSFlexi.Col = MSFlexi.Col + 1
    Wend
    
    If boolMSFLastLineText = True Then
        While lngX <= MSFlexi.Rows - 1
            MSFlexi.Row = lngX
            RecSet.AddNew
            While lngY < MSFlexi.Cols
                MSFlexi.Col = lngY
                If lngY = intFieldNo And Len(strAdditionalInfo) > 0 Then
                    RecSet(lngY) = strAdditionalInfo
                Else
                    RecSet(lngY) = MSFlexi.Text
                End If
                lngY = lngY + 1
            Wend
            lngY = 0
            If Len(intFieldNo) > 0 And Len(strAdditionalInfo) > 0 Then
                RecSet(intFieldNo) = strAdditionalInfo
            End If
            RecSet.Update
            lngX = lngX + 1
        Wend
    Else
        While lngX < MSFlexi.Rows - 1
            MSFlexi.Row = lngX
            RecSet.AddNew
            While lngY < MSFlexi.Cols
                MSFlexi.Col = lngY
                If lngY = intFieldNo And Len(strAdditionalInfo) > 0 Then
                    RecSet(lngY) = strAdditionalInfo
                Else
                    RecSet(lngY) = MSFlexi.Text
                End If
                lngY = lngY + 1
            Wend
            lngY = 0
            If Len(intFieldNo) > 0 And Len(strAdditionalInfo) > 0 Then
                RecSet(intFieldNo) = strAdditionalInfo
            End If
            RecSet.Update
            lngX = lngX + 1
        Wend
    End If
End If
End Sub

Function GetFileTitleFromPath(strPath As String) As String
Dim strFile() As String
strFile = Split(strPath, "\")
GetFileTitleFromPath = strFile(UBound(strFile))
End Function

Function chkFileExtension(strFileInfo As String) As String
If strFileInfo <> "" Then
    chkFileExtension = Mid(strFileInfo, Len(strFileInfo) - 2, 3)
End If
End Function

Function inDebugMode() As Boolean
On Error GoTo err:
inDebugMode = False
Debug.Print 0 / 0
Exit Function

err:
inDebugMode = True
End Function

Sub MSF_SortFlexiNoArrows(MSFGrid As MSFlexGrid, boolLastRowBlank As Boolean, Optional sortColNo As Integer)
With MSFGrid
Static mboolAsc As Boolean
'set the col no if passed as parameter
If sortColNo > 0 And sortColNo <= MSFGrid.Cols Then
    MSFGrid.Col = sortColNo
End If

'remove blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows - 1
End If

'sort
If mboolAsc = False Then
    .Sort = 6
'    .Row = 0
    mboolAsc = True
Else
    .Sort = 7
'    .Row = 0
    mboolAsc = False
End If

'add blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows + 1
End If

End With
End Sub

