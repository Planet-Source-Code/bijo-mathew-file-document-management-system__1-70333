VERSION 5.00
Object = "{94A0E92D-43C0-494E-AC29-FD45948A5221}#1.0#0"; "wiaaut.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00800000&
   Caption         =   " Main"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   4
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   960
         Top             =   0
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1320
         Top             =   0
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   480
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin WIACtl.CommonDialog winCD 
         Left            =   1680
         Top             =   0
      End
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0FFFF&
      Caption         =   "á"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14520
      MouseIcon       =   "frmMain.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   220
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.ComboBox cboSortBy 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmMain.frx":074C
      Left            =   12615
      List            =   "frmMain.frx":0762
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   210
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFFiles 
      Height          =   9615
      Left            =   2280
      TabIndex        =   17
      Top             =   600
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   16960
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   32896
      ForeColorFixed  =   -2147483628
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"frmMain.frx":07A7
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Caption         =   "Options"
      ForeColor       =   &H80000018&
      Height          =   10215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.Frame Frame1 
         BackColor       =   &H00800000&
         Caption         =   "File Options"
         ForeColor       =   &H80000018&
         Height          =   5295
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   1935
         Begin VB.CommandButton cmdScanUpload 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Scan && Upload"
            Height          =   375
            HelpContextID   =   5
            Left            =   120
            MouseIcon       =   "frmMain.frx":085C
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton cmdLock 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Lock File"
            Height          =   375
            HelpContextID   =   10
            Left            =   120
            MouseIcon       =   "frmMain.frx":0B66
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   4080
            Width           =   1695
         End
         Begin VB.CommandButton cmdUpload 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Upload"
            Height          =   375
            HelpContextID   =   5
            Left            =   120
            MouseIcon       =   "frmMain.frx":0E70
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdDownload 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Download for Editing"
            Height          =   495
            HelpContextID   =   6
            Left            =   120
            MouseIcon       =   "frmMain.frx":117A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton cmdOpen 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&View"
            Height          =   375
            HelpContextID   =   7
            Left            =   120
            MouseIcon       =   "frmMain.frx":1484
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2280
            Width           =   1695
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Delete"
            Height          =   375
            HelpContextID   =   8
            Left            =   120
            MouseIcon       =   "frmMain.frx":178E
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2880
            Width           =   1695
         End
         Begin VB.CommandButton cmdCopyOut 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Copy Out"
            Height          =   375
            HelpContextID   =   9
            Left            =   120
            MouseIcon       =   "frmMain.frx":1A98
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3480
            Width           =   1695
         End
         Begin VB.CommandButton cmdProperties 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Properties"
            Height          =   375
            HelpContextID   =   11
            Left            =   120
            MouseIcon       =   "frmMain.frx":1DA2
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   4680
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   240
         MouseIcon       =   "frmMain.frx":20AC
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   9720
         Width           =   1695
      End
      Begin VB.CommandButton cmdVersions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Versions"
         Height          =   375
         HelpContextID   =   12
         Left            =   240
         MouseIcon       =   "frmMain.frx":23B6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2715
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Search Files"
         Height          =   375
         HelpContextID   =   13
         Left            =   240
         MouseIcon       =   "frmMain.frx":26C0
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8880
         Width           =   1695
      End
      Begin VB.OptionButton cmdOthersFiles 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Other's Files"
         Height          =   375
         Left            =   240
         MouseIcon       =   "frmMain.frx":29CA
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1575
         Width           =   1695
      End
      Begin VB.OptionButton cmdMyInactiveFiles 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Inactive Files"
         Height          =   375
         Left            =   240
         MouseIcon       =   "frmMain.frx":2CD4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2130
         Width           =   1695
      End
      Begin VB.OptionButton cmdMyFiles 
         BackColor       =   &H00C0FFFF&
         Caption         =   "My &Files"
         Height          =   375
         Left            =   240
         MouseIcon       =   "frmMain.frx":2FDE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1035
         Width           =   1695
      End
      Begin VB.OptionButton cmdNewFiles 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&New Files"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         MouseIcon       =   "frmMain.frx":32E8
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   10335
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   17358
            MinWidth        =   2647
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "7:43 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "27-Mar-2008"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblNoOfFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   8160
      TabIndex        =   18
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By:"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   11880
      TabIndex        =   19
      Top             =   300
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   240
      Left            =   2280
      TabIndex        =   22
      Top             =   300
      Width           =   1035
   End
   Begin VB.Menu mnuViewFile 
      Caption         =   "&View Files    "
      Begin VB.Menu mnuNewFiles 
         Caption         =   "&New Files"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMyFiles 
         Caption         =   "&My Files"
      End
      Begin VB.Menu mnuOthersFiles 
         Caption         =   "&Other's Files"
      End
      Begin VB.Menu mnuMyInactiveFiles 
         Caption         =   "&Inactive Files"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search   "
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options   "
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password   "
         HelpContextID   =   14
      End
      Begin VB.Menu mnuSetUserRights 
         Caption         =   "&Set Default User Rights"
         HelpContextID   =   15
      End
      Begin VB.Menu mnuDblClick 
         Caption         =   "&Double Click To View File Properties"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Administer   "
      Visible         =   0   'False
      Begin VB.Menu mnuUsers 
         Caption         =   "&Users"
         HelpContextID   =   18
      End
      Begin VB.Menu mnuDepartments 
         Caption         =   "&Departments"
         HelpContextID   =   19
      End
      Begin VB.Menu mnuPurgeRestore 
         Caption         =   "&Purge / Re-store"
         HelpContextID   =   20
      End
      Begin VB.Menu sepa3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckLicense 
         Caption         =   "&Check License Status"
         HelpContextID   =   21
      End
      Begin VB.Menu mnuForceLogout 
         Caption         =   "&Force Logout"
         HelpContextID   =   22
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help   "
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Help Contents"
      End
      Begin VB.Menu sepa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit   "
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Log Off"
         HelpContextID   =   16
      End
      Begin VB.Menu sepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitFMS 
         Caption         =   "E&xit FMS"
      End
   End
   Begin VB.Menu mnuFileOptions 
      Caption         =   "File Options"
      Visible         =   0   'False
      Begin VB.Menu mnuFileName 
         Caption         =   ""
      End
      Begin VB.Menu sepa2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "&Upload"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "&Download for Editing"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuCopyOut 
         Caption         =   "&Copy Out"
      End
      Begin VB.Menu mnuLockFile 
         Caption         =   "&Lock File"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
      End
      Begin VB.Menu sepa4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersions 
         Caption         =   "&Versions"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSortBy_Change()
strSortBy = "[Created Date],[Created Time] desc"
MSFFiles.Row = 1
If MSFFiles.Rows >= 4 Then
    Call setColRef
    MSFFiles.RowSel = MSFFiles.Rows - 2
    If cmdSort.Caption = "á" Then
        MSFFiles.Sort = 1
    ElseIf cmdSort.Caption = "â" Then
        MSFFiles.Sort = 2
    End If
End If

MSFFiles.Row = 1
MSFFiles.RowSel = 1
End Sub

Private Sub cboSortBy_Click()
Call cboSortBy_Change
End Sub

Private Sub cmdCopyOut_Click()
boolConfirmDelete = MsgBox("Are you sure you want to copy this file " & MSFFiles.TextMatrix(MSFFiles.Row, 1) & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
If boolConfirmDelete = vbYes Then
    If lblFiles.Caption = "My Files" Then
        strTempServerPath = strServerPath & strUserName
        strTempServerPath = strTempServerPath & "\" & MSFFiles.TextMatrix(MSFFiles.Row, 1)
    Else
        strTempServerPath = strServerPath & MSFFiles.TextMatrix(MSFFiles.Row, 4)
        strTempServerPath = strTempServerPath & "\" & MSFFiles.TextMatrix(MSFFiles.Row, 1)
    End If
    
    Call chkFolderPath(App.Path & "\Download", True)
    
    CommonDialog2.CancelError = True
    CommonDialog2.DialogTitle = "Copy File..."
    CommonDialog2.Filter = "(*." & chkFileExtension(strTempServerPath) & ")|*." & chkFileExtension(strTempServerPath)
    CommonDialog2.FileName = App.Path & "\Download\" & MSFFiles.TextMatrix(MSFFiles.Row, 1)
    On Error GoTo err:
    CommonDialog2.ShowSave
  
    If Len(CommonDialog2.FileName) > 0 Then
        If chkFilePath(CommonDialog2.FileName) = True Then
            boolConfirmDelete = MsgBox("The file specified by already exist. Do you want to replace this file ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
            If boolConfirmDelete = vbYes Then
                Call CopyFile(strTempServerPath, CommonDialog2.FileName, 0)
                Call updateActivityTable("Copy")
                MsgBox "File copied successfully.", vbInformation
            Else
                MsgBox "File not copied.", vbExclamation
            End If
        Else
            Call CopyFile(strTempServerPath, CommonDialog2.FileName, 0)
            Call updateActivityTable("Copy")
            MsgBox "File copied successfully.", vbInformation
        End If
    End If
End If
Exit Sub

err:
End Sub

Private Sub cmdDelete_Click()
If Len(strDocID) <= 0 Then Exit Sub
Set rsCheckOut = DBFileManager.OpenRecordset("Select [Document Id],[Check Out Path],[User] from [Documents CheckOut] where [Document Id]=" & strDocID)
If rsCheckOut.EOF = True Then
    
    If lblFiles.Caption = "My Files" Then
        strTempServerPath = strServerPath & strUserName
        strTempServerPath = strTempServerPath & "\" & MSFFiles.TextMatrix(MSFFiles.Row, 1)
    Else
        strTempServerPath = strServerPath & MSFFiles.TextMatrix(MSFFiles.Row, 3)
        strTempServerPath = strTempServerPath & "\" & MSFFiles.TextMatrix(MSFFiles.Row, 1)
    End If
        
    If Len(MSFFiles.TextMatrix(MSFFiles.Row, 1)) > 0 And strDocID <> "" Then
        boolConfirmDelete = MsgBox("Are you sure you want to delete this file " & MSFFiles.TextMatrix(MSFFiles.Row, 1) & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirmDelete = vbYes Then
            boolConfirmDelete = MsgBox("If you delete the file the following will happen..." & _
            vbCrLf & "1. This will delete all the user access from the files." & _
            vbCrLf & "2. The file details will be deleted from the server." & _
            vbCrLf & "3. The physical file will be deleted from the server." & _
            vbCrLf & "4. This is an irreversible action." & vbCrLf & vbCrLf & _
            "Are you sure you want to continue deleting this file...?", vbYesNoCancel + vbDefaultButton3 + vbQuestion + vbCritical)
            If boolConfirmDelete = vbYes Then
                Call chkDeleteFile(strTempServerPath, CStr(strDocID))
                Call cmdMyFiles_Click
            End If
        End If
    End If
Else
    rsCheckOut.MoveFirst
    MsgBox "The user " & rsCheckOut!User & " has already downloaded this document for editing. Please delete once the user has updated the file.", vbExclamation
End If
End Sub

Private Sub cmdDownload_Click()
If cmdDownload.Caption = "&Download for Editing" Then
    boolConfirmDelete = MsgBox("Are you sure you want to download this file " & MSFFiles.TextMatrix(MSFFiles.Row, 1) & " for editing ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirmDelete = vbYes And Len(strDocID) > 0 Then
        Set rsCheckOut = DBFileManager.OpenRecordset("Select [User],[Document Id],[Check Out Path],[System Name]  from [Documents CheckOut] where " & _
            " [Documents CheckOut].[Document Id]=" & strDocID)
        If rsCheckOut.EOF = True Then
            If lblFiles.Caption = "My Files" Then
                strTempServerPath = strServerPath & strUserName & "\"
            Else
                strTempServerPath = strServerPath & MSFFiles.TextMatrix(MSFFiles.Row, 4) & "\"
            End If
            Dim strLocalFileName As String
            Call chkFolderPath(App.Path & "\Checkout\" & strUserName & "\", True)
            strLocalFileName = MSFFiles.TextMatrix(MSFFiles.Row, 1)
            
            Dim intX As Integer
            For intX = 0 To GetNextRev(strDocID) - 1
                strLocalFileName = Replace(strLocalFileName, " v" & intX & Right(strLocalFileName, 4), Right(strLocalFileName, 4), , , vbBinaryCompare)
            Next
            
            strLocalFileName = App.Path & "\Checkout\" & strUserName & "\" & Left(strLocalFileName, Len(strLocalFileName) - 4) & " v" & GetNextRev(strDocID) & Right(strLocalFileName, 4)
            
            If chkFilePath(strLocalFileName) = True Then
                MsgBox "The file already exists in the folder '..\Checkout\'" & strUserName & ". Please delete this file and download again.", vbExclamation
            Else
                Call CopyFile(strTempServerPath & MSFFiles.TextMatrix(MSFFiles.Row, 1), strLocalFileName, 1)
                If chkFilePath(strLocalFileName) = True Then
                    rsCheckOut.AddNew
                        rsCheckOut![Document Id] = strDocID
                        rsCheckOut![Check Out Path] = strLocalFileName
                        rsCheckOut![User] = strUserName
                        rsCheckOut![System Name] = gstrComputerName
                    rsCheckOut.Update
                    
                    MsgBox "Downloaded the file successfully. Press OK to open the file.", vbInformation
                    Call ShellExecute(0, vbNullString, strLocalFileName, vbNullString, "c:\temp", 3)
                Else
                    MsgBox "The file could not be downloaded for editing. Please try again later.", vbExclamation
                End If
            End If
        Else
            rsCheckOut.MoveFirst
            MsgBox "The user " & rsCheckOut!User & " has already downloaded this document for editing. Please download once the user has updated the file.", vbExclamation
        End If
    End If
Else
    boolConfirmDelete = MsgBox("Are you sure you want to upload this file " & MSFFiles.TextMatrix(MSFFiles.Row, 1) & " after editing ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirmDelete = vbYes And Len(strDocID) > 0 Then
        boolFromRevision = True
        Dim rsCheckOutDoc As Recordset
        If cmdMyFiles.Value = True Then
            Set rsCheckOutDoc = DBFileManager.OpenRecordset("SELECT [Document Id],[Check Out Path],[System Name] from [Documents CheckOut] where [Document id]=(select [document id] from documents where [Name]='" & MSFFiles.TextMatrix(MSFFiles.Row, 1) & "' and [User]='" & strUserName & "')")
        Else
            Set rsCheckOutDoc = DBFileManager.OpenRecordset("SELECT [Document Id],[Check Out Path],[System Name] from [Documents CheckOut] where [Document id]=(select [document id] from documents where [Name]='" & MSFFiles.TextMatrix(MSFFiles.Row, 1) & "' and [User]='" & MSFFiles.TextMatrix(MSFFiles.Row, 4) & "')")
        End If
        
        If rsCheckOutDoc.EOF = False Then
            If LCase(rsCheckOutDoc![System Name]) <> LCase(gstrComputerName) Then
                MsgBox "This document has been checked out from system: " & rsCheckOutDoc![System Name] & vbCrLf & "Try to check in the document from the same system.", vbExclamation
            Else
                If chkFilePath(rsCheckOutDoc![Check Out Path]) = True Then
                    Call setDocumentRevID
                    If Len(strDocRevID) > 0 Then
                        boolFromRevision = True
                        frmFileProperties.Show , Me
                        With frmFileProperties
                            frmMain.CommonDialog1.FileName = rsCheckOutDoc![Check Out Path]
                            .lblFileName.Caption = GetFileTitleFromPath(rsCheckOutDoc![Check Out Path])
                            .lblRevision.Caption = GetNextRev(CStr(strDocRevID))
                            .lblCreatedBy.Caption = strUserName
                            .lblCreatedOn.Caption = nowDate & "  " & nowTime
                            .DTPInactive.Value = DateAdd("m", 1, Format(nowDate, "dd-MMM-yyyy"))
                            .cboPriority.Text = .cboPriority.List(1)
                            .lblModifiedBy.Caption = ""
                            .lblModifiedOn.Caption = ""
                            .txtDescription.Text = MSFFiles.TextMatrix(MSFFiles.Row, 2)
                            Call setActiveUsersRev
                        End With
                    End If
                Else
                    boolConfirmDelete = MsgBox("Checked out document not found. Do you want to cancel your checkout on this file ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
                    If boolConfirmDelete = vbYes Then
                        DBFileManager.Execute ("Delete * from [Documents CheckOut] where [Document id]=" & CLng(rsCheckOutDoc![Document Id]))
                    End If
                End If
            End If
        Else
            MsgBox "Checked out details for document not found. Try again later", vbInformation
        End If
    End If
End If

End Sub

Private Sub cmdExit_Click()
Unload frmLogin
Unload Me
End Sub

Private Sub cmdLock_Click()
If Len(strDocID) > 0 And Len(MSFFiles.TextMatrix(MSFFiles.Row, 1)) > 0 Then
    Set rsCheckOut = DBFileManager.OpenRecordset("Select [Document Id],[Check Out Path],[User] from [Documents CheckOut] where [Document Id]=" & strDocID)
    If rsCheckOut.EOF = True Then
        If cmdLock.Caption = "&Lock File" Then
            boolConfirmDelete = MsgBox("Do you want to lock this file " & MSFFiles.TextMatrix(MSFFiles.Row, 1) & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
            If boolConfirmDelete = vbYes Then
                Call LockFile
            End If
        ElseIf cmdLock.Caption = "&Unlock File" Then
            boolConfirmDelete = MsgBox("Do you want to unlock this file " & MSFFiles.TextMatrix(MSFFiles.Row, 1) & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
            If boolConfirmDelete = vbYes Then
                Call unLockFile
                MsgBox "File unlocked successfully.", vbInformation
            End If
        End If
        Call MSFFiles_Click
    Else
        rsCheckOut.MoveFirst
        MsgBox "The user " & rsCheckOut!User & " has already downloaded this document for editing. Please lock once the user has updated the file.", vbExclamation
    End If
End If
End Sub

Private Sub cmdMyFiles_Click()
lblFiles.Caption = "My Files"
Call setMenusButtons(mnuMyFiles, cmdMyFiles)

MSFFiles.FormatString = "   |File Name                              |Description                                                        |Rev|Created On   |Last Modified By| Modified On | Expiry On    "
MSFFiles.Cols = 8
'MSFFiles.Rows = 1

MSFFiles.Rows = 1
MSFFiles.Rows = 2

Set rsLoadFiles = DBFileManager.OpenRecordset("select Name,Description,Revision,[Created Date],[Modified By],[Modified Date],Expiry from Documents where user='" & strUserName & "' and Expiry >= #" & Format(Now, "dd-MMM-yyyy") & "# order by " & strSortBy)
Call loadRSFlexiValues(MSFFiles, rsLoadFiles, 1)

MSFFiles.Redraw = True
Me.MousePointer = vbNormal
lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"

Call AlignCols

MSFFiles.Redraw = False
Dim intX As Integer
Set rsCheckOut = Nothing
Set rsCheckOut = DBFileManager.OpenRecordset("SELECT  [Documents CheckOut].user as Checked_Out_User, Documents.Name as Doc_Name,Documents.User as Doc_Owner " & _
" FROM Documents INNER JOIN [Documents CheckOut] ON Documents.[Document Id] = [Documents CheckOut].[Document Id]")
If rsCheckOut.EOF = False Then
    While rsCheckOut.EOF = False
        For intX = 1 To MSFFiles.Rows - 1
            If LCase(MSFFiles.TextMatrix(intX, 1)) = LCase(rsCheckOut!Doc_Name) Then
                 MSFFiles.Col = 0
                 MSFFiles.Row = intX
                 MSFFiles.CellFontName = "Webdings"
                 MSFFiles.Text = "Ï"
            End If
        Next
        rsCheckOut.MoveNext
    Wend
End If
MSFFiles.Redraw = True
End Sub

Private Sub cmdMyInactiveFiles_Click()
lblFiles.Caption = "My Inactive Files"
Call setMenusButtons(mnuMyInactiveFiles, cmdMyInactiveFiles)

MSFFiles.FormatString = "   |File Name                            |Description                                                        |Rev| User           |Department        |Created On  | Last Modified By  | Modified On|Expired On    "
MSFFiles.Cols = 10
'MSFFiles.Rows = 2
MSFFiles.Rows = 1
MSFFiles.Rows = 2

Call loadInactiveFiles

MSFFiles.Redraw = True
Me.MousePointer = vbNormal
lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"

Call AlignCols

MSFFiles.Redraw = False
Dim intX As Integer
Set rsCheckOut = Nothing
Set rsCheckOut = DBFileManager.OpenRecordset("SELECT  [Documents CheckOut].user as Checked_Out_User, Documents.Name as Doc_Name,Documents.User as Doc_Owner " & _
" FROM Documents INNER JOIN [Documents CheckOut] ON Documents.[Document Id] = [Documents CheckOut].[Document Id]")
If rsCheckOut.EOF = False Then
    While rsCheckOut.EOF = False
        For intX = 1 To MSFFiles.Rows - 1
            If LCase(MSFFiles.TextMatrix(intX, 1)) = LCase(rsCheckOut!Doc_Name) And LCase(MSFFiles.TextMatrix(intX, 4)) = LCase(rsCheckOut!Doc_Owner) Then
                 MSFFiles.Col = 0
                 MSFFiles.Row = intX
                 MSFFiles.CellFontName = "Webdings"
                 MSFFiles.Text = "Ï"
            End If
        Next
        rsCheckOut.MoveNext
    Wend
End If
MSFFiles.Redraw = True
End Sub

Private Sub cmdNewFiles_Click()
lblFiles.Caption = "New Files"
Call setMenusButtons(mnuNewFiles, cmdNewFiles)

MSFFiles.FormatString = "   |File Name                            |Description                                                        |Rev| User           |Department        |Created On   |Expiry On     "
MSFFiles.Cols = 8
'MSFFiles.Rows = 2
MSFFiles.Rows = 1
MSFFiles.Rows = 2

Call loadNewFiles

MSFFiles.Redraw = True
Me.MousePointer = vbNormal

lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"

Call AlignCols

MSFFiles.Redraw = False
Dim intX As Integer
Set rsCheckOut = Nothing
Set rsCheckOut = DBFileManager.OpenRecordset("SELECT  [Documents CheckOut].user as Checked_Out_User, Documents.Name as Doc_Name,Documents.User as Doc_Owner " & _
" FROM Documents INNER JOIN [Documents CheckOut] ON Documents.[Document Id] = [Documents CheckOut].[Document Id]")
If rsCheckOut.EOF = False Then
    While rsCheckOut.EOF = False
        For intX = 1 To MSFFiles.Rows - 1
            If LCase(MSFFiles.TextMatrix(intX, 1)) = LCase(rsCheckOut!Doc_Name) And LCase(MSFFiles.TextMatrix(intX, 4)) = LCase(rsCheckOut!Doc_Owner) Then
                 MSFFiles.Col = 0
                 MSFFiles.Row = intX
                 MSFFiles.CellFontName = "Webdings"
                 MSFFiles.Text = "Ï"
            End If
        Next
        rsCheckOut.MoveNext
    Wend
End If
MSFFiles.Redraw = True
End Sub

Private Sub cmdOthersFiles_Click()
lblFiles.Caption = "Other's Files"
Call setMenusButtons(mnuOthersFiles, cmdOthersFiles)

MSFFiles.FormatString = "   |File Name                            |Description                                                        |Rev| User           |Department        |Created On   |Expiry On     "
MSFFiles.Cols = 7
MSFFiles.Rows = 1
MSFFiles.Rows = 2

Call loadOtherUsersFiles

MSFFiles.Redraw = True
Me.MousePointer = vbNormal
lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"

Call AlignCols

MSFFiles.Redraw = False
Dim intX As Integer
Set rsCheckOut = Nothing
Set rsCheckOut = DBFileManager.OpenRecordset("SELECT  [Documents CheckOut].user as Checked_Out_User, Documents.Name as Doc_Name,Documents.User as Doc_Owner " & _
" FROM Documents INNER JOIN [Documents CheckOut] ON Documents.[Document Id] = [Documents CheckOut].[Document Id]")
If rsCheckOut.EOF = False Then
    While rsCheckOut.EOF = False
        For intX = 1 To MSFFiles.Rows - 1
            If LCase(MSFFiles.TextMatrix(intX, 1)) = LCase(rsCheckOut!Doc_Name) And LCase(MSFFiles.TextMatrix(intX, 4)) = LCase(rsCheckOut!Doc_Owner) Then
                 MSFFiles.Col = 0
                 MSFFiles.Row = intX
                 MSFFiles.CellFontName = "Webdings"
                 MSFFiles.Text = "Ï"
            End If
        Next
        rsCheckOut.MoveNext
    Wend
End If
MSFFiles.Redraw = True
End Sub

Private Sub cmdOpen_Click()
strTempServerPath = ""
boolPrint = False
If Len(strDocID) > 0 Then
    If lblFiles.Caption = "My Files" Then
        strTempServerPath = strServerPath & strUserName
        boolPrint = True
    Else
        strTempServerPath = strServerPath & MSFFiles.TextMatrix(MSFFiles.Row, 4)
        On Error Resume Next
        If rsUserFileRights.EOF = False Then
            On Error Resume Next
            If UCase(rsUserFileRights!Print) = "Y" Then
                boolPrint = True
            End If
        End If
    End If
End If

If MSFFiles.TextMatrix(MSFFiles.Row, 1) <> "" Then
    If chkFilePath(strTempServerPath & "\" & MSFFiles.TextMatrix(MSFFiles.Row, 1)) = True Then
        Call updateActivityTable("View")
        If LCase(Right(MSFFiles.TextMatrix(MSFFiles.Row, 1), 4)) <> ".pdf" Then
            Dim sValues  As String
            Call fEnumKey("HKCR", LCase(Right(MSFFiles.TextMatrix(MSFFiles.Row, 1), 4)), sValues)

            Dim strTypes() As String
            strTypes = Split(sValues, vbNullChar)
            Dim intX As Integer
            For intX = 0 To UBound(strTypes)
                Call fWriteValue("HKCU", "Software\Microsoft\Windows\Shell\AttachmentExecute\{0002DF01-0000-0000-C000-000000000046}", _
                  strTypes(intX), "B", 0)
            Next
            
            Unload frmDocument
            Load frmDocument
            Me.Enabled = False
            frmDocument.WebBrowser1.Navigate strTempServerPath & "\" & MSFFiles.TextMatrix(MSFFiles.Row, 1)
            frmDocument.Show , Me
            frmDocument.WebBrowser1.Visible = True
            frmDocument.Caption = "Opened " & MSFFiles.TextMatrix(MSFFiles.Row, 1)
            frmDocument.cmdPrint.Enabled = boolPrint

            For intX = 0 To UBound(strTypes)
                 Call fDeleteValue("HKCU", "Software\Microsoft\Windows\Shell\AttachmentExecute\{0002DF01-0000-0000-C000-000000000046}", strTypes(intX))
            Next
        Else
            Call ShellExecute(0, vbNullString, strTempServerPath & "\" & MSFFiles.TextMatrix(MSFFiles.Row, 1), vbNullString, "c:\temp", 3)
        End If
    Else
        MsgBox "The requested file could not be found on the server. Please contact your administrator if you think the file should be available.", vbExclamation
    End If
End If

'Exit Sub
'
'err:
End Sub

Private Sub cmdProperties_Click()
If Len(strDocID) > 0 Then
    boolFromFileUpload = False
    Me.Enabled = False
    Call setFileProperties
    frmFileProperties.Show , Me
End If
End Sub

Private Sub cmdScanUpload_Click()
Me.Enabled = False
frmScan.Show , Me
End Sub

Private Sub cmdSearch_Click()
frmSearch.Show , Me
End Sub

Private Sub cmdSort_Click()
If cmdSort.Caption = "á" Then
    cmdSort.Caption = "â"
Else
    cmdSort.Caption = "á"
End If
Call cboSortBy_Change
End Sub

Private Sub cmdUpload_Click()
boolFromFileUpload = False
CommonDialog1.DialogTitle = "Upload File..."
CommonDialog1.Filter = "All Files (*.html;*.htm;*.xls;*.csv;*.pps;*.doc;*.pdf;*.jpg;*.jpeg;*.gif;*.rtf;*.txt;*.tif;*.tiff)|*.html;*.htm;*.xls;*.csv;*.pps;*.doc;*.pdf;*.jpg;*.jpeg;*.gif;*.rtf;*.txt;*.tif;*.tiff| HTM/HTML (*.html;*.htm)|*.html;*.htm|Microsoft Excel (*.xls;*.csv)|*.xls;*.csv|Microsoft PowerPoint Show (*.pps)|*.pps|Microsoft Word (*.doc)|*.doc|PDF (*.pdf)|*.pdf|Pictures (*.jpg;*.jpeg;*.gif)|*.jpg;*.jpeg;*.gif|Rich Text File (*.rtf)|*.rtf|Text (*.txt)|*.txt|TIF/TIFF (*.tif;*.tiff)|*.tif;*.tiff"
CommonDialog1.CancelError = True
On Error GoTo err:
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    If chkFilePath(CommonDialog1.FileName) = True Then
'        If Replace(UCase(CommonDialog1.FileName), UCase(CommonDialog1.FileTitle), "") = UCase(App.Path & "\DOWNLOAD\") Then
'            boolConfirmDelete = MsgBox("Is this a revision document for any existing file ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
'            If boolConfirmDelete = vbYes Then
'                Me.Enabled = False
'                frmUploadRevision.Show , Me
'            ElseIf boolConfirmDelete = vbNo Then
                Call uploadSingleFile
'            End If
'        Else
'            Call uploadSingleFile
'        End If
    Else
        MsgBox "The file name specified by you does not exist. Please enter a different file name.", vbExclamation
    End If
End If
Exit Sub
err:
End Sub

Private Sub cmdVersions_Click()
Timer2.Enabled = False
cmdVersions.BackColor = &HC0FFFF
cmdVersions.Caption = "&Versions"
With frmVersions.MSFFiles
If Len(strDocID) > 0 Then
    Set rsOldRevisions = DBFileManager.OpenRecordset("Select Name,Description,Revision,User,[Created Date]  from Documents where [Parent Id]=" & strDocParentId & " or [Document Id]=" & strDocParentId & " order by revision desc")
    If rsOldRevisions.EOF = False Then
        frmVersions.Show , Me
        Me.Enabled = False
        Call loadRSFlexiValues(frmVersions.MSFFiles, rsOldRevisions)
    Else
        MsgBox "No revisions for this file.", vbExclamation
    End If
End If
End With
End Sub

Public Sub Form_Activate()
boolFromRevision = False
chkLicenseLog = True
If cmdNewFiles.Value = True Then
    Call cmdNewFiles_Click
ElseIf cmdMyFiles.Value = True Then
    Call cmdMyFiles_Click
ElseIf cmdOthersFiles.Value = True Then
    Call cmdOthersFiles_Click
ElseIf cmdMyInactiveFiles.Value = True Then
    Call cmdMyInactiveFiles_Click
End If

If Len(strSearchFileName) > 0 Then
    Call setHighlightFileName
End If
End Sub

Private Sub Form_Load()
strSortBy = "[Created Date],[Created Time] desc"
boolFromRevision = False

cmdNewFiles.Value = True
mnuLogOff.Caption = "&Log Off " & strUserName
Me.Caption = " Main " & " User : " & strUserName

If Len(GetSetting(App.EXEName, "DblClick", "Props")) > 0 Then
    mnuDblClick.Checked = CBool(GetSetting(App.EXEName, "DblClick", "Props"))
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmLogin.Visible = False And chkLicenseLog = True Then
    boolConfirmDelete = MsgBox("Are you sure you want to exit ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirmDelete = vbYes Then
        DBFileManager.Execute ("Delete * from Logins where Login='" & strUserName & "'")
        Unload frmLogin
        Dim objFrm As Form
        For Each objFrm In Forms
            Unload objFrm
        Next
    Else
        Cancel = 1
    End If
End If
End Sub

Private Sub mnuAbout_Click()
frmAboutSoft.Show , Me
End Sub

Private Sub mnuChangePassword_Click()
Call setNewUserPassword
End Sub

Private Sub mnuCheckLicense_Click()
Set rsLicenses = DBFileManager.OpenRecordset("Select clients from [No Of Clients]")
Set rsUsers = DBFileManager.OpenRecordset("Select * from Master_Users")
If rsLicenses.EOF = False Then
    rsLicenses.MoveFirst
    If IsNumeric(Decrypt(rsLicenses!Clients, strPassword)) = True Then
        strNoOfLicenses = Decrypt(rsLicenses!Clients, strPassword)
    Else
        strNoOfLicenses = "Not Configured."
    End If
Else
    strNoOfLicenses = "Not Configured."
End If

If rsUsers.EOF = False Then
    rsUsers.MoveLast
    strNoOfUsers = rsUsers.RecordCount
Else
    strNoOfUsers = 0
End If

Dim strDemo As String
Dim rsDemo As Recordset
strDemo = ""
Set rsDemo = DBFileManager.OpenRecordset("Select Demo,Recs from [Demo]")
If rsDemo.EOF = False Then
    rsDemo.MoveFirst
    If LCase(Decrypt(rsDemo(0), strPassword)) = "yes" Then
        strDemo = vbCrLf & "Demo version, for " & Decrypt(rsDemo(1), strPassword) & " records."
    End If
End If

If IsNumeric(strNoOfLicenses) = True And IsNumeric(strNoOfUsers) = True Then
    MsgBox "No of Licenses : " & strNoOfLicenses & vbCrLf & vbCrLf & "No of Used Licenses : " & strNoOfUsers & vbCrLf & "No of Free Licenses  : " & strNoOfLicenses - strNoOfUsers & vbCrLf & vbCrLf & "Please contact your vendor in case if you want to update your licenses." & strDemo, vbInformation
Else
    MsgBox "No of Licenses : " & strNoOfLicenses & vbCrLf & vbCrLf & "No of Used Licenses : " & strNoOfUsers & vbCrLf & vbCrLf & "Please contact your vendor in case if you want to update your licenses." & strDemo, vbInformation
End If
End Sub

Private Sub mnuCopyOut_Click()
Call cmdCopyOut_Click
End Sub

Private Sub mnuDblClick_Click()
If mnuDblClick.Checked = True Then
    mnuDblClick.Checked = False
Else
    mnuDblClick.Checked = True
End If

SaveSetting App.EXEName, "DblClick", "Props", mnuDblClick.Checked
End Sub

Private Sub mnuDelete_Click()
Call cmdDelete_Click
End Sub

Private Sub mnuDepartments_Click()
Me.Enabled = False
frmDeptNames.Show , Me
End Sub

Private Sub mnuDownload_Click()
Call cmdDownload_Click
End Sub

Private Sub mnuExitFMS_Click()
Call cmdExit_Click
End Sub

Private Sub mnuForceLogout_Click()
Me.Enabled = False
frmForceLogout.Show , Me
End Sub

Private Sub mnuHelpContents_Click()
SendKeys "{F1}"
End Sub

Private Sub mnuLockFile_Click()
Call cmdLock_Click
End Sub

Private Sub mnuLogOff_Click()
boolConfirmDelete = MsgBox("Are you sure you want to Log Off ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
If boolConfirmDelete = vbYes Then
    Load frmLogin
    frmLogin.Visible = True
    Unload Me
    DBFileManager.Execute ("Delete * from Logins where Login='" & strUserName & "'")
End If
End Sub

Private Sub mnuMyFiles_Click()
Call cmdMyFiles_Click
End Sub

Private Sub mnuMyInactiveFiles_Click()
Call cmdMyInactiveFiles_Click
End Sub

Private Sub mnuNewFiles_Click()
Call cmdNewFiles_Click
End Sub

Private Sub mnuOthersFiles_Click()
Call cmdOthersFiles_Click
End Sub

Private Sub mnuProperties_Click()
Call cmdProperties_Click
End Sub

Private Sub mnuPurgeRestore_Click()
frmPurgeRestore.Show , Me
Me.Enabled = False
End Sub

Private Sub mnuSearch_Click()
Call cmdSearch_Click
End Sub

Private Sub mnuSetUserRights_Click()
frmSetUserRights.Show , Me
Me.Enabled = False
End Sub

Private Sub mnuUpload_Click()
 Call cmdUpload_Click
End Sub

Private Sub mnuUsers_Click()
frmUserRoles.Show , Me
End Sub

Private Sub MSFInactiveFiles_LostFocus()
strLastFlexiName = "MSFInactiveFiles"
End Sub

Private Sub MSFMyFiles_LostFocus()
strLastFlexiName = "MSFMyFiles"
End Sub

Private Sub MSFNewFiles_LostFocus()
strLastFlexiName = "MSFNewFiles"
End Sub

Private Sub MSFOthersFiles_LostFocus()
strLastFlexiName = "MSFOtherFiles"
End Sub

Private Sub mnuVersions_Click()
Call cmdVersions_Click
End Sub

Private Sub mnuView_Click()
Call cmdOpen_Click
End Sub

Private Sub MSFFiles_Click()
If MSFFiles.Rows > 2 And MSFFiles.TextMatrix(MSFFiles.Row, 1) <> "" Then
    Call setDocumentID
    Call setUserDocRights
'    MSFFiles.ColSel = MSFFiles.Cols - 1
    cmdProperties.Enabled = True
Else
    cmdProperties.Enabled = False
    'Exit Sub
End If

'checking whether newer versions are available
cmdVersions.Caption = "&Versions"
cmdVersions.BackColor = &HC0FFFF
Timer2.Enabled = False
If Len(strDocID) > 0 And Len(MSFFiles.TextMatrix(MSFFiles.Row, 1)) > 0 Then
    Set rsOldRevisions = DBFileManager.OpenRecordset("Select max(Revision) as Rev from Documents where [Document Id]=" & strDocID & " or [Parent Id]=(select [Parent Id] from Documents where [Document Id]=" & strDocID & ")")
    If rsOldRevisions.EOF = False Then
        If rsOldRevisions!Rev > CLng(MSFFiles.TextMatrix(MSFFiles.Row, 3)) Then
            Timer2.Enabled = True
        End If
    End If
End If

'checking whether the document is locked
If Len(strDocID) > 0 And Len(MSFFiles.Text) > 0 Then
    Set rsLockDoc = DBFileManager.OpenRecordset("Select * from [Documents LockedUsers] where [Document Id]=" & strDocID)
    If rsLockDoc.EOF = False Then
        cmdLock.Caption = "&Unlock File"
        cmdDownload.Enabled = False
        cmdCopyOut.Enabled = False
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdLock.Caption = "&Lock File"
    End If
End If

'MSFFiles.Col = 0
'MSFFiles.ColSel = MSFFiles.Cols - 1
End Sub

Private Sub MSFFiles_DblClick()
If mnuDblClick.Checked = True Then
    Call cmdProperties_Click
ElseIf cmdOpen.Enabled = True Then
    Call cmdOpen_Click
End If
End Sub

Private Sub MSFFiles_LostFocus()
Call MSFFiles_Click
End Sub

Private Sub MSFFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If MSFFiles.MouseRow = 0 And Button = 1 Then Call MSF_SortFlexiNoArrows(MSFFiles, True)
End Sub

Private Sub MSFFiles_Mouseup(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call MSFFiles_Click
If Button = 2 And Len(MSFFiles.TextMatrix(MSFFiles.Row, 1)) > 0 Then
'    PopupMenu mnuFileOptions
    mnuFileName.Caption = "Options for " & MSFFiles.TextMatrix(MSFFiles.Row, 1)
    mnuUpload.Enabled = cmdUpload.Enabled
    mnuDownload.Enabled = cmdDownload.Enabled
    mnuView.Enabled = cmdOpen.Enabled
    mnuDelete.Enabled = cmdDelete.Enabled
    mnuCopyOut.Enabled = cmdCopyOut.Enabled
    mnuProperties.Enabled = cmdProperties.Enabled
    mnuVersions.Enabled = cmdVersions.Enabled
    mnuVersions.Caption = cmdVersions.Caption
    mnuLockFile.Enabled = cmdLock.Enabled
    mnuLockFile.Caption = cmdLock.Caption
End If

If Button = 2 And Len(MSFFiles.TextMatrix(MSFFiles.Row, 1)) > 0 Then
    PopupMenu mnuFileOptions
    mnuFileName.Caption = "Options for " & MSFFiles.TextMatrix(MSFFiles.Row, 1)
    mnuUpload.Enabled = cmdUpload.Enabled
    mnuDownload.Enabled = cmdDownload.Enabled
    mnuView.Enabled = cmdOpen.Enabled
    mnuDelete.Enabled = cmdDelete.Enabled
    mnuCopyOut.Enabled = cmdCopyOut.Enabled
    mnuProperties.Enabled = cmdProperties.Enabled
    mnuVersions.Enabled = cmdVersions.Enabled
    mnuVersions.Caption = cmdVersions.Caption
    mnuLockFile.Enabled = cmdLock.Enabled
    mnuLockFile.Caption = cmdLock.Caption
End If

End Sub

Private Sub Timer1_Timer()
chkLicenseLog = True
Dim rsChkLoginStatus As Recordset
Dim rsChkDemoRecs As Recordset
If Me.Enabled = True And UCase(strUserName) <> "ADMIN" Then
    Set rsChkLoginStatus = DBFileManager.OpenRecordset("Select * from Logins where Login='" & strUserName & "'")
    If rsChkLoginStatus.EOF = True Then
        MsgBox "You have been forcefully logged out by the administrator. " & vbCrLf & "You will now be logged off automatically.", vbCritical + vbOKOnly
        chkLicenseLog = False
        Timer1.Enabled = False
        Unload frmLogin
        Unload frmMain
    End If
End If

Set rsChkLoginStatus = DBFileManager.OpenRecordset("Select Demo,Recs from Demo")
Set rsChkDemoRecs = DBFileManager.OpenRecordset("Select * from Documents")
If rsChkDemoRecs.EOF = False And rsChkLoginStatus.EOF = False Then
    rsChkDemoRecs.MoveLast
    rsChkLoginStatus.MoveFirst
    If Decrypt(rsChkLoginStatus!Demo, strPassword) <> "No" Then
        If rsChkDemoRecs.RecordCount > CLng(Decrypt(rsChkLoginStatus!Recs, strPassword)) Then
            MsgBox "You demo version has expired. Please contact your vendor to update your licenses.", vbCritical + vbOKOnly
            chkLicenseLog = False
            Timer1.Enabled = False
            Unload frmLogin
            Unload frmMain
        End If
    End If
ElseIf rsChkDemoRecs.EOF = False Then
    MsgBox "You demo version has expired. Please contact your vendor to update your licenses.", vbCritical + vbOKOnly
    chkLicenseLog = False
    Timer1.Enabled = False
    Unload frmLogin
    Unload frmMain
End If
End Sub

Private Sub Timer2_Timer()
If cmdVersions.BackColor = &HC0FFFF Then
    cmdVersions.BackColor = &HFF&
    cmdVersions.Caption = "New &Versions"
Else
    cmdVersions.BackColor = &HC0FFFF
    cmdVersions.Caption = "New &Versions"
End If
End Sub

Sub AlignCols()
Dim intX As Integer
For intX = 0 To MSFFiles.Cols - 1
    MSFFiles.ColAlignment(intX) = 1
Next
End Sub
