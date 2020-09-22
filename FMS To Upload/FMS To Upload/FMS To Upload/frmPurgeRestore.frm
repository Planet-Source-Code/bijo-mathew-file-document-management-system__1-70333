VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPurgeRestore 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Purge / Restore Files"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   HelpContextID   =   20
   Icon            =   "frmPurgeRestore.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Purge / Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.TextBox txtFileName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3840
         MaxLength       =   255
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtDepartment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   240
         MaxLength       =   255
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cboSortBy 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmPurgeRestore.frx":0442
         Left            =   7680
         List            =   "frmPurgeRestore.frx":0455
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1935
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
         Left            =   9720
         MouseIcon       =   "frmPurgeRestore.frx":0490
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   450
      End
      Begin VB.CommandButton cmdPurge 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Purge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         MouseIcon       =   "frmPurgeRestore.frx":079A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton cmdRestore 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Restore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         MouseIcon       =   "frmPurgeRestore.frx":0AA4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         MaskColor       =   &H00C0FFFF&
         MouseIcon       =   "frmPurgeRestore.frx":0DAE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5280
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFFiles 
         Height          =   3735
         Left            =   240
         TabIndex        =   8
         Top             =   1275
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   32896
         ForeColorFixed  =   -2147483624
         BackColorSel    =   8388608
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         BorderStyle     =   0
         FormatString    =   "Department            |User                     |  File Name                        | Created On    | Date Deleted | Time Deleted"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblNoOfFiles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Left            =   9360
         TabIndex        =   12
         Top             =   5160
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type the department or user name or file name to search:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Left            =   6960
         TabIndex        =   5
         Top             =   930
         Width           =   690
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   6180
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9844
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
            TextSave        =   "7:59 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "6/2/2005"
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
End
Attribute VB_Name = "frmPurgeRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSortBy_Change()
    Call cboSortBy_Click
End Sub

Private Sub cboSortBy_Click()
MSFFiles.Row = 1
If MSFFiles.Rows >= 4 Then
    If cboSortBy.Text = "Department" Then
        MSFFiles.Col = 0
    ElseIf cboSortBy.Text = "User" Then
        MSFFiles.Col = 1
    ElseIf cboSortBy.Text = "File Name" Then
        MSFFiles.Col = 2
    ElseIf cboSortBy.Text = "Created On" Then
        MSFFiles.Col = 3
    ElseIf cboSortBy.Text = "Date Deleted" Then
        MSFFiles.Col = 5
        MSFFiles.ColSel = 4
    End If
    
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

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPurge_Click()
MSFFiles.Col = 2
If Len(MSFFiles.Text) > 0 Then
    boolConfirmDelete = MsgBox("Do you want to purge the file " & MSFFiles.Text & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirmDelete = vbYes Then
        If CLng(strDocID) > 0 Then
            MSFFiles.Col = 1
            strTempServerPath = strServerPath & MSFFiles.Text
            MSFFiles.Col = 2
            strTempServerPath = strTempServerPath & "\" & MSFFiles.Text
            
            Call DeleteFile(strTempServerPath)
            
            DBFileManager.Execute ("Delete * from [Documents Deleted] where [Document Id]=" & strDocID)
            MsgBox "File purged.", vbInformation
            Call Form_Load
        Else
            MsgBox "This file could not be purged.", vbExclamation
        End If
    End If
End If
End Sub

Private Sub cmdRestore_Click()
MSFFiles.Col = 2
If Len(MSFFiles.Text) > 0 Then
    boolConfirmDelete = MsgBox("Do you want to restore the file " & MSFFiles.Text & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirmDelete = vbYes Then
        If CLng(strDocID) > 0 Then
            Set rsPurgeRestore = DBFileManager.OpenRecordset("select [Document Id],Revision,Name,Path,User,[Created Date],[Created Time],Department,[Modified By],[Modified Date],[Modified Time],[Description],Status,Expiry,Priority,[Parent Id] from [Documents Deleted] where [Document Id]=" & strDocID)
            Set rsDocDetails = DBFileManager.OpenRecordset("select [Document Id],Revision,Name,Path,User,[Created Date],[Created Time],Department,[Modified By],[Modified Date],[Modified Time],[Description],Status,Expiry,Priority,[Parent Id] from [Documents]")
            If rsPurgeRestore.EOF = False Then
                rsPurgeRestore.MoveFirst
                rsDocDetails.AddNew
                'since Document Id is Auto Number
                'rsDocDetails![Document Id] = rsPurgeRestore![Document Id]
                rsDocDetails![Revision] = rsPurgeRestore![Revision]
                rsDocDetails![Name] = rsPurgeRestore![Name]
                rsDocDetails![Path] = rsPurgeRestore![Path]
                rsDocDetails![User] = rsPurgeRestore![User]
                rsDocDetails![Created Date] = rsPurgeRestore![Created Date]
                rsDocDetails![Created Time] = rsPurgeRestore![Created Time]
                rsDocDetails![Department] = rsPurgeRestore![Department]
                rsDocDetails![Modified By] = rsPurgeRestore![Modified By]
                rsDocDetails![Modified Date] = rsPurgeRestore![Modified Date]
                rsDocDetails![Modified Time] = rsPurgeRestore![Modified Time]
                rsDocDetails![Description] = rsPurgeRestore![Description]
                rsDocDetails![Status] = rsPurgeRestore![Status]
                rsDocDetails![Expiry] = Format(DateAdd("m", 1, Format(nowDate, "dd-MMM-yyyy")), "dd-MMM-yyyy")
                rsDocDetails![Priority] = rsPurgeRestore![Priority]
                rsDocDetails![Parent Id] = rsPurgeRestore![Parent Id]
                rsDocDetails.Update
            End If
            DBFileManager.Execute ("delete * from [Documents Deleted] where [Document Id]=" & strDocID)
            MsgBox "File restored.", vbInformation
            Call Form_Load
        Else
            MsgBox "This file could not be re-stored.", vbExclamation
        End If
    End If
End If
End Sub

Private Sub cmdSort_Click()
If cmdSort.Caption = "á" Then
    cmdSort.Caption = "â"
Else
    cmdSort.Caption = "á"
End If
Call cboSortBy_Change

End Sub

Private Sub Form_Load()
Set rsPurgeRestore = DBFileManager.OpenRecordset("Select Department,User,Name,[Created Date],[Deleted Date],[Deleted Time] from [Documents Deleted] order by [Deleted Date],[Deleted Time]desc")
Call loadRSFlexiValues(MSFFiles, rsPurgeRestore)
lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub MSFFiles_Click()
Call MSFFiles_LostFocus
End Sub

Private Sub MSFFiles_LostFocus()
If MSFFiles.Rows > 2 And MSFFiles.Text <> "" Then
    Call setPurgeDocumentID
    MSFFiles.Col = 0
    MSFFiles.ColSel = MSFFiles.Cols - 1
End If
End Sub

Private Sub txtDepartment_Change()
Set rsPurgeRestore = DBFileManager.OpenRecordset("Select Department,User,Name,[Created Date],[Deleted Date],[Deleted Time] from [Documents Deleted] where Department like '*" & Replace(txtDepartment.Text, "'", "''") & "*' and  User Like '*" & Replace(txtUserName.Text, "'", "''") & "*' and Name like '*" & Replace(txtFileName.Text, "'", "''") & "*' order by [Deleted Date],[Deleted Time]desc")
Call loadRSFlexiValues(MSFFiles, rsPurgeRestore)
lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"
End Sub

Private Sub txtFileName_Change()
Call txtDepartment_Change
End Sub

Private Sub txtUserName_Change()
Call txtDepartment_Change
End Sub
