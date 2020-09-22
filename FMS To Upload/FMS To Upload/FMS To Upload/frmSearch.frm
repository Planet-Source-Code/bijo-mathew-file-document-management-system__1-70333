VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BackColor       =   &H00800000&
   Caption         =   " Document Search"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAccess 
      BackColor       =   &H00800000&
      Caption         =   "Show files which I have access to."
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
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   1680
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkExpiry 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   1040
      Width           =   255
   End
   Begin VB.CheckBox chkCreatedOn 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1040
      Width           =   255
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Search"
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
      Left            =   12060
      MouseIcon       =   "frmSearch.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   2535
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
      Left            =   6600
      MouseIcon       =   "frmSearch.frx":074C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9720
      Width           =   1695
   End
   Begin VB.TextBox txtDescription 
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
      Left            =   5520
      MaxLength       =   255
      TabIndex        =   3
      Top             =   360
      Width           =   5875
   End
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
      Left            =   1800
      MaxLength       =   255
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox txtUser 
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
      Left            =   12120
      MaxLength       =   255
      TabIndex        =   5
      Top             =   360
      Width           =   2415
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
      Left            =   14160
      MouseIcon       =   "frmSearch.frx":0A56
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
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
      ItemData        =   "frmSearch.frx":0D60
      Left            =   12240
      List            =   "frmSearch.frx":0D79
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFFiles 
      Height          =   7455
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   32896
      ForeColorFixed  =   -2147483628
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"frmSearch.frx":0DC6
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
   Begin MSComCtl2.DTPicker DTPCreated1 
      Height          =   345
      Left            =   2090
      TabIndex        =   8
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yy"
      Format          =   52953091
      CurrentDate     =   38496
   End
   Begin MSComCtl2.DTPicker DTPCreated2 
      Height          =   345
      Left            =   4150
      TabIndex        =   10
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yy"
      Format          =   52953091
      CurrentDate     =   38496
   End
   Begin MSComCtl2.DTPicker DTPickerExpiry1 
      Height          =   345
      Left            =   7650
      TabIndex        =   13
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yy"
      Format          =   52953091
      CurrentDate     =   38496
   End
   Begin MSComCtl2.DTPicker DTPickerExpiry2 
      Height          =   345
      Left            =   9840
      TabIndex        =   15
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yy"
      Format          =   52953091
      CurrentDate     =   38496
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   10335
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   16776
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
            TextSave        =   "2:57 AM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "3/15/2008"
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double click to goto the file..."
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
      Left            =   12120
      TabIndex        =   24
      Top             =   9520
      Width           =   2565
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
      Left            =   7080
      TabIndex        =   18
      Top             =   1800
      Width           =   405
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      TabIndex        =   14
      Top             =   1080
      Width           =   330
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Between:"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Left            =   3720
      TabIndex        =   9
      Top             =   1080
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created Between:"
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
      Left            =   225
      TabIndex        =   6
      Top             =   1080
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
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
      Left            =   11595
      TabIndex        =   4
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   915
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
      Left            =   11520
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSortBy_Change()
MSFFiles.Row = 1

If MSFFiles.Rows >= 4 Then
    If cboSortBy.Text = "File Name" Then
        MSFFiles.Col = 0
    ElseIf cboSortBy.Text = "Description" Then
        MSFFiles.Col = 1
    ElseIf cboSortBy.Text = "Revision" Then
        MSFFiles.Col = 2
    ElseIf cboSortBy.Text = "User" Then
        MSFFiles.Col = 3
    ElseIf cboSortBy.Text = "Created On" Then
        MSFFiles.Col = 4
    ElseIf cboSortBy.Text = "Expiry On" Then
        MSFFiles.Col = 5
    ElseIf cboSortBy.Text = "Location" Then
        MSFFiles.Col = 6
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

Private Sub cboSortBy_Click()
Call cboSortBy_Change
End Sub

Private Sub chkAccess_Click()
Call cmdSearch_Click
End Sub

Private Sub chkCreatedOn_Click()
If chkCreatedOn.Value = 0 Then
    DTPCreated1.Enabled = False
    DTPCreated1.Value = DTPCreated1.MinDate
    
    DTPCreated2.Enabled = False
    DTPCreated2.Value = DTPCreated2.MinDate
Else
    DTPCreated1.Enabled = True
    DTPCreated1.Value = Now
    
    DTPCreated2.Enabled = True
    DTPCreated2.Value = Now
End If
End Sub

Private Sub chkExpiry_Click()
If chkExpiry.Value = 0 Then
    DTPickerExpiry1.Enabled = False
    DTPickerExpiry1.Value = DTPickerExpiry1.MinDate
    
    DTPickerExpiry2.Enabled = False
    DTPickerExpiry2.Value = DTPickerExpiry2.MinDate
Else
    DTPickerExpiry1.Enabled = True
    DTPickerExpiry1.Value = Now
    
    DTPickerExpiry2.Enabled = True
    DTPickerExpiry2.Value = Now
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
boolRemoveFirst = False
If DTPCreated1.Value > DTPCreated2.Value Then
    MsgBox "The Date Created From should be lesser than Date Created To.", vbExclamation
    Exit Sub
ElseIf DTPickerExpiry1.Value > DTPickerExpiry2.Value Then
    MsgBox "The Expiry Date From should be lesser than Expiry Date To.", vbExclamation
    Exit Sub
End If

If chkCreatedOn.Value = 1 And chkExpiry.Value = 1 Then
    Set rsSearch = DBFileManager.OpenRecordset("Select Name,Description,Revision,User,[Created Date],[Expiry],[Document Id] from [Documents] where" _
    & " Name like'*" & Replace(txtFileName.Text, "'", "''") & "*' and" _
    & " Description like '*" & Replace(txtDescription.Text, "'", "''") & "*' and" _
    & " User like '*" & Replace(txtUser.Text, "'", "''") & "*' and" _
    & " ([Created Date]>=#" & Format(DTPCreated1.Value, "dd-MMM-yyyy") & "# and [Created Date]<=#" & Format(DTPCreated2.Value, "dd-MMM-yyyy") & "#) and" _
    & " ([Expiry]>=#" & Format(DTPickerExpiry1.Value, "dd-MMM-yyyy") & "# and [Expiry]<=#" & Format(DTPickerExpiry2.Value, "dd-MMM-yyyy") & "#)")
ElseIf chkCreatedOn.Value = 1 And chkExpiry.Value = 0 Then
    Set rsSearch = DBFileManager.OpenRecordset("Select Name,Description,Revision,User,[Created Date],[Expiry],[Document Id] from [Documents] where" _
    & " Name like'*" & Replace(txtFileName.Text, "'", "''") & "*' and" _
    & " Description like '*" & Replace(txtDescription.Text, "'", "''") & "*' and" _
    & " User like '*" & Replace(txtUser.Text, "'", "''") & "*' and" _
    & " ([Created Date]>=#" & Format(DTPCreated1.Value, "dd-MMM-yyyy") & "# and [Created Date]<=#" & Format(DTPCreated2.Value, "dd-MMM-yyyy") & "#)")
ElseIf chkCreatedOn.Value = 0 And chkExpiry.Value = 1 Then
    Set rsSearch = DBFileManager.OpenRecordset("Select Name,Description,Revision,User,[Created Date],[Expiry],[Document Id] from [Documents] where" _
    & " Name like'*" & Replace(txtFileName.Text, "'", "''") & "*' and" _
    & " Description like '*" & Replace(txtDescription.Text, "'", "''") & "*' and" _
    & " User like '*" & Replace(txtUser.Text, "'", "''") & "*' and" _
    & " ([Expiry]>=#" & Format(DTPickerExpiry1.Value, "dd-MMM-yyyy") & "# and [Expiry]<=#" & Format(DTPickerExpiry2.Value, "dd-MMM-yyyy") & "#)")
ElseIf chkCreatedOn.Value = 0 And chkExpiry.Value = 0 Then
    Set rsSearch = DBFileManager.OpenRecordset("Select Name,Description,Revision,User,[Created Date],[Expiry],[Document Id] from [Documents] where" _
    & " Name like'*" & Replace(txtFileName.Text, "'", "''") & "*' and" _
    & " Description like '*" & Replace(txtDescription.Text, "'", "''") & "*' and" _
    & " User like '*" & Replace(txtUser.Text, "'", "''") & "*'")
End If

With MSFFiles
    .Redraw = False
    strFormatString = .FormatString
    .Clear
    .FormatString = strFormatString
    .Rows = 2
    If rsSearch.EOF = False Then
        rsSearch.MoveFirst
        While rsSearch.EOF = False
            strFileLocation = ""
            .Row = .Rows - 1
            
            .Col = 0
            .Text = rsSearch!Name
            .Col = 1
            .Text = rsSearch!Description
            .Col = 2
            .Text = rsSearch!Revision
            .Col = 3
            .Text = rsSearch!User
            .Col = 4
            .Text = Format(rsSearch![Created Date], "dd-MMM-yyyy")
            .Col = 5
            .Text = Format(rsSearch![expiry], "dd-MMM-yyyy")
            
            If (UCase(rsSearch!User) = UCase(strUserName)) And DateDiff("d", Format(nowDate, "dd-MMM-yyyy"), Format(rsSearch![expiry], "dd-MMM-yyyy")) >= 0 Then
                strFileLocation = "My Files"
            ElseIf UCase(rsSearch!User) = UCase(strUserName) And DateDiff("d", Format(nowDate, "dd-MMM-yyyy"), Format(rsSearch![expiry], "dd-MMM-yyyy")) < 0 Then
                strFileLocation = "Inactive Files"
            Else
                Set rsLocation = DBFileManager.OpenRecordset("Select [User Name] from [Documents FilesUsers] where [User Name]='" & strUserName & "' and [Document Id]=" & rsSearch![Document Id])
                If rsLocation.EOF = True Then
                    strFileLocation = "Access Denied"
                Else
                    If DateDiff("d", Format(nowDate, "dd-MMM-yyyy"), Format(rsSearch![expiry], "dd-MMM-yyyy")) >= 0 Then
                        Set rsLocation = DBFileManager.OpenRecordset("Select * from [Documents ActivityLog] where [User]='" & rsLocation![User Name] & "' and [Document Id]=" & rsSearch![Document Id])
                        If rsLocation.EOF = True Then
                            strFileLocation = "New Files"
                        Else
                            strFileLocation = "Other's Files"
                        End If
                    ElseIf DateDiff("d", Format(nowDate, "dd-MMM-yyyy"), Format(rsSearch![expiry], "dd-MMM-yyyy")) < 0 Then
                        strFileLocation = "Inactive Files"
                    End If
                End If
            End If
            
            .Col = 6
            .Text = strFileLocation
            If strFileLocation = "Access Denied" And chkAccess.Value = 0 Then
                lngX = 0
                While lngX <= 6
                    .CellForeColor = &H808080
                    .Col = lngX
                    lngX = lngX + 1
                Wend
            ElseIf strFileLocation = "Access Denied" And chkAccess.Value = 1 Then
                On Error GoTo err:
                .RemoveItem .Row
            End If
    
            rsSearch.MoveNext
            .Rows = .Rows + 1
        Wend
    End If
    .Redraw = True
'this done so coz the first row cannot be removed dynamically when only 1 row is added
If boolRemoveFirst = True Then
    .RemoveItem 1
End If
lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"

Dim intX As Integer
For intX = 0 To MSFFiles.Cols - 1
    MSFFiles.ColAlignment(intX) = 1
Next
End With
Exit Sub

err:
boolRemoveFirst = True
Resume Next
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
DTPCreated1.Value = Now: DTPCreated2 = Now: DTPickerExpiry1 = Now: DTPickerExpiry2 = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub MSFFiles_DblClick()
With MSFFiles
    .Col = 0
    strSearchFileName = .Text
    .Col = 3
    strSearchUserName = .Text
    .Col = 6
    If .Text = "Access Denied" Or Len(.Text) <= 0 Then
        Exit Sub
    ElseIf .Text = "Other's Files" Then
        Unload frmSearch
        frmMain.Show
        frmMain.cmdOthersFiles.Value = True
    ElseIf .Text = "My Files" Then
        Unload frmSearch
        frmMain.Show
        frmMain.cmdMyFiles.Value = True
        strSearchUserName = ""
    ElseIf .Text = "Inactive Files" Then
        Unload frmSearch
        frmMain.Show
        frmMain.cmdMyInactiveFiles.Value = True
    ElseIf .Text = "New Files" Then
        Unload frmSearch
        frmMain.Show
        frmMain.cmdNewFiles.Value = True
    End If
End With
End Sub

Private Sub MSFFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If MSFFiles.MouseRow = 0 And Button = 1 Then Call MSF_SortFlexiNoArrows(MSFFiles, True)
End Sub
