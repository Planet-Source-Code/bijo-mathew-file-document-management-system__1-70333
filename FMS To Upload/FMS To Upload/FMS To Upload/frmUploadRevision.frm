VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmUploadRevision 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Select Revision Document"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   Icon            =   "frmUploadRevision.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmUploadRevision.frx":0442
      Left            =   9255
      List            =   "frmUploadRevision.frx":0452
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   1815
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
      Left            =   11160
      MouseIcon       =   "frmUploadRevision.frx":0480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   375
      Width           =   450
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&OK"
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
      Left            =   4080
      MouseIcon       =   "frmUploadRevision.frx":078A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8760
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
      Left            =   6120
      MouseIcon       =   "frmUploadRevision.frx":0A94
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8760
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFFiles 
      Height          =   7695
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13573
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   32896
      ForeColorFixed  =   -2147483628
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"frmUploadRevision.frx":0D9E
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   9345
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11775
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
            TextSave        =   "9:52 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "3/11/2008"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   480
      Width           =   405
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
      Left            =   8520
      TabIndex        =   1
      Top             =   450
      Width           =   690
   End
End
Attribute VB_Name = "frmUploadRevision"
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
    If cboSortBy.Text = "File Name" Then
        MSFFiles.Col = 0
    ElseIf cboSortBy.Text = "Description" Then
        MSFFiles.Col = 1
    ElseIf cboSortBy.Text = "User" Then
        MSFFiles.Col = 3
    ElseIf cboSortBy.Text = "Created On" Then
        MSFFiles.Col = 4
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
frmMain.Enabled = True
frmMain.SetFocus
Unload Me
End Sub

Private Sub cmdOK_Click()
boolFromRevision = False
If MSFFiles.Rows >= 2 Then
    boolFromRevision = True
    Call setDocumentRevID
    If Len(strDocRevID) > 0 Then
        Unload frmUploadRevision
        frmFileProperties.Show , Me
        With frmFileProperties
            .lblFileName.Caption = frmMain.CommonDialog1.FileTitle
            .lblRevision.Caption = setMaxRevNo(CLng(strDocRevID))
            .lblCreatedBy.Caption = strUserName
            .lblCreatedOn.Caption = nowDate & "  " & nowTime
            .DTPInactive.Value = DateAdd("m", 1, Format(nowDate, "dd-MMM-yyyy"))
            .cboPriority.Text = .cboPriority.List(1)
            .lblModifiedBy.Caption = ""
            .lblModifiedOn.Caption = ""
            MSFFiles.Col = 1
            .txtDescription.Text = MSFFiles.Text
            Call setActiveUsersRev
            
        End With
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
Set rsCheckOutIDs = DBFileManager.OpenRecordset("Select [Document Id] from [Documents CheckOut] where user='" & strUserName & "'")
If rsCheckOutIDs.EOF = False Then
    rsCheckOutIDs.MoveFirst
    MSFFiles.Rows = 2
    While rsCheckOutIDs.EOF = False
        MSFFiles.Row = MSFFiles.Rows - 1
        Set rsChecOutDocs = DBFileManager.OpenRecordset("Select Name,Description,Revision,User,[Created Date] from Documents where [Document Id]=" & rsCheckOutIDs![Document Id] & " order by Name")
        If rsChecOutDocs.EOF = False Then
            rsChecOutDocs.MoveFirst
            MSFFiles.Col = 0
            MSFFiles.Text = rsChecOutDocs!Name
            
            MSFFiles.Col = 1
            If Len(rsChecOutDocs!Description) > 0 Then
                MSFFiles.Text = rsChecOutDocs!Description
            End If
            
            MSFFiles.Col = 2
            MSFFiles.Text = rsChecOutDocs!Revision
            
            MSFFiles.Col = 3
            MSFFiles.Text = rsChecOutDocs!User
            
            MSFFiles.Col = 4
            MSFFiles.Text = Format(rsChecOutDocs![Created Date], "dd-MMM-yyyy")
        End If
        rsCheckOutIDs.MoveNext
        MSFFiles.Rows = MSFFiles.Rows + 1
    Wend
    lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"
Else
    MsgBox "You do not have any check out documents to create a revision.", vbExclamation
    frmMain.Enabled = True
    Unload Me
End If

Dim intX As Integer
For intX = 0 To MSFFiles.Cols - 1
    MSFFiles.ColAlignment(intX) = 1
Next
End Sub

Private Sub MSFFiles_DblClick()
    Call cmdOK_Click
End Sub
