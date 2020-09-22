VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileProperties 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File Properties"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   Icon            =   "frmFileProperties.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "File Properties"
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
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.TextBox txtDescription 
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
         ForeColor       =   &H80000009&
         Height          =   320
         Left            =   2400
         MaxLength       =   255
         TabIndex        =   14
         Top             =   2160
         Width           =   8055
      End
      Begin VB.ComboBox cboPriority 
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
         ItemData        =   "frmFileProperties.frx":0442
         Left            =   2880
         List            =   "frmFileProperties.frx":044F
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Cancel"
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
         Left            =   9840
         MouseIcon       =   "frmFileProperties.frx":0466
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7320
         Width           =   1695
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
         Left            =   7920
         MouseIcon       =   "frmFileProperties.frx":0770
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   7320
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveUsers 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Remove User >>"
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
         Left            =   8760
         MouseIcon       =   "frmFileProperties.frx":0A7A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddUsers 
         BackColor       =   &H00C0FFFF&
         Caption         =   "<< &Add User"
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
         Left            =   8760
         MouseIcon       =   "frmFileProperties.frx":0D84
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3120
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPInactive 
         Height          =   345
         Left            =   2880
         TabIndex        =   28
         Top             =   7320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
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
         Format          =   53018627
         CurrentDate     =   38496
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDocCopies 
         Height          =   1575
         Left            =   7920
         TabIndex        =   24
         Top             =   4920
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   32896
         ForeColorFixed  =   -2147483624
         BackColorSel    =   8388608
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   3
         BorderStyle     =   0
         FormatString    =   "User             | Date           | Time      "
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
      Begin MSFlexGridLib.MSFlexGrid MSFDocViews 
         Height          =   1575
         Left            =   240
         TabIndex        =   20
         Top             =   4920
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   32896
         ForeColorFixed  =   -2147483624
         BackColorSel    =   8388608
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   3
         BorderStyle     =   0
         FormatString    =   "User             | Date           | Time      "
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
      Begin MSFlexGridLib.MSFlexGrid MSFDocPrints 
         Height          =   1575
         Left            =   4080
         TabIndex        =   22
         Top             =   4920
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   32896
         ForeColorFixed  =   -2147483624
         BackColorSel    =   8388608
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   3
         BorderStyle     =   0
         FormatString    =   "User             | Date           | Time      "
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
      Begin MSFlexGridLib.MSFlexGrid MSFActiveUsers 
         Height          =   1455
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   32896
         ForeColorFixed  =   -2147483624
         BackColorSel    =   8388608
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   3
         BorderStyle     =   0
         FormatString    =   "Department                              |User                                      |  View  |Print   |Edit   |Copy  "
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
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
         Left            =   9480
         TabIndex        =   3
         Top             =   420
         Width           =   420
      End
      Begin VB.Label lblRevision 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   9960
         TabIndex        =   4
         Top             =   420
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Priority:"
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
         Left            =   2040
         TabIndex        =   25
         Top             =   6960
         Width           =   660
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Active Users for this file:"
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
         TabIndex        =   15
         Top             =   2640
         Width           =   2115
      End
      Begin VB.Label lblModifiedOn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label lblModifiedBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   1500
         Width           =   510
      End
      Begin VB.Label lblCreatedOn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   1140
         Width           =   510
      End
      Begin VB.Label lblCreatedBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   2400
         TabIndex        =   6
         Top             =   780
         Width           =   510
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   2400
         TabIndex        =   2
         Top             =   420
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description / Keywords:"
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
         TabIndex        =   13
         Top             =   2235
         Width           =   2055
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Copied By:"
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
         Left            =   7920
         TabIndex        =   23
         Top             =   4680
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Printed By:"
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
         Left            =   4080
         TabIndex        =   21
         Top             =   4680
         Width           =   1860
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Viewed By:"
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
         TabIndex        =   19
         Top             =   4680
         Width           =   1875
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified On:"
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
         TabIndex        =   11
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified By:"
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
         TabIndex        =   9
         Top             =   1500
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Created On:"
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
         TabIndex        =   7
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Created By:"
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
         TabIndex        =   5
         Top             =   780
         Width           =   1005
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
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File becomes inactive after:"
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
         Left            =   360
         TabIndex        =   27
         Top             =   7440
         Width           =   2370
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   8130
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12331
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
            TextSave        =   "1:19 PM"
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
End
Attribute VB_Name = "frmFileProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddUsers_Click()
frmSelectUsers.Show , Me
Me.Enabled = False
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If MSFActiveUsers.Rows <= 2 Then
    boolConfirmDelete = MsgBox("Are you sure you want to exit without adding any users ?" & vbCrLf & "To add users click YES or " & vbCrLf & "To continue without adding users click NO.", vbYesNo + vbQuestion)
    If boolConfirmDelete = vbYes Then
        frmSelectUsers.Show , Me
        Me.Enabled = False
        Exit Sub
    End If
End If
If DateDiff("d", Format(CDate(nowDate), "dd-MMM-yyyy"), Format(DTPInactive.Value, "dd-MMM-yyyy")) <= 0 Then
    MsgBox "Please specify an expiry date atleast 1 day or more than the date created.", vbExclamation
    Exit Sub
End If

If boolFromRevision = True Then
    strFilePath = Replace(frmMain.CommonDialog1.FileName, GetFileTitleFromPath(frmMain.CommonDialog1.FileName), "", , , vbBinaryCompare)
    If chkFilePath(frmMain.CommonDialog1.FileName) = True Then
        Call uploadNewFile
        Call DeleteFile(frmMain.CommonDialog1.FileName)
        MsgBox "File uploaded successfully", vbInformation
    Else
        MsgBox "The file specified by you does not exist.", vbExclamation
    End If
ElseIf cmdAddUsers.Enabled = True And boolFromFileUpload = True Then
    Call uploadNewFile
    MsgBox "File uploaded successfully", vbInformation
ElseIf cmdAddUsers.Enabled = True And boolFromFileUpload = False Then
    Call uploadFileProperties
    MsgBox "File uploaded successfully", vbInformation
End If

Call frmMain.Form_Activate
End Sub

Private Sub cmdRemoveUsers_Click()
If Len(MSFActiveUsers.Text) > 0 Then
    MSFActiveUsers.Col = 1
    boolConfirmDelete = MsgBox("Are you sure you want to remove this user " & MSFActiveUsers.Text & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirmDelete = vbYes Then
        MSFActiveUsers.RemoveItem MSFActiveUsers.Row
    End If
End If
End Sub

Private Sub Form_Load()
cboPriority.Text = cboPriority.List(1)

If Len(strDocID) > 0 And boolFromFileUpload = False Then
    Call setActiveUsers
    MSFActiveUsers.Rows = MSFActiveUsers.Rows + 1
    
    Set rsUsers = DBFileManager.OpenRecordset("select distinct User,Date,Time from [Documents ActivityLog] where Type='View' and [Document Id]=" & strDocID)
    Call loadRSUserActivity(MSFDocViews, rsUsers)
    
    Set rsUsers = DBFileManager.OpenRecordset("select distinct User,Date,Time from [Documents ActivityLog] where Type='Print' and [Document Id]=" & strDocID)
    Call loadRSUserActivity(MSFDocPrints, rsUsers)
    
    Set rsUsers = DBFileManager.OpenRecordset("select distinct User,Date,Time from [Documents ActivityLog] where Type='Copy' and [Document Id]=" & strDocID)
    Call loadRSUserActivity(MSFDocCopies, rsUsers)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.SetFocus
End Sub

