VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSetUserRights 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Set User Rights"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   HelpContextID   =   15
   Icon            =   "frmSetUserRights.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Select Users"
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
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
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
         Width           =   3495
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
         Left            =   3960
         MaxLength       =   100
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton cmdAddUser 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Add User"
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
         Left            =   2760
         MouseIcon       =   "frmSetUserRights.frx":0442
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4080
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
         Left            =   3720
         MouseIcon       =   "frmSetUserRights.frx":074C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7560
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveUser 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Remove User"
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
         Left            =   2760
         MouseIcon       =   "frmSetUserRights.frx":0A56
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6960
         Width           =   1695
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Save"
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
         Left            =   1800
         MouseIcon       =   "frmSetUserRights.frx":0D60
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7560
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFUsers 
         Height          =   2295
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   6930
         _ExtentX        =   12224
         _ExtentY        =   4048
         _Version        =   393216
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
         FormatString    =   "Department                                           |User                                              "
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
         Height          =   1935
         Left            =   240
         TabIndex        =   8
         Top             =   4800
         Width           =   6930
         _ExtentX        =   12224
         _ExtentY        =   3413
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
         FormatString    =   "Department                   |User                          |  View  |Print   |Edit   |Copy  "
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available Users:"
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
         TabIndex        =   4
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type the department or user name to search:"
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
         Width           =   3840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Users:"
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
         Top             =   4560
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   8370
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4314
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
            TextSave        =   "10:14 PM"
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
End
Attribute VB_Name = "frmSetUserRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddUser_Click()
MSFUsers.Col = 1
If UCase(MSFUsers.Text) = UCase(strUserName) Then
    MsgBox "You cannot add yourself as a user.", vbExclamation
ElseIf Len(MSFUsers.Text) > 0 Then
    frmDefaultUserRights.Show , Me
    Me.Enabled = False
    frmDefaultUserRights.Frame1.Caption = "Rights for: " & MSFUsers.Text
    strUserFileName = MSFUsers.Text
    MSFUsers.Col = 0
    strUserDeptName = MSFUsers.Text
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRemoveUser_Click()
If Len(MSFActiveUsers.Text) > 0 Then
    MSFActiveUsers.Col = 1
    boolConfirmDelete = MsgBox("Are you sure you want to remove this user " & MSFActiveUsers.Text & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirmDelete = vbYes Then
        MSFActiveUsers.RemoveItem MSFActiveUsers.Row
    End If
End If
End Sub

Private Sub cmdSave_Click()
With frmSetUserRights
DBFileManager.Execute ("Delete * from [Default User Rights] where [User Name]='" & strUserName & "'")

Set rsUsers = DBFileManager.OpenRecordset("Select Department,[Set User Name],View,Print,Edit,Copy,[User Name] from [Default User Rights]")
Dim intX As Integer
For intX = 1 To .MSFActiveUsers.Rows - 2
    DBFileManager.Execute ("insert into [Default User Rights] (Department,[Set User Name],View,Print,Edit,Copy,[User Name]) values ('" & _
        MSFActiveUsers.TextMatrix(intX, 0) & "','" & MSFActiveUsers.TextMatrix(intX, 1) & "','" & .MSFActiveUsers.TextMatrix(intX, 2) & "','" & _
        .MSFActiveUsers.TextMatrix(intX, 3) & "','" & .MSFActiveUsers.TextMatrix(intX, 4) & "','" & _
        .MSFActiveUsers.TextMatrix(intX, 5) & "','" & strUserName & "')")
Next
Unload Me
End With
End Sub

Private Sub Form_Load()
Set rsUsers = DBFileManager.OpenRecordset("select Department,Login from Master_Users where Login<>'Admin' order by Department,Login")
Call loadRSFlexiValues(MSFUsers, rsUsers)

Set rsUsers = DBFileManager.OpenRecordset("Select Department,[Set User Name],View,Print,Edit,Copy from [Default User Rights] where [User Name]='" & strUserName & "' order by Department,[Set User Name]")
Call loadRSFlexiValues(MSFActiveUsers, rsUsers)

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub MSFActiveUsers_LostFocus()
MSFActiveUsers.Col = 0
MSFActiveUsers.ColSel = MSFActiveUsers.Cols - 1
End Sub

Private Sub MSFUsers_DblClick()
Call cmdAddUser_Click
End Sub

Private Sub MSFUsers_LostFocus()
MSFUsers.Col = 0
MSFUsers.ColSel = MSFUsers.Cols - 1
End Sub

Private Sub txtDepartment_Change()
Set rsUsers = DBFileManager.OpenRecordset("select Department,Login from Master_Users where department like '*" & Replace(txtDepartment.Text, "'", "''") & "*' and login like '*" & Replace(txtUser.Text, "'", "''") & "*' and Login<>'Admin' order by Department,Login")
Call loadRSFlexiValues(MSFUsers, rsUsers)
End Sub

Private Sub txtUser_Change()
Set rsUsers = DBFileManager.OpenRecordset("select Department,Login from Master_Users where department like '*" & Replace(txtDepartment.Text, "'", "''") & "*' and login like '*" & Replace(txtUser.Text, "'", "''") & "*' and Login<>'Admin' order by Department,Login")
Call loadRSFlexiValues(MSFUsers, rsUsers)
End Sub

