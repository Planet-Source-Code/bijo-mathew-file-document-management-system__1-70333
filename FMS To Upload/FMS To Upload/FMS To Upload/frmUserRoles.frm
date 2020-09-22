VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserRoles 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " User Roles"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7440
   HelpContextID   =   18
   Icon            =   "frmUserRoles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "User Role Description"
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
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox cboDepartment 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3000
         Width           =   3255
      End
      Begin VB.CheckBox ckhMaskPassword 
         BackColor       =   &H00800000&
         Caption         =   "&Mask Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3600
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H80000018&
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MouseIcon       =   "frmUserRoles.frx":0442
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Update existing information"
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H80000018&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         MouseIcon       =   "frmUserRoles.frx":074C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete user details"
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000018&
         Caption         =   "&Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         MouseIcon       =   "frmUserRoles.frx":0A56
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Add new user"
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000018&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         MouseIcon       =   "frmUserRoles.frx":0D60
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Exit to Main"
         Top             =   4200
         Width           =   1455
      End
      Begin VB.ComboBox cmbUserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtNewPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   6
         ToolTipText     =   "Enter login name here"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtConfirmPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "Enter login name here"
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox txtOldPassword 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Enter login name here"
         Top             =   1200
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1080
         TabIndex        =   9
         Top             =   3120
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   510
         TabIndex        =   7
         Top             =   2475
         Width           =   1785
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   840
         TabIndex        =   5
         Top             =   1875
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   555
         TabIndex        =   3
         Top             =   1275
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1200
         TabIndex        =   1
         Top             =   660
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   5070
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4076
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
            TextSave        =   "12:06 AM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "3/10/2008"
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
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10920
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   10920
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options   "
      Begin VB.Menu mnuAddUser 
         Caption         =   "&Add User   "
      End
      Begin VB.Menu mnuUpdateUser 
         Caption         =   "&Update User   "
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "&Delete User   "
      End
      Begin VB.Menu sepa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaskPassword 
         Caption         =   "&Mask Password   "
         Checked         =   -1  'True
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help   "
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Help Contents"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout   "
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "frmUserRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ckhMaskPassword_Click()
If ckhMaskPassword.Value = 1 Then
    txtOldPassword.PasswordChar = "*"
    txtNewPassword.PasswordChar = "*"
    txtConfirmPassword.PasswordChar = "*"
ElseIf ckhMaskPassword.Value = 0 Then
    txtOldPassword.PasswordChar = ""
    txtNewPassword.PasswordChar = ""
    txtConfirmPassword.PasswordChar = ""
End If
End Sub

Private Sub cmbUserName_Change()
Call changeUserRoleData
End Sub

Private Sub cmbUserName_Click()
Call changeUserRoleData
End Sub

Private Sub cmbUserName_LostFocus()
If Len(cmbUserName.Text) > 255 Then
    MsgBox "User name should be less than 255 characters!!!"
    cmbUserName.SetFocus
End If
End Sub

Private Sub cmdAdd_Click()
If chkNoOfLicenses = True Then
    If userRoleValidateControls = True Then
        If InStr(1, cmbUserName.Text, "'") = 0 Then
            Call addUserRoleData
        Else
            MsgBox "User name cannot conatain the special character quote('). Please remove the quote from user name and continue.", vbExclamation
        End If
    End If
    Call loadUserRoleComboBoxes
End If
End Sub

Private Sub cmdDelete_Click()
Call deleteUser
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
If userRoleValidateControls = True Then
    Call updateUserRoleData
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Form_Load()
Call loadUserRoleComboBoxes
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub mnuAbout_Click()
frmAboutSoft.Show , Me
End Sub

Private Sub mnuAddUser_Click()
Call cmdAdd_Click
End Sub

Private Sub mnuDeleteUser_Click()
Call cmdDelete_Click
End Sub

Private Sub mnuExit_Click()
Call cmdExit_Click
End Sub

Private Sub mnuHelpContents_Click()
SendKeys "{F1}"
End Sub

Private Sub mnuMaskPassword_Click()
If ckhMaskPassword.Value = 1 Then
    ckhMaskPassword.Value = 0
ElseIf ckhMaskPassword.Value = 0 Then
    ckhMaskPassword.Value = 1
End If
End Sub

Private Sub mnuUpdateUser_Click()
Call cmdUpdate_Click
End Sub
