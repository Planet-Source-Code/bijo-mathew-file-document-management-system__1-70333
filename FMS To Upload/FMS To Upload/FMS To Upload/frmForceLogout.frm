VERSION 5.00
Begin VB.Form frmForceLogout 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Force Logout"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   HelpContextID   =   22
   Icon            =   "frmForceLogout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.ComboBox cboUserName 
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
         ItemData        =   "frmForceLogout.frx":0442
         Left            =   3000
         List            =   "frmForceLogout.frx":0444
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   550
         Width           =   3015
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
         Left            =   3480
         MouseIcon       =   "frmForceLogout.frx":0446
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdForceLogOut 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Force Logout"
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
         Left            =   1440
         MouseIcon       =   "frmForceLogout.frx":0750
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name to Force Logout:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmForceLogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdForceLogOut_Click()
boolConfirmDelete = MsgBox("Are you sure you want to force logout on " & cboUserName.Text & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
If boolConfirmDelete = vbYes Then
    DBFileManager.Execute ("Delete * from Logins where Login='" & cboUserName.Text & "'")
    
    Call Form_Load
End If
End Sub

Private Sub Form_Load()
cboUserName.Clear
Set rsUserForceLogout = DBFileManager.OpenRecordset("Select Login from Logins where Login<>'Admin'")
With rsUserForceLogout
If .EOF = False Then
    .MoveFirst
    While .EOF = False
        cboUserName.AddItem rsUserForceLogout!Login
        .MoveNext
    Wend
    cboUserName.Text = cboUserName.List(0)
Else
    cmdForceLogOut.Enabled = False
End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.SetFocus
End Sub

