VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDBase 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Configure Database"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   HelpContextID   =   250
   Icon            =   "frmDBase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Find the file ""File Manager.mdb"""
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
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Browse..."
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
         Left            =   6000
         MouseIcon       =   "frmDBase.frx":0442
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Browse for database"
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&OK"
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
         Left            =   1320
         MouseIcon       =   "frmDBase.frx":074C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Set database path"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Cancel"
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
         Left            =   3240
         MouseIcon       =   "frmDBase.frx":0A56
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit"
         Top             =   1800
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6240
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtBrowse 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Enter database path here"
         Top             =   600
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmDBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
CommonDialog1.DialogTitle = "Find the Microsoft Access file File Manager.mdb"
CommonDialog1.Filter = "File Manager.mdb|File Manager.mdb"
CommonDialog1.InitDir = App.Path
CommonDialog1.Action = 1

If CommonDialog1.FileName <> "" Then
    txtBrowse.Text = CommonDialog1.FileName
 ElseIf CommonDialog1.FileName = "" Then
    MsgBox "Please specify the database path.", vbCritical
    cmdBrowse.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
dbPathString = txtBrowse.Text
SaveSetting App.EXEName, "DB Path", "DB Path", dbPathString
On Error GoTo dbError
Set DBFileManager = OpenDatabase(dbPathString)
frmLogin.Show
Unload Me
Exit Sub

dbError:
MsgBox "Please configure your database again.", vbCritical
frmDBase.Show
End Sub

Private Sub Form_Load()
Me.MousePointer = vbNormal
End Sub
