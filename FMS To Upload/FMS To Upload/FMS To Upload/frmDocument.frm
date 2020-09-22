VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmDocument 
   BackColor       =   &H00800000&
   Caption         =   " Document"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmDocument.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRight 
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   7
      Top             =   10080
      Width           =   735
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "ç"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   6
      Top             =   10080
      Width           =   855
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14880
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14880
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   120
      Top             =   960
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Print"
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
      Left            =   5760
      MouseIcon       =   "frmDocument.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10200
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
      Left            =   7680
      MouseIcon       =   "frmDocument.frx":074C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   9975
      Left            =   120
      ScaleHeight     =   9915
      ScaleWidth      =   14715
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   10215
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   14985
         ExtentX         =   26432
         ExtentY         =   18018
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
On Error Resume Next
Picture1.Enabled = True
WebBrowser1.SetFocus
Call keybd_event(VK_CONTROL, 0, 0, 0)   'for CTRL press
Call keybd_event(80, 0, 0, 0)           'for P press
Call keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
Call updateActivityTable("Print")
End Sub

Private Sub cmdRight_Click()
On Error Resume Next
Picture1.Enabled = True
WebBrowser1.SetFocus
SendKeys "{RIGHT 20}", False
Timer3.Enabled = True
End Sub

Private Sub cmdLeft_Click()
On Error Resume Next
Timer3.Enabled = True
Picture1.Enabled = True
WebBrowser1.SetFocus
SendKeys "{LEFT 20}", False
Timer3.Enabled = True
End Sub

Private Sub cmdUp_Click()
On Error Resume Next
Timer3.Enabled = True
Picture1.Enabled = True
WebBrowser1.SetFocus
'SendKeys "{PGUP}", False
SendKeys "{PGUP}", False
Timer3.Enabled = True
End Sub

Private Sub cmdDown_Click()
On Error Resume Next
Timer3.Enabled = True
Picture1.Enabled = True
WebBrowser1.SetFocus
'SendKeys "{PGDN 25}", True
SendKeys "{PGDN}", False
Timer3.Enabled = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub Timer3_Timer()
Picture1.Enabled = False
Timer3.Enabled = False
End Sub

'Private Sub Timer1_Timer()
'On Error Resume Next
'If Me.WindowState <> 1 Then
'    Clipboard.Clear
''    Call keybd_event(VK_ESCAPE, 0, 0, 0)    'for ESCAPE press
'End If
'End Sub

