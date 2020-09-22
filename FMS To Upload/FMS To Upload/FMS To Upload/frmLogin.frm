VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Login"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6645
   HelpContextID   =   3
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4215
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin VB.CheckBox chkRememberPass 
         BackColor       =   &H80000007&
         Caption         =   "Remember my Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox chkRememberUN 
         BackColor       =   &H80000007&
         Caption         =   "Remember my User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   255
         PasswordChar    =   "¤"
         TabIndex        =   4
         ToolTipText     =   "Enter password here"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtLoginName 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   2
         ToolTipText     =   "Enter login name here"
         Top             =   480
         Width           =   3015
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   615
         Left            =   3120
         TabIndex        =   8
         ToolTipText     =   "Cancel"
         Top             =   3240
         Width           =   1815
         ForeColor       =   65280
         BackColor       =   0
         Caption         =   "Cancel    "
         PicturePosition =   196613
         Size            =   "3201;1085"
         MousePointer    =   99
         Picture         =   "frmLogin.frx":0442
         Accelerator     =   67
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdLogin 
         Height          =   615
         Left            =   960
         TabIndex        =   7
         ToolTipText     =   "Login"
         Top             =   3240
         Width           =   1815
         ForeColor       =   65280
         BackColor       =   0
         Caption         =   "Login"
         PicturePosition =   196613
         Size            =   "3201;1085"
         MousePointer    =   99
         Accelerator     =   76
         MouseIcon       =   "frmLogin.frx":0894
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   870
         TabIndex        =   3
         Top             =   1155
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   750
         TabIndex        =   1
         Top             =   540
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4980
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2673
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
            TextSave        =   "1:32 PM"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "With 128 bit encryption."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2160
      TabIndex        =   10
      Top             =   4680
      Width           =   2475
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
      Begin VB.Menu mnuLogin 
         Caption         =   "&Login   "
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit   "
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help   "
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Help Contents"
      End
      Begin VB.Menu sepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout   "
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "Cannot start two instances of File Manager from the same system.", vbExclamation
    Unload Me
Else
    On Error Resume Next
    App.HelpFile = App.Path & "\Help\FMS.hlp"
    Screen.MousePointer = vbHourglass
    Call setDbPath
    On Error GoTo err:
    Set DBFileManager = OpenDatabase(dbPathString)
    
   
    If GetSetting(App.EXEName, "Auto", "Auto User Name") = "True" Then
        txtLoginName.Text = GetSetting(App.EXEName, "Auto", "User Name")
        chkRememberUN.Value = 1
    Else
        chkRememberUN.Value = 0
    End If
    
    If GetSetting(App.EXEName, "Auto", "Auto Password") = "True" Then
        txtPassword.Text = GetSetting(App.EXEName, "Auto", "Password")
        chkRememberPass.Value = 1
    Else
        chkRememberPass.Value = 0
    End If
    
    strServerPath = Replace(dbPathString, "Database\File Manager.mdb", "")
    strServerPath = strServerPath & "Files\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Screen.MousePointer = vbNormal
    Call chkFolderPath(App.Path & "\Uploads", True)
    Call chkFolderPath(App.Path & "\Checkout", True)
    
    'IN Debug Mode delete all logins
    If inDebugMode = True Then DBFileManager.Execute ("Delete * from Logins")
   
End If
    
On Error GoTo err1:
If FileVersionChk(Replace(strServerPath, "\Files\", "\Updates\FMS.exe")) > FileVersionChk(App.Path & "\FMS.exe") Then
    boolConfirmDelete = MsgBox("There is a new update for FMS available in the server. Do you like to download it ?", vbYesNoCancel + vbDefaultButton1 + vbQuestion)
    If boolConfirmDelete = vbYes Then
        On Error GoTo err1:
        Set fsFile = FSO.CreateTextFile(App.Path & "\Update.ini", True)
        fsFile.WriteLine Replace(strServerPath, "\Files\", "\Updates\FMS.exe")
        chkLicenseLog = False
        Unload frmMain
        Unload frmLogin
        On Error Resume Next
        retVal = Shell(App.Path & "\Auto Update.exe", vbHide)
    End If
End If

Exit Sub

err:
Unload Me
frmDBase.Show
Exit Sub

err1:
End Sub

Private Sub cmdLogin_Click()
Screen.MousePointer = vbHourglass
    Call validateLoginPassword
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call mnuExit_Click
End Sub

Private Sub mnuAbout_Click()
frmAboutSoft.Show , Me
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHelpContents_Click()
SendKeys "{F1}"
End Sub

Private Sub mnuLogin_Click()
Call cmdLogin_Click
End Sub

Private Sub txtLoginName_GotFocus()
txtLoginName.SelStart = 0
txtLoginName.SelLength = Len(txtLoginName.Text)
End Sub

Private Sub txtLoginName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdLogin_Click
End If
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdLogin_Click
End If
End Sub

