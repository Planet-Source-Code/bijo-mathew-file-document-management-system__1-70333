VERSION 5.00
Begin VB.Form frmDefaultUserRights 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " User Rights"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   Icon            =   "frmDefaultUserRights.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Set Default User Rights"
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
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
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
         Left            =   2280
         MouseIcon       =   "frmDefaultUserRights.frx":0442
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
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
         Left            =   720
         MouseIcon       =   "frmDefaultUserRights.frx":074C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkCopy 
         BackColor       =   &H00800000&
         Caption         =   "&Copy Out"
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
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkEdit 
         BackColor       =   &H00800000&
         Caption         =   "&Download for Editing"
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
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox chkPrint 
         BackColor       =   &H00800000&
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
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkView 
         BackColor       =   &H00800000&
         Caption         =   "&View"
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
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmDefaultUserRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkCopy_Click()
If chkCopy.Value = 1 Then
    chkView.Value = 1
End If
End Sub

Private Sub chkEdit_Click()
If chkEdit.Value = 1 Then
    chkView.Value = 1
End If
End Sub

Private Sub chkPrint_Click()
If chkPrint.Value = 1 Then
    chkView.Value = 1
End If
End Sub

Private Sub chkView_Validate(Cancel As Boolean)
If chkPrint.Value = 1 Or chkCopy.Value = 1 Or chkEdit.Value Then
    chkView.Value = 1
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If chkCopy.Value = 0 And chkView.Value = 0 And chkPrint.Value = 0 And chkEdit.Value = 0 Then
    MsgBox "Please add at least 1 right to the user to add.", vbExclamation
    Exit Sub
End If

With frmSetUserRights
    lngX = 1
    While lngX < .MSFActiveUsers.Rows
        .MSFActiveUsers.Row = lngX
        .MSFActiveUsers.Col = 0
        If UCase(.MSFActiveUsers.Text) = UCase(strUserDeptName) Then
            .MSFActiveUsers.Col = 1
            If UCase(.MSFActiveUsers.Text) = UCase(strUserFileName) Then
                .MSFActiveUsers.Col = 2
                If chkView.Value = 1 Then
                    .MSFActiveUsers.Text = "Y"
                Else
                    .MSFActiveUsers.Text = "N"
                End If
                
                .MSFActiveUsers.Col = 3
                If chkPrint.Value = 1 Then
                    .MSFActiveUsers.Text = "Y"
                Else
                    .MSFActiveUsers.Text = "N"
                End If
                
                .MSFActiveUsers.Col = 4
                If chkEdit.Value = 1 Then
                    .MSFActiveUsers.Text = "Y"
                Else
                    .MSFActiveUsers.Text = "N"
                End If
                
                .MSFActiveUsers.Col = 5
                If chkCopy.Value = 1 Then
                    .MSFActiveUsers.Text = "Y"
                Else
                    .MSFActiveUsers.Text = "N"
                End If
                Unload Me
                Exit Sub
            End If
        End If
        lngX = lngX + 1
    Wend
    
    
    .MSFActiveUsers.Rows = .MSFActiveUsers.Rows + 1
    .MSFActiveUsers.Row = .MSFActiveUsers.Rows - 2
    
    .MSFActiveUsers.Col = 0
    .MSFActiveUsers.Text = strUserDeptName
    
    .MSFActiveUsers.Col = 1
    .MSFActiveUsers.Text = strUserFileName
    
    .MSFActiveUsers.Col = 2
    If chkView.Value = 1 Then
        .MSFActiveUsers.Text = "Y"
    Else
        .MSFActiveUsers.Text = "N"
    End If
    
    .MSFActiveUsers.Col = 2
    If chkView.Value = 1 Then
        .MSFActiveUsers.Text = "Y"
    Else
        .MSFActiveUsers.Text = "N"
    End If
    
    .MSFActiveUsers.Col = 3
    If chkPrint.Value = 1 Then
        .MSFActiveUsers.Text = "Y"
    Else
        .MSFActiveUsers.Text = "N"
    End If
    
    .MSFActiveUsers.Col = 4
    If chkEdit.Value = 1 Then
        .MSFActiveUsers.Text = "Y"
    Else
        .MSFActiveUsers.Text = "N"
    End If
    
    .MSFActiveUsers.Col = 5
    If chkCopy.Value = 1 Then
        .MSFActiveUsers.Text = "Y"
    Else
        .MSFActiveUsers.Text = "N"
    End If
    
    'to avoid that extra blank line at end
    '.MSFActiveUsers.Rows = .MSFActiveUsers.Rows - 1
End With
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdOK_Click
End If
End Sub

Private Sub Form_Load()
chkCopy.Value = 0
chkEdit.Value = 0
chkPrint.Value = 0
chkView.Value = 0

Set rsDefaultUserRight = DBFileManager.OpenRecordset("Select View,Print,Edit,Copy from [Default User Rights] where [Set User Name]='" & frmSetUserRights.MSFUsers.Text & "' and [User Name]='" & strUserName & "'")
If rsDefaultUserRight.EOF = False Then
    rsDefaultUserRight.MoveFirst
    If UCase(rsDefaultUserRight![View]) = "Y" Then
        chkView.Value = 1
    End If
    
    If UCase(rsDefaultUserRight![Print]) = "Y" Then
        chkPrint.Value = 1
    End If
    
    If UCase(rsDefaultUserRight![Edit]) = "Y" Then
        chkEdit.Value = 1
    End If
    
    If UCase(rsDefaultUserRight![Copy]) = "Y" Then
        chkCopy.Value = 1
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmSetUserRights.Enabled = True
frmSetUserRights.SetFocus
End Sub

