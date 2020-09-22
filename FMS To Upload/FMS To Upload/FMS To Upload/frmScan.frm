VERSION 5.00
Begin VB.Form frmScan 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Scan & Upload"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      HelpContextID   =   5
      Left            =   2520
      MouseIcon       =   "frmScan.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdStartScan 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Start Scan"
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
      HelpContextID   =   5
      Left            =   480
      MouseIcon       =   "frmScan.frx":074C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
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
      Left            =   1680
      MaxLength       =   255
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.ComboBox cboFileType 
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
      ItemData        =   "frmScan.frx":0A56
      Left            =   1680
      List            =   "frmScan.frx":0A60
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Type:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1170
      Width           =   855
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
frmMain.SetFocus
End Sub

Private Sub cmdStartScan_Click()
Me.Enabled = False
frmMain.Picture1.Height = 800
frmMain.Picture1.Width = 600
Dim strUploadFileName As String

Dim strFileExt As String
If cboFileType.ListIndex = 0 Then
    strFileExt = ".jpg"
Else
    strFileExt = ".pdf"
End If
    
If Len(strUploadFileName) <= 0 Then strUploadFileName = Format(Now(), "yymmddhhmmss")

On Error GoTo ScnErr
frmMain.Picture1.Picture = frmMain.winCD.ShowAcquireImage.FileData.Picture

On Error GoTo SaveErr
Me.Caption = "Scanned file successfully...Saving file.."
Call SAVEJPEG(App.Path & "\Uploads\" & strUploadFileName & ".jpg", 100, frmMain.Picture1)

frmMain.Enabled = False
With frmFileProperties
    boolFromFileUpload = True
    If strFileExt = ".pdf" Then
        Me.Caption = "Saved file successfully...Converting to PDF.."
        If IMG2PDF(App.Path & "\Uploads\" & strUploadFileName & ".jpg", App.Path & "\Uploads\" & strUploadFileName & ".pdf", True) Then
            Kill App.Path & "\Uploads\" & strUploadFileName & ".jpg"
            frmMain.CommonDialog1.FileName = App.Path & "\Uploads\" & strUploadFileName & ".pdf"
        Else
            Dim confirmMsg As VbMsgBoxResult
            confirmMsg = MsgBox("The PDF conversion failed. Do you want to upload the file as JPG instead ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
            If confirmMsg = vbYes Then
                frmMain.CommonDialog1.FileName = App.Path & "\Uploads\" & strUploadFileName & ".jpg"
            Else
                Exit Sub
            End If
        End If
    Else
        frmMain.CommonDialog1.FileName = App.Path & "\Uploads\" & strUploadFileName & ".jpg"
    End If
    .Show , frmMain
    .lblFileName.Caption = strUploadFileName & strFileExt
    .lblCreatedBy.Caption = strUserName
    .lblCreatedOn.Caption = nowDate & "  " & nowTime
    .lblModifiedBy.Caption = ""
    .lblModifiedOn.Caption = ""
    .lblRevision.Caption = "0"
    .cboPriority.Text = .cboPriority.List(1)
    .DTPInactive.Value = Format(DateAdd("m", 1, Now()), "dd-MMM-yyyy")
    strFilePath = App.Path & "\Uploads\"
End With
frmMain.Picture1.Height = 10
frmMain.Picture1.Width = 10
Me.Enabled = True
Unload Me
Exit Sub

ScnErr:
Me.Enabled = True
MsgBox "File not scanned!!!", vbExclamation
frmMain.Picture1.Height = 10
frmMain.Picture1.Width = 10
Exit Sub

SaveErr:
Me.Enabled = True
MsgBox "Scanned file was not saved successfully", vbExclamation
frmMain.Picture1.Height = 10
frmMain.Picture1.Width = 10
Exit Sub

End Sub

Private Sub Form_Load()
cboFileType.Text = cboFileType.List(0)
txtFileName.Text = Format(Now(), "yymmddhhmmss")
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub txtFileName_GotFocus()
txtFileName.SelStart = 0
txtFileName.SelLength = Len(txtFileName.Text)
End Sub
