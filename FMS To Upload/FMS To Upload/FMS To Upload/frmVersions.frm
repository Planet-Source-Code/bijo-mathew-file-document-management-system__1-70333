VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmVersions 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Document Versions"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   HelpContextID   =   12
   Icon            =   "frmVersions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5160
      MouseIcon       =   "frmVersions.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8520
      Width           =   1695
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
      Left            =   11280
      MouseIcon       =   "frmVersions.frx":074C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   255
      Width           =   450
   End
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
      ItemData        =   "frmVersions.frx":0A56
      Left            =   9375
      List            =   "frmVersions.frx":0A66
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFFiles 
      Height          =   7695
      Left            =   240
      TabIndex        =   3
      Top             =   600
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
      FormatString    =   $"frmVersions.frx":0A91
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
      TabIndex        =   6
      Top             =   9150
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12066
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
            TextSave        =   "9:51 PM"
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
      TabIndex        =   5
      Top             =   360
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
      Left            =   8640
      TabIndex        =   0
      Top             =   330
      Width           =   690
   End
End
Attribute VB_Name = "frmVersions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSortBy_Change()
MSFFiles.Row = 1
If MSFFiles.Rows >= 4 Then
    If cboSortBy.Text = "File Name" Then
        MSFFiles.Col = 0
    ElseIf cboSortBy.Text = "Revision" Then
        MSFFiles.Col = 2
    ElseIf cboSortBy.Text = "User" Then
        MSFFiles.Col = 3
    ElseIf cboSortBy.Text = "" Then
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

lblNoOfFiles.Caption = MSFFiles.Rows - 2 & " File(s)"

Dim intX As Integer
For intX = 0 To MSFFiles.Cols - 1
    MSFFiles.ColAlignment(intX) = 1
Next
End Sub

Private Sub cboSortBy_Click()
Call cboSortBy_Change
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdSort_Click()
If cmdSort.Caption = "á" Then
    cmdSort.Caption = "â"
Else
    cmdSort.Caption = "á"
End If
Call cboSortBy_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.SetFocus
End Sub
