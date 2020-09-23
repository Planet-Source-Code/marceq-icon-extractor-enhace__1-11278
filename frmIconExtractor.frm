VERSION 5.00
Begin VB.Form frmIconExtractor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Extractor"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmIconExtractor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMakeSubPaths 
      Caption         =   "&Make Subpaths"
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CheckBox chkRecurse 
      Caption         =   "&Recurse Subdirectories"
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Top             =   2760
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3120
      Width           =   840
   End
   Begin VB.PictureBox TPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   540
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2760
      Width           =   840
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go!"
      Default         =   -1  'True
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtExtract 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdBrowseExtract 
      Caption         =   "..."
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdBrowsePath 
      Caption         =   "..."
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtMessages 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   840
      Width           =   4680
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label 
      Caption         =   "Extract Size"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Save Files To"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path To Search"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path To Extract To"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmIconExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowsePath_Click()
Dim F As Folder
Dim X As Shell
Set X = New Shell
Set F = X.BrowseForFolder(hWnd, "Select a folder", 1)
If Not F Is Nothing Then
    If Right(F.Self.Path, 1) = "\" Then
        txtPath = F.Self.Path
    Else
        txtPath = F.Self.Path & "\"
    End If
End If
Set F = Nothing
Set X = Nothing
End Sub

Private Sub cmdBrowseExtract_Click()
Dim F As Folder
Dim X As Shell
Set X = New Shell
Set F = X.BrowseForFolder(hWnd, "Select a folder", 1)
If Not F Is Nothing Then
    If Right(F.Self.Path, 1) = "\" Then
        txtExtract = F.Self.Path
    Else
        txtExtract = F.Self.Path & "\"
    End If
End If
Set F = Nothing
Set X = Nothing
End Sub

Private Sub cmdGo_Click()
Dim Icons As Long
Static Working  As Boolean
If Working Then Exit Sub
Working = True
MousePointer = vbHourglass
txtMessages = ""
Icons = ExtractIcons(txtPath, txtExtract, cboType = "BMP", cboSize = "Large", chkRecurse = vbChecked, chkMakeSubPaths = vbChecked, TPic, txtMessages)
If ExitCalled Then Exit Sub
txtMessages.SelText = "Total icons found: " & Icons & vbCrLf
MousePointer = vbDefault
Working = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
cboType.AddItem "ICO"
cboType.AddItem "BMP"
cboType.ListIndex = 0
cboSize.AddItem "Large"
cboSize.AddItem "Small"
cboSize.ListIndex = 0
txtPath = "c:\winnt\"
txtExtract = "C:\Documents and Settings\mramirez\My Documents\Iconos\"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ExitCalled = True
End Sub
