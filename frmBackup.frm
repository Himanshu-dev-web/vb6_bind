VERSION 5.00
Begin VB.Form frmbackup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Backup Data"
   ClientHeight    =   5820
   ClientLeft      =   3975
   ClientTop       =   780
   ClientWidth     =   4395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   3240
      ScaleHeight     =   345
      ScaleWidth      =   960
      TabIndex        =   4
      Top             =   5280
      Width           =   990
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   360
         Left            =   -105
         TabIndex        =   11
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2055
      ScaleHeight     =   345
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   5280
      Width           =   1170
      Begin VB.CommandButton Command1 
         Caption         =   "Backup"
         Height          =   360
         Left            =   15
         TabIndex        =   10
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.TextBox txtfolder 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   270
      MaxLength       =   50
      TabIndex        =   2
      Top             =   4785
      Width           =   3840
   End
   Begin VB.DirListBox Dirb 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3330
      Left            =   315
      TabIndex        =   1
      Top             =   1155
      Width           =   3825
   End
   Begin VB.DriveListBox Driveb 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   300
      TabIndex        =   0
      Top             =   780
      Width           =   3840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00C0E0FF&
      Height          =   5715
      Left            =   45
      Top             =   30
      Width           =   4275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Top             =   4500
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   1590
      TabIndex        =   8
      Top             =   4500
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Path ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   270
      TabIndex        =   7
      Top             =   510
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   75
      X2              =   4320
      Y1              =   435
      Y2              =   435
   End
   Begin VB.Label LblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C29221&
      Caption         =   "DATA BACKUP"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   4245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   90
      X2              =   4350
      Y1              =   5190
      Y2              =   5175
   End
   Begin VB.Label lblbackup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   5355
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MDB_Path As String

Private Sub Command1_Click()
BackupData
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Driveb_Change()
On Error GoTo last
    Dirb.Path = Driveb.Drive
    Exit Sub
last:
    MsgBox "Drive not Accessable", vbCritical
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Screen.ActiveControl.TabIndex > -1 Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Me.Move 6800, 1500
'MDB_Path = VB.App.Path + "\" + Trim(main.Directory) + "\data.mdb"
MDB_Path = VB.App.Path + "\data.mdb"
End Sub

Sub BackupData()
'On Error GoTo last
    zoo = InStr(txtfolder.Text, "/")
    If zoo > 0 Then
        MsgBox "Invalid Folder Name", vbCritical
        txtfolder.Text = ""
        txtfolder.SetFocus
        Exit Sub
    End If
    DoEvents
    lblbackup.Visible = True
    Dim bfile As New FileSystemObject
    If Trim(txtfolder.Text) <> "" Then
        If bfile.FolderExists(Replace(Dirb.Path & "\", "\\", "\") & txtfolder.Text) = True Then
            lblbackup.Visible = False
            MsgBox "Folder allready exists. Please change the folder name.", vbCritical
            Exit Sub
        Else
            bfile.CreateFolder Replace(Dirb.Path & "\", "\\", "\") & txtfolder.Text
            bfile.CopyFile MDB_Path, Replace(Dirb.Path & "\", "\\", "\") & txtfolder.Text & "\", True
        End If
    Else
        txtfolder.SetFocus
        lblbackup.Visible = False
        MsgBox "Please Write The Folder Name", vbInformation
        Exit Sub
    End If
    Set bfile = Nothing
lblbackup.Visible = False
txtfolder.Text = ""
MsgBox "Backup completed.", vbInformation
Exit Sub
last:
    lblbackup.Visible = False
    txtfolder.Text = ""
    MsgBox "Error During Backup...", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub L1_Click(Index As Integer)
Unload Me
End Sub

Private Sub p1_GotFocus(Index As Integer)
   ' P1(Index).BackColor = &HC0C0C0
   ' L1(Index).ForeColor = vbBlack
End Sub
Private Sub P1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 And Screen.ActiveControl.TabIndex > -1 Then
        SendKeys "{TAB}"
    End If
    If KeyCode = 37 And Screen.ActiveControl.TabIndex > -1 Then
        SendKeys "+{TAB}"
    End If
    If KeyCode = 13 Then
        Select Case Index
        Case 0
            BackupData
        Case 1
            Unload Me
        End Select
    End If
End Sub
Private Sub P1_LostFocus(Index As Integer)
    'P1(Index).BackColor = vbBlack
    'L1(Index).ForeColor = vbWhite '&HC0C0C0
End Sub
Private Sub txtfolder_LostFocus()
    txtfolder.BackColor = vbBlack
    txtfolder.ForeColor = vbWhite '&HC0C0C0
End Sub

Private Sub Txtfolder_GotFocus()
    txtfolder.BackColor = &HC0C0C0
    txtfolder.ForeColor = vbBlack
End Sub
