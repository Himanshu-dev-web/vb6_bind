VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "User Login"
   ClientHeight    =   2376
   ClientLeft      =   3816
   ClientTop       =   3720
   ClientWidth     =   4032
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2376
   ScaleWidth      =   4032
   Begin VB.ComboBox cboyrs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   180
      Width           =   2175
   End
   Begin VB.CommandButton cmdPass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   570
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   3525
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1500
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   0
      Top             =   660
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   690
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
  If Check1.Value = 1 Then
   txtchange.Visible = True
   Command1.Visible = True
  Else
   txtchange.Visible = False
   Label3.Visible = False
   Label2.Caption = "Password :"
   Command1.Visible = False
  End If
End Sub
Private Sub cmdPass_Click()

    ' Step 1 - Check License first
    Dim str_ As String
    str_ = check_session
    If str_ <> "" Then
        MsgBox " " & str_ & "", vbCritical
    End If

    ' Step 2 - Open Database Connection
    Call DSN

    ' Step 3 - Check Connection
    If con.State = 0 Then
        MsgBox "Database connection failed!", vbCritical
        Exit Sub
    End If

    ' Step 4 - Login Check
    If rs.State = 1 Then rs.Close
    rs.Open "select * from loginpassword where (usrpassword='" & txtPass.text & "') and usrname='" & txtUser.text & "'", con, adOpenDynamic, adLockOptimistic

    If rs.EOF = False Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from loginpassword where usrname='" & txtUser.text & "' and usrpassword='" & txtPass.text & "'", con, adOpenDynamic, adLockOptimistic

        If rs.EOF = False Then
            If rs.Fields("usrname") = "admin" Then
                frmMain.mnureset.Enabled = True
                frmMain.Show
            Else
                frmMain.mnureset.Enabled = False
                frmMain.Show
            End If
        End If
        Unload Me
    Else
        MsgBox "Please check the Password !!", vbCritical
        Exit Sub
    End If

End Sub

Sub TransictionDetails()
   
   If rs.State = 1 Then rs.Close
   rs.Open "select * from LoginDetail", con, adOpenDynamic
   rs.AddNew
   rs!usrname = txtUser.text
   rs!Dates = Date
   rs!Work = "Null"
   rs.Update
   
End Sub



Private Sub txtchange_GotFocus()
   txtchange.BackColor = &HFFFF80
End Sub
Private Sub txtchange_LostFocus()
 txtchange.BackColor = &HFFFFFF
End Sub

Private Sub Form_Load()

    Dim direct As New ADODB.Connection
    direct.CursorLocation = adUseClient
    With direct
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\info.mdb"
        .Open
    End With

    Set rs = New ADODB.Recordset
    rs.Open "select * from datadir order by year desc", direct, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Do While Not rs.EOF
            cboyrs.AddItem rs!Directory
            If Not rs.EOF Then rs.MoveNext
        Loop
        cboyrs.ListIndex = 0
    End If

    rs.Close      ' <<< IMPORTANT - rs close karo
    direct.Close  ' <<< IMPORTANT - direct connection close karo

    ch_ = True
    chk_done = False

End Sub

Private Sub txtPass_GotFocus()
  txtPass.BackColor = &HFFFF80
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call cmdPass_Click
  End If
End Sub
Private Sub txtPass_LostFocus()
 txtPass.BackColor = &HFFFFFF
End Sub

Private Sub txtUser_GotFocus()
txtUser.BackColor = &HFFFF80
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtPass.SetFocus
End Sub
Private Sub txtUser_LostFocus()
  txtUser.BackColor = &HFFFFFF
End Sub
