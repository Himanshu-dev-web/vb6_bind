VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAdmin 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   570
   ClientWidth     =   11880
   Icon            =   "frmAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab adminTab 
      Height          =   5175
      Left            =   3510
      TabIndex        =   20
      Top             =   1500
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9128
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "New User"
      TabPicture(0)   =   "frmAdmin.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Delete User"
      TabPicture(1)   =   "frmAdmin.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Change Profile"
      TabPicture(2)   =   "frmAdmin.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   -75000
         TabIndex        =   23
         Top             =   330
         Width           =   5520
         Begin VB.CheckBox chkRead2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2160
            TabIndex        =   5
            Top             =   3120
            Width           =   840
         End
         Begin VB.CheckBox chkWrite2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3015
            TabIndex        =   6
            Top             =   3120
            Width           =   855
         End
         Begin VB.CheckBox chkModify2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Modify"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3960
            TabIndex        =   7
            Top             =   3120
            Width           =   870
         End
         Begin MSForms.ComboBox cboUserOld 
            Height          =   375
            Left            =   2190
            TabIndex        =   0
            Top             =   540
            Width           =   2565
            VariousPropertyBits=   1820346395
            BackColor       =   16777215
            DisplayStyle    =   3
            Size            =   "4524;661"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdChange 
            Height          =   570
            Left            =   2160
            TabIndex        =   8
            Top             =   3720
            Width           =   2535
            ForeColor       =   16711680
            VariousPropertyBits=   19
            Caption         =   "Change Profile"
            Size            =   "4471;1005"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.TextBox txtConfirmPass2 
            Height          =   375
            Left            =   2190
            TabIndex        =   4
            Top             =   2625
            Width           =   2565
            VariousPropertyBits=   1825589275
            BackColor       =   16777215
            Size            =   "4524;661"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtNewPass2 
            Height          =   375
            Left            =   2190
            TabIndex        =   3
            Top             =   2085
            Width           =   2565
            VariousPropertyBits=   1825589275
            BackColor       =   16777215
            Size            =   "4524;661"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtOldPass 
            Height          =   375
            Left            =   2190
            TabIndex        =   2
            Top             =   1575
            Width           =   2565
            VariousPropertyBits=   1825589275
            BackColor       =   16777215
            Size            =   "4524;661"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtNewUser2 
            Height          =   375
            Left            =   2190
            TabIndex        =   1
            Top             =   1080
            Width           =   2565
            VariousPropertyBits=   1825589275
            BackColor       =   16777215
            Size            =   "4524;661"
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New User Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   720
            TabIndex        =   36
            Top             =   1110
            Width           =   1380
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   825
            TabIndex        =   35
            Top             =   2085
            Width           =   1260
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   570
            TabIndex        =   34
            Top             =   2640
            Width           =   1545
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   900
            TabIndex        =   33
            Top             =   1575
            Width           =   1170
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old User Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   750
            TabIndex        =   32
            Top             =   540
            Width           =   1290
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rights"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1470
            TabIndex        =   31
            Top             =   3120
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   4920
         Left            =   -75105
         TabIndex        =   22
         Top             =   345
         Width           =   5505
         Begin MSForms.TextBox txtDelConfirm 
            Height          =   390
            Left            =   2085
            TabIndex        =   11
            Top             =   1755
            Width           =   2265
            VariousPropertyBits=   1825589275
            Size            =   "3995;688"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtDelPass 
            Height          =   390
            Left            =   2085
            TabIndex        =   10
            Top             =   1260
            Width           =   2265
            VariousPropertyBits=   1825589275
            Size            =   "3995;688"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdDelete 
            Height          =   585
            Left            =   2100
            TabIndex        =   12
            Top             =   2670
            Width           =   2205
            ForeColor       =   16711680
            VariousPropertyBits=   19
            Caption         =   "Delete User"
            Size            =   "3889;1032"
            FontName        =   "Arial"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.TextBox txtDelUser 
            Height          =   390
            Left            =   2100
            TabIndex        =   9
            Top             =   765
            Width           =   2265
            VariousPropertyBits=   1825589275
            Size            =   "3995;688"
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label lblDelUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1050
            TabIndex        =   30
            Top             =   750
            Width           =   960
         End
         Begin VB.Label lblDelPass 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1125
            TabIndex        =   29
            Top             =   1305
            Width           =   840
         End
         Begin VB.Label lblDelConfirm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   435
            TabIndex        =   28
            Top             =   1770
            Width           =   1545
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   4920
         Left            =   -150
         TabIndex        =   21
         Top             =   360
         Width           =   5490
         Begin VB.CheckBox chkRead 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1965
            TabIndex        =   16
            Top             =   2265
            Width           =   870
         End
         Begin VB.CheckBox chkWrite 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2880
            TabIndex        =   17
            Top             =   2265
            Width           =   915
         End
         Begin VB.CheckBox chkModify 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Modify"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3795
            TabIndex        =   18
            Top             =   2265
            Width           =   945
         End
         Begin MSForms.CommandButton cmdCreate 
            Height          =   585
            Left            =   2010
            TabIndex        =   19
            Top             =   3030
            Width           =   2355
            ForeColor       =   16711680
            VariousPropertyBits=   19
            Caption         =   "Create User"
            Size            =   "4154;1032"
            FontName        =   "Arial"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.TextBox txtNewConfirm1 
            Height          =   390
            Left            =   1965
            TabIndex        =   15
            Top             =   1725
            Width           =   2370
            VariousPropertyBits=   1825589275
            Size            =   "4180;688"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtNewPass1 
            Height          =   390
            Left            =   1965
            TabIndex        =   14
            Top             =   1200
            Width           =   2370
            VariousPropertyBits=   1825589275
            Size            =   "4180;688"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtNewUser1 
            Height          =   390
            Left            =   1965
            TabIndex        =   13
            Top             =   675
            Width           =   2370
            VariousPropertyBits=   1825589275
            Size            =   "4180;688"
            SpecialEffect   =   6
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label lblnewUsr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   885
            TabIndex        =   27
            Top             =   705
            Width           =   975
         End
         Begin VB.Label lblNewPass 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   975
            TabIndex        =   26
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label lblConfirmPass 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   270
            TabIndex        =   25
            Top             =   1755
            Width           =   1575
         End
         Begin VB.Label lblRights 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rights"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1245
            TabIndex        =   24
            Top             =   2220
            Width           =   555
         End
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Height          =   7605
      Left            =   1920
      Top             =   180
      Width           =   8265
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator Option"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3780
      TabIndex        =   38
      Top             =   510
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Esc To Exit"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10800
      TabIndex        =   37
      Top             =   7440
      Width           =   975
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoRS As New ADODB.Recordset
Dim adoCmd1 As New ADODB.Command
Dim strSql1 As String

Private Sub adminTab_Click(PreviousTab As Integer)
On Error Resume Next
query = "select * from loginpassword"
If adoRS.State = 1 Then adoRS.Close
adoRS.Open query, con, adOpenDynamic, adLockOptimistic, adCmdText
cboUserOld.Clear
If adoRS.EOF = True Then
MsgBox "No Record Found"
Exit Sub
End If
Do While Not adoRS.EOF = True
cboUserOld.AddItem adoRS!usrname
adoRS.MoveNext
Loop
adoRS.Close
End Sub
Private Sub cboUserOld_DropButtonClick()
'On Error Resume Next
'query = "select * from loginpassword"
'If adoRS.State = 1 Then adoRS.Close
'adoRS.Open query, con, adOpenDynamic, adLockOptimistic, adCmdText
'
'If adoRS.EOF = True Then
'MsgBox "No Record Found"
'Exit Sub
'End If
'Do While Not adoRS.EOF = True
'cboUserOld.AddItem adoRS!usrname
'adoRS.MoveNext
'Loop
'adoRS.Close
End Sub

Private Sub cboUserOld_LostFocus()
cboUserOld.Text = StrConv(cboUserOld.Text, vbProperCase)
End Sub

Private Sub chkModify_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdCreate.SetFocus
End If
End Sub

Private Sub chkModify2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdChange.SetFocus
End If
End Sub

Private Sub chkRead_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
chkWrite.SetFocus
End If
End Sub



Private Sub chkRead2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
chkWrite2.SetFocus
End If
End Sub

Private Sub chkWrite_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
chkModify.SetFocus
End If
End Sub


Private Sub chkWrite2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
chkModify2.SetFocus
End If
End Sub

Private Sub cmdChange_Click()
On Error Resume Next
If (Not cboUserOld.Text = "") And (Not txtNewUser2.Text = "") And (Not txtOldPass.Text = "") And (Not txtNewPass2.Text = "") And (Not txtConfirmPass2.Text = "") Then

    If (txtConfirmPass2.Text = txtNewPass2.Text) Then

            If adoRS.State = 1 Then adoRS.Close
            adoRS.Open "select * from loginpassword where usrname= '" & cboUserOld.Text & "' and usrpassword = '" & txtOldPass.Text & "'", con, adOpenStatic, adLockOptimistic, adCmdText

            If adoRS.BOF And adoRS.EOF Then
            MsgBox "No Matching Record Found"
            txtNewUser2.Text = ""
            txtOldPass.Text = ""
            txtNewPass2.Text = ""
            txtConfirmPass2.Text = ""
            chkRead2.Value = 0
            chkWrite2.Value = 0
            chkModify2.Value = 0
            cboUserOld.SetFocus
            Exit Sub
            End If

            If adoRS.EOF = False Then
                adoRS.Close
                Set adoRS = Nothing
                '''''''''''''''''''''''''''''''''''''''
                If chkRead2.Value = 1 Then
                Zread = "S"
                Else
                Zread = "N"
                End If
                If chkWrite2.Value = 1 Then
                Zwrite = "D"
                Else
                Zwrite = "N"
                End If
                If chkModify2.Value = 1 Then
                Zmodify = "M"
                Else
                Zmodify = "N"
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''
                strSql1 = "update loginpassword set  [usrname] = '" & txtNewUser2.Text & "',[usrpassword] = '" & txtNewPass2.Text & "', [read] = '" & Zread & "', [write] = '" & Zwrite & "', [modify] = '" & Zmodify & "'  where usrname = '" & cboUserOld.Text & "' and usrpassword='" & txtOldPass.Text & "'"


                adoCmd1.CommandText = strSql1
                adoCmd1.CommandType = adCmdText
                adoCmd1.ActiveConnection = con
                adoCmd1.Execute
                ThankMsg = "Updation"
                frmThanksDialog.Show vbModal
                cboUserOld.Text = ""
                txtNewUser2.Text = ""
                txtOldPass.Text = ""
                txtNewPass2.Text = ""
                txtConfirmPass2.Text = ""
                chkRead2.Value = 0
                chkWrite2.Value = 0
                chkModify2.Value = 0
                cboUserOld.SetFocus

        Else:
            adoRS.Close
            Set adoRS = Nothing
            MsgBox "Incorrect Username Or Password"
            cboUserOld.Text = ""
            txtNewUser2.Text = ""
            txtOldPass.Text = ""
            txtNewPass2.Text = ""
            txtConfirmPass2.Text = ""
            chkRead2.Value = 0
            chkWrite2.Value = 0
            chkModify2.Value = 0
            cboUserOld.SetFocus
        End If
    Else
        MsgBox "New Password And Confirm Password Are Not Match", vbOKOnly + vbInformation, "Error"
        cboUserOld.Text = ""
            txtNewUser2.Text = ""
            txtOldPass.Text = ""
            txtNewPass2.Text = ""
            txtConfirmPass2.Text = ""
            chkRead2.Value = 0
            chkWrite2.Value = 0
            chkModify2.Value = 0
            cboUserOld.SetFocus
    End If

Else
MsgBox "Please Fill Up All The Entries Properly", vbOKOnly + vbQuestion, "Be Alert"
cboUserOld.SetFocus
End If
Exit Sub

End Sub

Private Sub CmdCreate_Click()
On Error Resume Next
query = "select * from loginpassword"
If adoRS.State = 1 Then adoRS.Close
adoRS.Open query, con, adOpenDynamic, adLockOptimistic, adCmdText

If adoRS.EOF = True Then
MsgBox "No Record Found"
Exit Sub
End If

If (Not txtNewUser1.Text = "") And (Not txtNewPass1.Text = "") And (Not txtNewConfirm1.Text = "") Then
    If (txtNewPass1.Text = txtNewConfirm1.Text) Then
        adoRS.AddNew
        adoRS!usrname = txtNewUser1.Text
        adoRS!usrpassword = txtNewPass1.Text
         If chkRead.Value = 1 Then
            adoRS!Read = "S"
           Else
            adoRS!Read = "N"
         End If
        If chkWrite.Value = 1 Then
            adoRS!Write = "D"
           Else
            adoRS!Write = "N"
        End If

        If chkModify.Value = 1 Then
            adoRS!modify = "M"
           Else
            adoRS!modify = "N"
        End If

        adoRS.Update
        MsgBox "User Create Successfully !!", vbInformation
        txtNewUser1.Text = ""
        txtNewPass1.Text = ""
        txtNewConfirm1.Text = ""
        chkRead.Value = 0
        chkWrite.Value = 0
        chkModify.Value = 0
        txtNewUser1.SetFocus
    Else
        MsgBox "Password And Confirm Password Are Not Match", vbOKOnly + vbInformation, "Error"
        txtNewPass1.Text = ""
        txtNewConfirm1.Text = ""
        txtNewPass1.SetFocus
    End If
Else
MsgBox "Please Fill Up All The Entries Properly", vbOKOnly + vbQuestion, "Be Alert"
txtNewUser1.SetFocus
Exit Sub
End If
End Sub



Private Sub cmdDelete_Click()
If (Not txtDelUser.Text = "") And (Not txtDelPass.Text = "") And (Not txtDelConfirm.Text = "") Then
    If (txtDelPass.Text = txtDelConfirm.Text) Then
      query = "delete from loginpassword where usrname='" & txtDelUser.Text & "' and usrpassword='" & txtDelPass.Text & "'"
        If adoRS.State = 1 Then adoRS.Close
        adoRS.Open query, con, adOpenDynamic, adLockOptimistic, adCmdText
        ThankMsg = "Deletion"
        MsgBox "User Create Successfully !!", vbInformation
        txtDelUser.Text = ""
        txtDelPass.Text = ""
        txtDelConfirm.Text = ""
        txtDelUser.SetFocus
    Else
        MsgBox "Password And Confirm Password Are Not Match", vbOKOnly + vbInformation, "Error"
        txtDelPass.Text = ""
        txtDelConfirm.Text = ""
        txtDelPass.SetFocus
    End If
Else
MsgBox "Please Fill Up All The Entries Properly", vbOKOnly + vbQuestion, "Be Alert"
txtDelUser.SetFocus
Exit Sub
End If
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Unload Me
End If
End Sub


Private Sub txtDelConfirm_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = vbKeyEscape Then
txtDelConfirm.Text = ""
End If
End Sub

Private Sub txtDelPass_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = vbKeyEscape Then
txtDelPass.Text = ""
End If
End Sub

Private Sub txtDelUser_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = vbKeyEscape Then
txtDelUser.Text = ""
End If
End Sub

Private Sub txtDelUser_LostFocus()
txtDelUser.Text = StrConv(txtDelUser.Text, vbProperCase)

End Sub

Private Sub txtNewConfirm1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = vbKeyEscape Then
txtNewConfirm1.Text = ""
End If

If KeyAscii = 13 Then
chkRead.SetFocus
End If
End Sub
Private Sub txtNewPass1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = vbKeyEscape Then
txtNewPass1.Text = ""
End If
End Sub
Private Sub txtNewUser1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = vbKeyEscape Then
txtNewUser1.Text = ""
End If
End Sub

Private Sub txtNewUser1_LostFocus()
'''''''''''check user name in database'''''''
query = "select * from loginpassword where usrname='" & txtNewUser1.Text & "'"
Set adoRS = New ADODB.Recordset
If adoRS.State = 1 Then adoRS.Close
adoRS.Open query, con, adOpenDynamic, adLockOptimistic, adCmdText
If adoRS.EOF = False Then
MsgBox "User Already Exits"
txtNewUser1.Text = ""
txtNewUser1.SetFocus
Exit Sub
End If
'''''''''''''''''
txtNewUser1.Text = StrConv(txtNewUser1.Text, vbProperCase)
End Sub

Private Sub txtNewUser2_LostFocus()
txtNewUser2.Text = StrConv(txtNewUser2.Text, vbProperCase)
End Sub

