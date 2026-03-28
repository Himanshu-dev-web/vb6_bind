VERSION 5.00
Begin VB.Form frmStudentDetail 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCollagelID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2925
   End
   Begin VB.TextBox txtAnniversaryDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Top             =   4815
      Width           =   1695
   End
   Begin VB.TextBox txtDOB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Top             =   4530
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5700
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   495
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5700
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   495
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5700
      Width           =   1215
   End
   Begin VB.TextBox txtaddress3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   2820
      Width           =   4140
   End
   Begin VB.TextBox txtaddress2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2535
      Width           =   4140
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   4245
      Width           =   4155
   End
   Begin VB.TextBox txtfax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Top             =   3960
      Width           =   4155
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   3675
      Width           =   4155
   End
   Begin VB.TextBox txtstate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   3390
      Width           =   2895
   End
   Begin VB.TextBox txtaddress1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2250
      Width           =   4140
   End
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   3105
      Width           =   2895
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   1965
      Width           =   4125
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "PRESS F2 TO SEARCH RECORD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   29
      Top             =   1680
      Width           =   3090
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "PRESS F2 TO ADD COLLEGE ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   28
      Top             =   360
      Width           =   3105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "School ID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   27
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Anniversary Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   26
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1470
      TabIndex        =   25
      Top             =   4575
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address3:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1770
      TabIndex        =   24
      Top             =   2820
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address2:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1770
      TabIndex        =   23
      Top             =   2580
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   22
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2205
      TabIndex        =   21
      Top             =   4020
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1695
      TabIndex        =   20
      Top             =   3720
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2115
      TabIndex        =   19
      Top             =   3420
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2220
      TabIndex        =   18
      Top             =   3120
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2070
      TabIndex        =   17
      Top             =   1980
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address1:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1770
      TabIndex        =   9
      Top             =   2280
      Width           =   870
   End
End
Attribute VB_Name = "frmStudentDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDelete_Click()
con.Execute "delete from studentdetail where name='" & txtname.Text & "'"
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
Dim O As Object
    For Each O In Me
          If TypeOf O Is TextBox Then
          O.Text = ""
          End If
      Next
      txtCollagelID.SetFocus
End Sub

Private Sub cmdSave_Click()
SaveData
Call cmd
End Sub

Sub SaveData()
If txtname.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If
        
         
         If rs.State = 1 Then rs.Close
         
         rs.Open "select * from studentDetail where collagelId='" & txtCollagelID & "'", con, adOpenDynamic, adLockOptimistic
                 
         If rs.EOF = True Then
            rs.AddNew
                        rs!collagelId = txtCollagelID.Text
                        rs!Name = txtname.Text
                        rs!Address1 = txtaddress1.Text
                        rs!Address2 = txtaddress2.Text
                        rs!Address3 = txtaddress3.Text
                        rs!CITY = txtcity.Text
                        rs!State = txtstate.Text
                        rs!phone = txtphone.Text
                        rs!FAX = txtfax.Text
                        rs!email = txtemail.Text
                        rs!DOB = txtDOB.Text
                        rs!anniversaryDate = txtAnniversaryDate.Text
            rs.Update
          
          Else
                        rs!collagelId = txtCollagelID.Text
                        rs!Name = txtname.Text
                        rs!Address1 = txtaddress1.Text
                        rs!Address2 = txtaddress2.Text
                        rs!Address3 = txtaddress3.Text
                        rs!CITY = txtcity.Text
                        rs!State = txtstate.Text
                        rs!phone = txtphone.Text
                        rs!FAX = txtfax.Text
                        rs!email = txtemail.Text
                        rs!DOB = txtDOB.Text
                        rs!anniversaryDate = txtAnniversaryDate.Text
            rs.Update
         
         End If
     
End Sub
Sub Search()
            
           ' If rs.State = 1 Then rs.Close
           Set rs1 = New Recordset
            rs1.Open "select * from studentDetail where name='" & txtname.Text & "'", con
            If rs1.EOF = False Then
            txtCollagelID.Text = rs1!collagelId
            txtname.Text = rs1!Name
            txtaddress1.Text = rs1!Address1
            txtaddress2.Text = rs1!Address2
            txtaddress3.Text = rs1!Address3
            txtcity.Text = rs1!CITY
            txtstate.Text = rs1!State
            txtphone.Text = rs1!phone
            txtfax.Text = rs1!FAX
            txtemail.Text = rs1!email
            txtDOB.Text = rs1!DOB
            txtAnniversaryDate.Text = rs1!anniversaryDate
            
            End If
End Sub

Private Sub txtSchoolID_GotFocus()

End Sub

Private Sub txtSchoolID_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub


Private Sub txtaddress1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtaddress2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtaddress3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtAnniversaryDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtcity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtCollagelID_GotFocus()
If PopUpValue1 <> "" Then
txtCollagelID.Text = PopUpValue1
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
PopUpValue4 = ""
End If
End Sub

Private Sub txtCollagelID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If

If KeyCode = 113 Then
popuplist2 "select CollegeID from college", con
End If
End Sub

Private Sub txtDOB_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtfax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtname_GotFocus()
If PopUpValue1 > "" Then
txtname.Text = PopUpValue2
Search
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
PopUpValue4 = ""
End If
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
If KeyCode = 113 Then
popuplist2 "select * from studentdetail", con
End If
End Sub

Private Sub txtphone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub
