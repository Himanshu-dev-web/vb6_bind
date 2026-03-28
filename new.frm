VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSchoolID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2640
      TabIndex        =   26
      Top             =   480
      Width           =   2925
   End
   Begin VB.TextBox txtAnniversaryDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   24
      Top             =   3975
      Width           =   1695
   End
   Begin VB.TextBox txtDOB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   22
      Top             =   3690
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4860
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   495
      Left            =   2655
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4860
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   495
      Left            =   5325
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4860
      Width           =   1215
   End
   Begin VB.TextBox txtaddress3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      ScrollBars      =   3  'Both
      TabIndex        =   16
      Top             =   1980
      Width           =   4140
   End
   Begin VB.TextBox txtaddress2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   1695
      Width           =   4140
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   3405
      Width           =   4155
   End
   Begin VB.TextBox txtfax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   3120
      Width           =   4155
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   2835
      Width           =   4155
   End
   Begin VB.TextBox txtstate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   2550
      Width           =   2895
   End
   Begin VB.TextBox txtaddress1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1410
      Width           =   4140
   End
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   2265
      Width           =   2895
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   1125
      Width           =   4125
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
      Left            =   1560
      TabIndex        =   27
      Top             =   480
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
      Left            =   960
      TabIndex        =   25
      Top             =   4080
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
      Left            =   1350
      TabIndex        =   23
      Top             =   3735
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
      Left            =   1650
      TabIndex        =   17
      Top             =   1980
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
      Left            =   1650
      TabIndex        =   15
      Top             =   1740
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
      Left            =   1920
      TabIndex        =   13
      Top             =   3480
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
      Left            =   2085
      TabIndex        =   12
      Top             =   3180
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
      Left            =   1575
      TabIndex        =   11
      Top             =   2880
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
      Left            =   1995
      TabIndex        =   10
      Top             =   2580
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
      Left            =   2100
      TabIndex        =   9
      Top             =   2280
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
      Left            =   1950
      TabIndex        =   8
      Top             =   1140
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
      Left            =   1650
      TabIndex        =   7
      Top             =   1440
      Width           =   870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
savedata
End Sub

Sub savedata()
If txtname.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If
        
         
         If rs.State = 1 Then rs.Close
         
         rs.Open "select * from newDetails where schoolId='" & txtSchoolID.Text & "'", con, adOpenDynamic, adLockOptimistic
                 
         If rs.EOF = True Then
            rs.AddNew
                        rs!schoolId = txtSchoolID.Text
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
                        rs!schoolId = txtSchoolID.Text
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

Private Sub Label11_Click()

End Sub
