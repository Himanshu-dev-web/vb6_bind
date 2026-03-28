VERSION 5.00
Begin VB.Form frmSupplier 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Binder Details"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11610
   Icon            =   "frmSupplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTinNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11295
      TabIndex        =   7
      Top             =   5940
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.TextBox txtCST 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11265
      TabIndex        =   8
      Top             =   6300
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6795
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11280
      TabIndex        =   6
      Top             =   5595
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.TextBox txtfax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11280
      TabIndex        =   5
      Top             =   5235
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11295
      TabIndex        =   4
      Top             =   4905
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.TextBox txtstate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   14415
      TabIndex        =   3
      Top             =   4545
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3555
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2340
      Width           =   6240
   End
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11295
      TabIndex        =   2
      Top             =   4545
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3555
      TabIndex        =   0
      Top             =   1980
      Width           =   6225
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tin No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9855
      TabIndex        =   22
      Top             =   5970
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "CST No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9855
      TabIndex        =   21
      Top             =   6330
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search Record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3585
      TabIndex        =   20
      Top             =   1410
      Width           =   3195
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Email "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9870
      TabIndex        =   19
      Top             =   5625
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9870
      TabIndex        =   18
      Top             =   5265
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   9855
      TabIndex        =   17
      Top             =   4935
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "CST:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   16
      Top             =   4530
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "UPTT:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9855
      TabIndex        =   15
      Top             =   4575
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   2115
      TabIndex        =   14
      Top             =   1980
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2115
      TabIndex        =   13
      Top             =   2340
      Width           =   1335
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim Str As String
Dim yrs As String
Dim FSO As New FileSystemObject


Private Sub cmdExit_Click()
Unload frmSupplier
End Sub

Private Sub Command1_Click()
    Dim O As Object
    For Each O In Me
          If TypeOf O Is TextBox Then
          O.Text = ""
          End If
      Next
      Record = ""
      txtname.SetFocus
      popuplist11 "Select name from Supplier order by Name", con, , True, colmn
End Sub

Private Sub Command2_Click()
      If txtname.Text = "" Then
         MsgBox "Please Search Records !!", vbInformation
         Exit Sub
      End If
      
      
      
      If rs.State = 1 Then rs.Close
      rs.Open "Select * from Supplier where Name='" & txtname.Text & "'", con, adOpenDynamic, adLockOptimistic
      If rs.RecordCount >= 1 Then
        If MsgBox("Do U Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
             rs.Delete adAffectCurrent
        End If
      End If
      Call Command1_Click
      
End Sub

Private Sub Command6_Click()
              
              Set rs = New ADODB.Recordset
              If txtname.Text = "" Then
                      MsgBox "Please Verify Entry !!!", vbInformation
                 Exit Sub
              End If
              
              
              
         
              
              If rs.State = 1 Then rs.Close
              rs.Open "Select * from Supplier where Name='" & txtname.Text & "'", con, adOpenDynamic, adLockOptimistic
                  If rs.RecordCount >= 1 Then
                     If MsgBox("Do U Want To Update !!!", vbInformation + vbYesNo) = vbYes Then
                        rs!Name = txtname.Text
                        rs!Address = txtaddress.Text
                        rs!CITY = txtcity.Text
                        rs!State = txtstate.Text
                        rs!phone = txtphone.Text
                        rs!FAX = txtfax.Text
                        rs!email = txtemail.Text
                        rs!cst = txtCST.Text
                        rs!tin = txtTinNo.Text

                        rs.Update
                     End If
                   Else
                     If MsgBox("Do U Want To Save !!!", vbInformation + vbYesNo) = vbYes Then
                        rs.AddNew
                        rs!Name = txtname.Text
                        rs!Address = txtaddress.Text
                        rs!CITY = txtcity.Text
                        rs!State = txtstate.Text
                        rs!phone = txtphone.Text
                        rs!FAX = txtfax.Text
                        rs!email = txtemail.Text
                        rs!cst = txtCST.Text
                        rs!tin = txtTinNo.Text

                        rs.Update
                     End If
                  End If
                  Call Command1_Click
                  
               
                  
                  

End Sub
Private Sub Command7_Click()
Unload Me
End Sub

Private Sub dp1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub dp2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub Form_Activate()
txtname.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
       Unload Me
    ElseIf KeyCode = 113 Then
       
       popuplist11 "Select name from Supplier order by Name", con, , True, colmn
       popuplist1.Show
    End If
End Sub

Private Sub Form_Load()
'''---------------------------------------------------
Me.Top = 1500
Me.Left = 2000
colmn = 1

Record = ""
popuplist11 "Select name from Supplier order by Name", con, , True, colmn

'LblName.Caption = "Supplier Details"
tbl = "Supplier"
pos Me

End Sub

Private Sub txtaddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtcity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtcstno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub
Private Sub txtCST_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtfax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtfinanYear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtname_GotFocus()
 SearchData
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{Tab}"
   SearchData
End If
End Sub
Sub SearchData()
On Error Resume Next

If rs.State = 1 Then rs.Close
rs.Open "Select * from Supplier where Name='" & Record & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
   txtname.Text = rs!Name
   txtaddress.Text = rs!Address
   txtcity.Text = rs!CITY
   txtstate.Text = rs!State
   txtphone.Text = rs!phone
   txtfax.Text = rs!FAX
   txtemail.Text = rs!email
   txtCST.Text = rs!cst
   txtTinNo.Text = rs!tin

End If

End Sub
Private Sub txtPen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtphone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtreg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtservice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub vs_Click()
            If rs.State = 1 Then rs.Close
              rs.Open "Select * from Supplier where SupplierName='" & vs.TextMatrix(vs.RowSel, 0) & "'and finanyear='" & vs.TextMatrix(vs.RowSel, 1) & "'", con, adOpenDynamic, adLockOptimistic
              If rs.RecordCount >= 1 Then
                 txtname.Text = rs!Name
                 txtaddress.Text = rs!Address
                 txtcity.Text = rs!CITY
                 txtstate.Text = rs!State
                 txtphone.Text = rs!phone
                 txtfinanYear.Text = rs!finanyear
                 txtreg.Text = rs!RegistNo
                 txtup.Text = rs!UpttNo
                 txtcstno.Text = rs!CstNo
                 dp1.Value = rs!CSTdates
                 txtservice.Text = rs!ServiceNo
                 dp2.Value = rs!ServiceDates
                 txtfax.Text = rs!FAX
                 txtemail.Text = rs!email
                 txtPen.Text = rs!penno
              End If

End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
              Call Command1_Click
              If rs.State = 1 Then rs.Close
              rs.Open "Select * from Supplier where SupplierName='" & vs.TextMatrix(vs.RowSel, 0) & "'and finanyear='" & vs.TextMatrix(vs.RowSel, 1) & "'", con, adOpenDynamic, adLockOptimistic
              If rs.RecordCount >= 1 Then
                 txtname.Text = rs!Name
                 txtaddress.Text = rs!Address
                 txtcity.Text = rs!CITY
                 txtstate.Text = rs!State
                 txtphone.Text = rs!phone
                 txtfinanYear.Text = rs!finanyear
                 txtreg.Text = rs!RegistNo
                 txtup.Text = rs!UpttNo
                 txtcstno.Text = rs!CstNo
                 dp1.Value = rs!CSTdates
                 txtservice.Text = rs!ServiceNo
                 dp2.Value = rs!ServiceDates
                 txtfax.Text = rs!FAX
                 txtemail.Text = rs!email
                 txtPen.Text = rs!penno
              End If
 End If

End Sub
Private Sub txtTinNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

