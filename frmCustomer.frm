VERSION 5.00
Begin VB.Form frmCustomer 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Customer Details"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   2857
      Width           =   6255
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2411
      Width           =   6240
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   0
      Top             =   2040
      Width           =   6225
   End
   Begin VB.TextBox txtTinNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      MaxLength       =   25
      TabIndex        =   3
      Top             =   3303
      Width           =   2655
   End
   Begin VB.TextBox txtCST 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   4
      Top             =   3660
      Width           =   2655
   End
   Begin VB.TextBox txtName1 
      Height          =   285
      Left            =   11700
      TabIndex        =   33
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtcharges 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10140
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12180
      TabIndex        =   15
      Top             =   7620
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtHead 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6420
      MaxLength       =   5
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CheckBox cartage 
      Caption         =   "Cartage"
      Height          =   315
      Left            =   13380
      TabIndex        =   14
      Top             =   7200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox unloading 
      Caption         =   "Unloading"
      Height          =   315
      Left            =   11700
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CheckBox lOADING 
      Caption         =   "Loading "
      Height          =   315
      Left            =   10140
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10140
      TabIndex        =   9
      Top             =   6300
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtstate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      MaxLength       =   35
      TabIndex        =   5
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox txtfax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10095
      TabIndex        =   10
      Top             =   5550
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10095
      TabIndex        =   11
      Top             =   5910
      Visible         =   0   'False
      Width           =   6255
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
      Height          =   540
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4905
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
      Height          =   540
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4905
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
      Height          =   540
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4905
      Width           =   1095
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
      Height          =   540
      Left            =   5955
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4905
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
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
      Left            =   2520
      TabIndex        =   34
      Top             =   4080
      Width           =   1035
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   315
      Left            =   9240
      TabIndex        =   32
      Top             =   8100
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   255
      Left            =   11700
      TabIndex        =   31
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Head"
      Height          =   315
      Left            =   9240
      TabIndex        =   30
      Top             =   7680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "PAN NO.:"
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
      Left            =   2520
      TabIndex        =   29
      Top             =   3660
      Width           =   1035
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN"
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
      Left            =   2520
      TabIndex        =   28
      Top             =   3300
      Width           =   735
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
      Left            =   2520
      TabIndex        =   27
      Top             =   2355
      Width           =   1275
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
      Left            =   2520
      TabIndex        =   26
      Top             =   1995
      Width           =   1335
   End
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   14640
      TabIndex        =   25
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   15240
      TabIndex        =   24
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
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
      Left            =   2520
      TabIndex        =   23
      Top             =   2820
      Width           =   1155
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
      Left            =   9165
      TabIndex        =   22
      Top             =   5580
      Visible         =   0   'False
      Width           =   915
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
      Left            =   9165
      TabIndex        =   21
      Top             =   5940
      Visible         =   0   'False
      Width           =   795
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
      Left            =   3780
      TabIndex        =   20
      Top             =   1485
      Width           =   3195
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2520
      TabIndex        =   19
      Top             =   495
      Width           =   7470
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim Str As String
Dim yrs As String
Dim FSO As New FileSystemObject
Private Sub cartage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub cmdExit_Click()
 Unload frmCustomer
End Sub

Private Sub Command1_Click()
    Dim O As Object
    For Each O In Me
          If TypeOf O Is TextBox Then
          O.Text = ""
          End If
      Next
      Me.lOADING.Value = 0
      Me.unloading.Value = 0
      Me.cartage.Value = 0
      
      Record = ""
      txtname.SetFocus
      popuplist11 "Select name from Customer order by Name", con, , True, colmn
End Sub

Private Sub Command2_Click()
      
      If txtname.Text = "" Then
         MsgBox "Please Search Records !!", vbInformation
         Exit Sub
      End If
      
      
      
      If rs.State = 1 Then rs.Close
      rs.Open "Select * from Customer where Name='" & txtname.Text & "'", con, adOpenDynamic, adLockOptimistic
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
              
              If txtName1.Text <> "" Then
              rs.Open "Select * from Customer where Name='" & txtName1.Text & "'", con, adOpenDynamic, adLockOptimistic
              Else
              rs.Open "Select * from Customer where Name='" & txtname.Text & "'", con, adOpenDynamic, adLockOptimistic
              End If
              
              
                  If rs.RecordCount >= 1 Then
                     If MsgBox("Do U Want To Update !!!", vbInformation + vbYesNo) = vbYes Then
                        
                        con.Execute "update FoldingDetails set PartyName='" & txtname.Text & "' where PartyName='" & txtName1.Text & "'"
                        
                        rs!Name = txtname.Text
                        rs!Address = txtaddress.Text
                        rs!CITY = txtcity.Text
                        rs!State = txtstate.Text
                        rs!phone = txtphone.Text
                        rs!FAX = txtfax.Text
                        rs!email = txtemail.Text
                        rs!cst = txtCST.Text
                        rs!tin = txtTinNo.Text
                        rs!lOADING = lOADING.Value
                        rs!unloading = unloading.Value
                        rs!cartage = cartage.Value
                        rs!TaxHead = txtHead.Text
                        rs!TaxRate = txtRate.Text
                        rs!OtherCharges = txtcharges.Text
                        
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
                        rs!lOADING = lOADING.Value
                        rs!unloading = unloading.Value
                        rs!cartage = cartage.Value
                        rs!TaxHead = Val(txtHead.Text)
                        rs!TaxRate = Val(txtRate.Text)
                        rs!OtherCharges = Val(txtcharges.Text)
                        
                        
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
       popuplist11 "Select name from Customer order by Name", con, , True, colmn
       popuplist1.Show
    End If
End Sub

Private Sub Form_Load()
'''---------------------------------------------------
Me.Top = 1500
Me.Left = 2000

Record = ""
colmn = 1
popuplist11 "Select name from Customer order by Name", con, , True, colmn

LblName.Caption = "Customer Details"
tbl = "Customer"
pos Me

End Sub
Private Sub lOADING_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtaddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub
Private Sub txtcharges_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
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
Private Sub txtHead_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
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
If rs.State = 1 Then rs.Close
rs.Open "Select * from Customer where Name='" & Record & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
   txtname.Text = rs!Name
   txtaddress.Text = rs!Address
   txtcity.Text = rs!CITY
   txtstate.Text = rs!State
   txtphone.Text = rs!phone
   txtfax.Text = rs!FAX
   txtemail.Text = rs!email
   txtCST.Text = rs!cst & ""
   txtTinNo.Text = rs!tin & ""
   
   txtName1 = rs!Name
   
   lOADING.Value = IIf(rs!lOADING = True, 1, 0)
   unloading.Value = IIf(rs!unloading = True, 1, 0)
   cartage.Value = IIf(rs!cartage = True, 1, 0)
   txtHead.Text = rs!TaxHead & ""
   txtRate.Text = rs!TaxRate & ""
   txtcharges.Text = rs!OtherCharges & ""

   
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
Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
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
              rs.Open "Select * from Customer where SupplierName='" & vs.TextMatrix(vs.RowSel, 0) & "'and finanyear='" & vs.TextMatrix(vs.RowSel, 1) & "'", con, adOpenDynamic, adLockOptimistic
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
              rs.Open "Select * from Customer where SupplierName='" & vs.TextMatrix(vs.RowSel, 0) & "'and finanyear='" & vs.TextMatrix(vs.RowSel, 1) & "'", con, adOpenDynamic, adLockOptimistic
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
Private Sub unloading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub
