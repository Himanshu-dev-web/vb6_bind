VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompany 
   BackColor       =   &H005194A9&
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtfinanYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   3975
   End
   Begin VB.TextBox txtPen 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7725
      TabIndex        =   7
      Top             =   1275
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10635
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8445
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtservice 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7725
      TabIndex        =   12
      Top             =   2715
      Width           =   2415
   End
   Begin VB.TextBox txtcstno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7725
      TabIndex        =   10
      Top             =   2355
      Width           =   2415
   End
   Begin VB.TextBox txtup 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7725
      TabIndex        =   9
      Top             =   1995
      Width           =   3975
   End
   Begin VB.TextBox txtreg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7725
      TabIndex        =   8
      Top             =   1635
      Width           =   3975
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7725
      TabIndex        =   15
      Top             =   3435
      Width           =   3975
   End
   Begin VB.TextBox txtfax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7725
      TabIndex        =   14
      Top             =   3075
      Width           =   3975
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1695
      TabIndex        =   5
      Top             =   3075
      Width           =   3975
   End
   Begin VB.TextBox txtstate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4335
      TabIndex        =   4
      Top             =   2715
      Width           =   1335
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   1695
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1635
      Width           =   3975
   End
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1695
      TabIndex        =   3
      Top             =   2715
      Width           =   1335
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1695
      TabIndex        =   1
      Top             =   1275
      Width           =   3975
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3255
      Left            =   75
      TabIndex        =   0
      Top             =   4335
      Width           =   11700
      _cx             =   20637
      _cy             =   5741
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   10667465
      ForeColor       =   -2147483640
      BackColorFixed  =   12640511
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   16777215
      BackColorBkg    =   5346473
      BackColorAlternate=   10667465
      GridColor       =   16744448
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   2000
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComCtl2.DTPicker dp1 
      Height          =   255
      Left            =   10245
      TabIndex        =   11
      Top             =   2370
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   22740995
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin MSComCtl2.DTPicker dp2 
      Height          =   255
      Left            =   10245
      TabIndex        =   13
      Top             =   2730
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   22740995
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   3465
      Width           =   1245
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "PAN No "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6405
      TabIndex        =   32
      Top             =   1275
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Details"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   195
      TabIndex        =   31
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax No :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6405
      TabIndex        =   30
      Top             =   2715
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "C.S.T. && ST No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6405
      TabIndex        =   29
      Top             =   2355
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "U.P.T.T. No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6405
      TabIndex        =   28
      Top             =   1995
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pest. Licence No.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6405
      TabIndex        =   27
      Top             =   1635
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Email "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6405
      TabIndex        =   26
      Top             =   3435
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6405
      TabIndex        =   25
      Top             =   3075
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   24
      Top             =   3075
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3375
      TabIndex        =   23
      Top             =   2715
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   22
      Top             =   2715
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   21
      Top             =   1275
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   20
      Top             =   1635
      Width           =   1335
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim Str As String
Dim yrs As String
Dim FSO As New FileSystemObject


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    Dim O As Object
    For Each O In Me
          If TypeOf O Is TextBox Then
          O.Text = ""
          End If
      Next
      txtname.SetFocus
End Sub

Private Sub Command2_Click()
      If txtname.Text = "" Then
         MsgBox "Please Search Records !!", vbInformation
         Exit Sub
      End If
      
      If Rs.State = 1 Then Rs.Close
      Rs.Open "Select * from Company where CompanyName='" & txtname.Text & "' and finanyear= '" & txtfinanYear.Text & "'", adoConn1, adOpenDynamic, adLockOptimistic
      If Rs.RecordCount >= 1 Then
        If MsgBox("Do U Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
             Rs.delete adAffectCurrent
        End If
      End If
      Call Command1_Click
      fillgrid
End Sub

Private Sub Command6_Click()
              
              Set Rs = New ADODB.Recordset
              If txtname.Text = "" Or txtaddress.Text = "" Or txtcity.Text = "" Or txtphone.Text = "" Then
                      MsgBox "Please Verify Entry !!!", vbInformation
                 Exit Sub
              End If
              If Rs.State = 1 Then Rs.Close
              Rs.Open "Select * from Company where CompanyName='" & txtname.Text & "'and finanyear='" & txtfinanYear.Text & "'", adoConn1, adOpenDynamic, adLockOptimistic
                  If Rs.RecordCount >= 1 Then
                     If MsgBox("Do U Want To Update !!!", vbInformation + vbYesNo) = vbYes Then
                        '''rs!CompanyName = txtName.Text
                        Rs!Address = txtaddress.Text
                        Rs!City = txtcity.Text
                        Rs!State = txtaddress.Text
                        Rs!phone = txtphone.Text
                        ''''rs!finanYear = Trim(txtfinanYear.Text)
                        Rs!RegistNo = txtreg.Text
                        Rs!UpttNo = txtup.Text
                        Rs!CstNo = txtcstno.Text
                        Rs!CSTdates = dp1.Value
                        Rs!ServiceNo = txtservice.Text
                        Rs!ServiceDates = dp2.Value
                        Rs!Fax = txtfax.Text
                        Rs!email = txtemail.Text
                        Rs!penno = txtPen.Text
                        Rs.Update
                     End If
                   Else
                     If MsgBox("Do U Want To Save !!!", vbInformation + vbYesNo) = vbYes Then
                        Rs.AddNew
                        ''''''''''''
                        Rs!CompanyName = UCase(txtname.Text)
                        Str = txtname.Text
                        '''''''''''
                        Rs!Address = txtaddress.Text
                        Rs!City = txtcity.Text
                        Rs!State = txtstate & ""
                        Rs!phone = txtphone.Text
                        '''''''''''''
                        Rs!finanyear = Trim(txtfinanYear.Text)
                        yrs = Trim(txtfinanYear.Text)
                        ''''''''''''
                        Rs!RegistNo = txtreg.Text
                        Rs!UpttNo = txtup.Text
                        Rs!CstNo = txtcstno.Text
                        Rs!CSTdates = dp1.Value
                        Rs!ServiceNo = txtservice.Text
                        Rs!ServiceDates = dp2.Value
                        Rs!Fax = txtfax.Text
                        Rs!email = txtemail.Text
                        Rs!penno = txtPen.Text
                        Rs.Update
                        ''''''''''''''save data in company record also''''''''''
                        SaveInCompanyRecord
                        '''''''''''''''''''''''''''''''''''''''''''''''
                        '''''''''''''''Folder Creation of the Company'''''''''''''''''''''''''''''''''''
                                NewSession
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     End If
                  End If
                
               Call Command1_Click
               fillgrid

End Sub
Sub fillgrid()
   Dim fill As New ADODB.Recordset
   If fill.State = 1 Then fill.Close
   fill.Open "select CompanyName as [Company Name],FinanYear,Address,Phone,City,State,Fax from company order by CompanyName", adoConn1, adOpenDynamic, adLockOptimistic
   If fill.EOF = False Then
      Set vs.DataSource = fill
   End If
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

Private Sub Form_Load()
'''---------------------------------------------------
 If adoConn1.State = 1 Then adoConn1.Close
    adoConn1.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Companyrecord.mdb;"
    adoConn1.CursorLocation = adUseClient
''''-----------------------------------------------------
frmCompany.Top = 10
frmCompany.Left = 10
dp1.Value = Format(Date, "dd/MM/yy")
dp2.Value = Format(Date, "dd/MM/yy")
fillgrid
End Sub

Private Sub txtaddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtcstno_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtPen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
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
            If Rs.State = 1 Then Rs.Close
              Rs.Open "Select * from Company where CompanyName='" & vs.TextMatrix(vs.RowSel, 0) & "'and finanyear='" & vs.TextMatrix(vs.RowSel, 1) & "'", adoConn1, adOpenDynamic, adLockOptimistic
              If Rs.RecordCount >= 1 Then
                 txtname.Text = Rs!CompanyName
                 txtaddress.Text = Rs!Address
                 txtcity.Text = Rs!City
                 txtstate.Text = Rs!State
                 txtphone.Text = Rs!phone
                 txtfinanYear.Text = Rs!finanyear
                 txtreg.Text = Rs!RegistNo
                 txtup.Text = Rs!UpttNo
                 txtcstno.Text = Rs!CstNo
                 dp1.Value = Rs!CSTdates
                 txtservice.Text = Rs!ServiceNo
                 dp2.Value = Rs!ServiceDates
                 txtfax.Text = Rs!Fax
                 txtemail.Text = Rs!email
                 txtPen.Text = Rs!penno
              End If

End Sub

Private Sub VS_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
              Call Command1_Click
              If Rs.State = 1 Then Rs.Close
              Rs.Open "Select * from Company where CompanyName='" & vs.TextMatrix(vs.RowSel, 0) & "'and finanyear='" & vs.TextMatrix(vs.RowSel, 1) & "'", adoConn1, adOpenDynamic, adLockOptimistic
              If Rs.RecordCount >= 1 Then
                 txtname.Text = Rs!CompanyName
                 txtaddress.Text = Rs!Address
                 txtcity.Text = Rs!City
                 txtstate.Text = Rs!State
                 txtphone.Text = Rs!phone
                 txtfinanYear.Text = Rs!finanyear
                 txtreg.Text = Rs!RegistNo
                 txtup.Text = Rs!UpttNo
                 txtcstno.Text = Rs!CstNo
                 dp1.Value = Rs!CSTdates
                 txtservice.Text = Rs!ServiceNo
                 dp2.Value = Rs!ServiceDates
                 txtfax.Text = Rs!Fax
                 txtemail.Text = Rs!email
                 txtPen.Text = Rs!penno
              End If
 End If

End Sub
Sub SaveInCompanyRecord()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If con.State = 1 Then con.Close
    con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CompanyRecord.mdb;"
    con.CursorLocation = adUseClient
        
    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from company", con, adOpenDynamic, adLockOptimistic, adCmdText
''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Rs.AddNew
                        Rs!CompanyName = txtname.Text
                        Rs!Address = txtaddress.Text
                        Rs!City = txtcity.Text
                        Rs!State = txtstate & ""
                        Rs!phone = txtphone.Text
                        Rs!finanyear = Trim(txtfinanYear.Text)
                        Rs!RegistNo = txtreg.Text
                        Rs!UpttNo = txtup.Text
                        Rs!CstNo = txtcstno.Text
                        Rs!CSTdates = dp1.Value
                        Rs!ServiceNo = txtservice.Text
                        Rs!ServiceDates = dp2.Value
                        Rs!Fax = txtfax.Text
                        Rs!email = txtemail.Text
                        Rs!penno = txtPen.Text
                        Rs.Update

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub NewSession()
Dim Co As Boolean

    '''''''''''''''''''''''''''check the Co. folder already exists''''''
     If FSO.FolderExists(App.Path & "\" & Str) Then
        Screen.MousePointer = vbDefault
        Co = True
    Else
        FSO.CreateFolder App.Path & "\" & Str
        FSO.CreateFolder App.Path & "\" & Str & "\" & yrs
        FSO.CopyFile App.Path & "\" & "NehaEnterprises.mdb", App.Path & "\" & Str & "\" & yrs & "\" & Str & ".mdb", False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ''''''''''''''''''''''''if Co Already Exists''''''''
    If Co = True Then
        '''''''''''''check the Year sub folder exists'''''''''
        If FSO.FolderExists(App.Path & "\" & Str & "\" & yrs) Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            FSO.CreateFolder App.Path & "\" & Str & "\" & yrs
            FSO.CopyFile App.Path & "\" & "NehaEnterprises.mdb", App.Path & "\" & Str & "\" & yrs & "\" & Str & ".mdb", False
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''

