VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmpayrec 
   Caption         =   "Receipt/Payment Amount"
   ClientHeight    =   9024
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9024
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2_dr 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2196
      TabIndex        =   27
      Top             =   1836
      Width           =   1116
   End
   Begin VB.OptionButton Option1_cr 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3384
      TabIndex        =   26
      Top             =   1836
      Value           =   -1  'True
      Width           =   1152
   End
   Begin VB.ComboBox cboFirm 
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtCustomer 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   2256
      Width           =   4035
   End
   Begin VB.TextBox txtBank 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6240
      TabIndex        =   6
      Top             =   2916
      Width           =   3096
   End
   Begin VB.TextBox txtCheckno 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Top             =   3336
      Width           =   3096
   End
   Begin VB.OptionButton bank 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Top             =   2856
      Width           =   975
   End
   Begin VB.OptionButton cash 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   2856
      Width           =   975
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   3276
      Width           =   1515
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4836
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4836
      Width           =   1275
   End
   Begin VB.CommandButton cdmDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4836
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4836
      Width           =   1215
   End
   Begin VB.TextBox txtrecNo 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   660
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4836
      Width           =   1215
   End
   Begin VB.ComboBox cboCustomer 
      Height          =   315
      Left            =   3780
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   660
      Visible         =   0   'False
      Width           =   375
   End
   Begin Crystal.CrystalReport cr 
      Left            =   1140
      Top             =   4896
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker recdate 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1020
      Width           =   1395
      _ExtentX        =   2455
      _ExtentY        =   550
      _Version        =   393216
      Format          =   139460609
      CurrentDate     =   39548
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Firm Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   25
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Receipt Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   24
      Top             =   1020
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Customer Name "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   360
      TabIndex        =   23
      Top             =   2256
      Width           =   1752
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Bank Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4440
      TabIndex        =   22
      Top             =   2916
      Width           =   1752
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cheque No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4440
      TabIndex        =   21
      Top             =   3336
      Width           =   1752
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Amount "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   360
      TabIndex        =   20
      Top             =   3276
      Width           =   1752
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Balance Amount :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   6300
      TabIndex        =   19
      Top             =   2256
      Visible         =   0   'False
      Width           =   2112
   End
   Begin VB.Label headBal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   8100
      TabIndex        =   18
      Top             =   2256
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Receipt No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   17
      Top             =   660
      Width           =   1755
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Customer Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   15
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmpayrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bank_Click()
If bank.Value = True Then
Me.txtBank.Enabled = True
Me.txtCheckno.Enabled = True
Me.txtCheckno.text = "Cheque"
End If
End Sub
Private Sub cash_Click()
If cash.Value = True Then
Me.txtBank.Enabled = False
Me.txtCheckno.Enabled = False
Me.txtCheckno.text = ""
End If
End Sub
Private Sub cdmDelete_Click()
If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
    If Pay_rec = "rec" Then
    con.Execute "delete from receipt where recno=" & Me.txtrecNo.text & ""
    Else
    con.Execute "delete from payment where recno=" & Me.txtrecNo.text & ""
    End If
    Call cmdRef_Click
End If
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdPrint_Click()
If Pay_rec = "rec" Then
   cr.Reset
   cr.ReportFileName = App.Path & "\receipt.rpt"
Else
   cr.ReportFileName = App.Path & "\payment.rpt"
   End If
   cr.ReplaceSelectionFormula "{Receipt.RecNo}=" & txtrecNo.text & ""
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1
   
End Sub

Private Sub cmdRef_Click()
Me.txtAmt.text = 0
Me.txtBank.text = ""
Me.txtCheckno.text = ""
Me.txtCustomer.text = ""
headBal.Caption = ""

If Pay_rec = "rec" Then
 txtrecNo.text = MaxSNo("receipt", "recno")
Else
 txtrecNo.text = MaxSNo("payment", "recno")
End If

cboFirm.text = ""

txtCustomer.SetFocus
End Sub

Private Sub cmdSave_Click()


If txtCustomer.text = "" Then
MsgBox "Plz. Select Customer Name !!", vbInformation
Exit Sub
txtCustomer.SetFocus
End If
If Val(Me.txtAmt.text) = 0 Then
MsgBox "Plz. Enter Amount !!", vbInformation
Exit Sub
Me.txtAmt.SetFocus
End If

If Pay_rec = "rec" Then
con.Execute "delete from receipt where recno=" & Me.txtrecNo.text & ""
Else
con.Execute "delete from payment where recno=" & Me.txtrecNo.text & ""
End If

If MsgBox("Want to Save ?", vbQuestion + vbYesNo) = vbYes Then
save
End If



End Sub
Sub save()
If rs.State = 1 Then rs.Close

If Pay_rec = "rec" Then
rs.Open "select * from receipt", con, adOpenDynamic, adLockOptimistic
Else
rs.Open "select * from payment", con, adOpenDynamic, adLockOptimistic
End If

rs.AddNew
rs!RecNo = txtrecNo.text
rs!Dates = Me.recdate.Value
rs!CustomerName = Me.txtCustomer.text
rs!bank = Me.txtBank.text
rs!Cheque = Me.txtCheckno.text

If Me.cash.Value = True Then
rs!Bank_Cash = Me.cash.Caption
Else
rs!Bank_Cash = Me.bank.Caption
End If
rs!firm = cboFirm.text

If Me.Option1_cr.Value = True Then
   rs!Amount = Me.txtAmt.text
   rs!drcr = "c"
Else
   rs!PayAmt = Me.txtAmt.text
   rs!drcr = "d"
End If


rs.Update

Call cmdRef_Click
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys "{tab}"
End Sub
Private Sub Form_Load()

On Error Resume Next

pos Me




If Pay_rec = "rec" Then
txtrecNo.text = MaxSNo("receipt", "recno")
Me.Caption = "Receipt Entry"
Label7.Caption = "Receipt No"
Label1.Caption = "Receipt Date"
Label2.Caption = "Customer Name"
cboCustomer.ListIndex = 0
Else
txtrecNo.text = MaxSNo("payment", "recno")
Me.Caption = "Payment Entry"
Label7.Caption = "Payment No"
Label1.Caption = "Payment Date"
Label2.Caption = "Supplier Name"
cboCustomer.ListIndex = 0
End If

Me.recdate.Value = Date
Me.txtCheckno.text = "Cheque"

If rs.State = 1 Then rs.Close
rs.Open "select Name,Address from firm order by name", con
While rs.EOF = False
cboFirm.AddItem rs(0)
rs.MoveNext
Wend

cboFirm.ListIndex = 1

End Sub

Private Sub txtCustomer_GotFocus()
If PopUpValue1 <> "" Then
txtCustomer.text = PopUpValue1

ShowBalanace
SupplierLedger

headBal = ""
If rs.State = 1 Then rs.Close
rs.Open "select sum(Receipt),sum(Payment) from winrpt", con, adOpenKeyset, adLockReadOnly
If Not IsNull(rs(0)) Then
headBal.Caption = Round((rs(1) - rs(0)), 0)
End If



PopUpValue1 = ""
End If
End Sub
Sub SupplierLedger()

Screen.MousePointer = vbHourglass

Dim op, rec, payment As Double
Dim rs1 As New ADODB.Recordset
con.Execute "DELETE FROM winrpt where len(Receipt)>0"
If rs1.State = 1 Then rs1.Close

If Me.txtCustomer.text = "" Then
rs1.Open "SELECT DISTINCT(remarks) FROM invoicea", con
Else
rs1.Open "SELECT DISTINCT(remarks) FROM invoicea WHERE remarks='" & Me.txtCustomer.text & "'", con
End If
While rs1.EOF = False
    op = 0
    rec = 0
    payment = 0
    If rs.State = 1 Then rs.Close
    rs.Open "select sum(Amount) FROM receipt where CustomerName='" & rs1(0) & "'", con
    If Not IsNull(rs(0)) Then
    rec = rs(0)
    Else
    rec = 0
    End If

    If rs.State = 1 Then rs.Close
    rs.Open "select sum(netamount) FROM invoicea where  remarks='" & rs1(0) & "'", con, adOpenKeyset
    If Not IsNull(rs(0)) Then
    payment = rs(0)
    Else
    payment = 0
    End If

    op = payment - rec
    rec = 0
    payment = 0
    If op < 0 Then
        payment = op
    Else
        rec = op
    End If
    
    
    
    'con.Execute "insert into winrpt(Narration,Receipt,Payment,Description,FromDate,ToDate,party) values( '" & "Opening Balance as on " & recdate.Value & "' ," & rec & "," & payment & ",'" & "SUB LEDGER ACCOUNT" & "','" & recdate.Value & "','" & recdate.Value & "','" & rs1(0) & "')"
    con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party)  SELECT Dates,str(RecNo) + space(10-len(RecNo)) + iif(Bank_Cash='Bank','','Cash') + bank + space(10-len(bank)) + Cheque as rec1,Amount,CustomerName FROM receipt WHERE CustomerName='" & txtCustomer & "'"
    'con.Execute "INSERT INTO WINRPT (date1,Narration,payment,party,MRVNo,opdes,dr,Qty,op)  SELECT invoiceDate,'',netamount,remarks,invoiceno,'',rate  FROM invoicea WHERE remarks='" & txtCustomer & "'  group by invoiceDate,invoiceDate,remarks,EntryNo,Descr,Qty,NoOfType,TypeOfFold,Rate"
    con.Execute "INSERT INTO WINRPT (date1,payment,party)  SELECT invoiceDate,netamount,remarks  FROM invoicea WHERE remarks='" & txtCustomer & "'  group by invoiceDate,netamount,remarks"
    
    op = 0
    
rs1.MoveNext
Wend

Screen.MousePointer = vbDefault


End Sub


Sub ShowBalanace()
''Me.headBal.Caption = 0
''Dim sum, sum1 As Double
''Dim Receive, payment As Double
''
''sum = 0
''sum1 = 0
''
''payment = 0
''Receive = 0
''
''
''If rs.State = 1 Then rs.Close
''rs.Open "select sum(net) from IssueConsumeItem where customerName='" & Me.txtCustomer.Text & "' and ReturnQty=False", con
''If Not IsNull(rs(0)) Then
''   Receive = rs(0)
''Else
''   Receive = 0
''End If
''
''If rs.State = 1 Then rs.Close
''rs.Open "select sum(net) from IssueConsumeItem where customerName='" & Me.txtCustomer.Text & "' and ReturnQty=true", con
''If Not IsNull(rs(0)) Then
''   payment = rs(0)
''Else
''   payment = 0
''End If
''
''
''
''
''If rs.State = 1 Then rs.Close
''rs.Open "select sum(Amount) from receipt where customerName='" & Me.txtCustomer.Text & "'", con
''If Not IsNull(rs(0)) Then
''   Receive = Receive + rs(0)
''End If
''
''
''If rs.State = 1 Then rs.Close
''rs.Open "select sum(Amount) from payment where customerName='" & Me.txtCustomer.Text & "'", con
''If Not IsNull(rs(0)) Then
''   payment = payment + rs(0)
''End If
''
''
''sum = (Receive - payment)
''Me.headBal.Caption = sum





End Sub
Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
'If (KeyCode = 255 Or KeyCode = 113) Then
'If Pay_rec = "rec" Then
'popuplist2 "Select name from Customer order by Name", con
'Else
'If cboCustomer.Text = "Customer" Then
'popuplist2 "Select name from Customer order by Name", con
'Else
If KeyCode = 113 Then
'popuplist2 "Select name from Supplier order by Name", con
popuplist2 "select Name,Address from customer order by name", con
End If
'End If
'End If
End Sub
Private Sub txtrecNo_GotFocus()

If PopUpValue1 <> "" Then
txtrecNo.text = PopUpValue1

If rs.State = 1 Then rs.Close
If Pay_rec = "rec" Then
rs.Open "select * from receipt where recno=" & txtrecNo.text & "", con
Else
rs.Open "select * from payment where recno=" & txtrecNo.text & "", con
End If

If rs.EOF = False Then
Me.recdate.Value = rs!Dates
Me.txtCustomer.text = rs!CustomerName & ""
Me.txtBank.text = rs!bank & ""
Me.txtCheckno.text = rs!Cheque & ""
'Me.txtAmt.text = rs!Amount & ""
Me.cboFirm.text = rs!firm & ""

If rs!Bank_Cash = "Cash" Then
Me.cash.Value = True
Me.bank.Value = False
Else
Me.cash.Value = False
Me.bank.Value = True
End If



If (rs!drcr = "c" Or IsNull(rs!drcr)) Then
   Me.txtAmt.text = rs!Amount
   Me.Option1_cr.Value = True
   
Else
   Me.txtAmt.text = rs!PayAmt
   Me.Option2_dr.Value = 1
End If


ShowBalanace
PopUpValue1 = ""
End If

End If

End Sub

Private Sub txtrecNo_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 255 Or KeyCode = 113) Then
If Pay_rec = "rec" Then
popuplist2 "Select RecNo,Dates,CustomerName,Amount from receipt order by RecNo", con
Else
popuplist2 "Select RecNo,Dates,CustomerName,Amount from payment order by RecNo", con
End If
End If
End Sub

