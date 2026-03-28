VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmVatRegister 
   Caption         =   "VAT Register"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmVatRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   2700
      TabIndex        =   2
      Top             =   1740
      Width           =   1710
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1740
      Width           =   1485
   End
   Begin VB.ComboBox cbofirm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   375
      Left            =   3060
      TabIndex        =   3
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   64487425
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   375
      Left            =   1260
      TabIndex        =   4
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   64487425
      CurrentDate     =   39545
   End
   Begin Crystal.CrystalReport cr 
      Left            =   240
      Top             =   2100
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   900
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Left            =   2820
      TabIndex        =   6
      Top             =   900
      Width           =   555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Firm Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "frmVatRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
  Unload Me
End Sub
Sub customerLedger()
''Screen.MousePointer = vbHourglass
''Dim op, rec, payment As Double
''Dim rs1 As New ADODB.Recordset
''con.Execute "DELETE FROM winrpt"
''If rs1.State = 1 Then rs1.Close
''
''If Me.txtCustomer.Text = "" Then
''rs1.Open "SELECT DISTINCT(Name) FROM Customer", con
''Else
''rs1.Open "SELECT DISTINCT(Name) FROM Customer WHERE Name='" & Me.txtCustomer.Text & "'", con
''End If
''
''While rs1.EOF = False
''    op = 0
''    rec = 0
''    payment = 0
''    If rs.State = 1 Then rs.Close
''    rs.Open "select sum(Amount) FROM Receipt where Dates<datevalue('" & Me.fromDate.Value & "') AND CustomerName='" & rs1(0) & "'", con
''    If Not IsNull(rs(0)) Then
''    rec = rs(0)
''    Else
''    rec = 0
''    End If
''
''    'If rs.State = 1 Then rs.Close
''    'rs.Open "select sum(NET) FROM IssueConsumeItem where IssueDate<datevalue('" & Me.fromDate.Value & "') AND CustomerName='" & rs1(0) & "' and ReturnQty=false", con, adOpenKeyset
''    'If Not IsNull(rs(0)) Then
''    'payment = rs(0)
''    'Else
''    'payment = 0
''    'End If
''
''    If rs.State = 1 Then rs.Close
''    rs.Open "select sum(NET) FROM IssueConsumeItem where IssueDate<datevalue('" & Me.fromDate.Value & "') AND CustomerName='" & rs1(0) & "' and ReturnQty=true", con, adOpenKeyset
''    If Not IsNull(rs(0)) Then
''    rec = rec + rs(0)
''    End If
''
''
''    op = rec - payment
''    rec = 0
''    payment = 0
''    If op < 0 Then
''        payment = op
''    Else
''        rec = op
''    End If
''
''    con.Execute "insert into winrpt(Narration,Payment,Receipt,Description,FromDate,ToDate,party) values( '" & "Opening Balance as on " & Me.todate.Value & "' ," & rec & "," & payment & ",'" & "SUB LEDGER ACCOUNT" & "','" & Me.fromDate.Value & "','" & Me.todate.Value & "','" & rs1(0) & "')"
''    con.Execute "INSERT INTO WINRPT (date1,Narration,Payment,party)  SELECT Dates,str(RecNo) + space(10-len(RecNo)) + iif(Bank_Cash='Bank','','Cash') + bank + space(10-len(bank)) + Cheque as rec1,Amount,CustomerName FROM RECEIPT WHERE CustomerName='" & rs1(0) & "' and Dates>=datevalue('" & Me.fromDate.Value & "')"
''    con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party)  SELECT Dates,str(RecNo) + space(10-len(RecNo)) + iif(Bank_Cash='Bank','','Cash') + bank + space(10-len(bank)) + Cheque as rec1,Amount,CustomerName FROM payment WHERE CustomerName='" & rs1(0) & "' and Dates>=datevalue('" & Me.fromDate.Value & "')"
''
''    con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party,IssueNo)  SELECT IssueDate,IssueNo,Net,customerName,IssueNo FROM IssueConsumeItem WHERE CustomerName='" & rs1(0) & "' and IssueDate>=datevalue('" & Me.fromDate.Value & "') and ReturnQty=False  group by IssueDate,TempData,Net,customerName,IssueNo"
''    con.Execute "INSERT INTO WINRPT (date1,Narration,Payment,party,IssueNo)  SELECT IssueDate,IssueNo,Net,customerName,IssueNo FROM IssueConsumeItem WHERE CustomerName='" & rs1(0) & "' and IssueDate>=datevalue('" & Me.fromDate.Value & "') and ReturnQty=true group by IssueDate,TempData,Net,customerName,IssueNo"
''
''    op = 0
''
''rs1.MoveNext
''Wend
''
''
''
''Screen.MousePointer = vbDefault


End Sub
Sub SupplierLedger()

''Screen.MousePointer = vbHourglass
''
''Dim op, rec, payment As Double
''Dim rs1 As New ADODB.Recordset
''Dim p As String
''Dim b12 As Boolean
''
''b12 = False
''
''con.Execute "DELETE FROM winrpt where len(Party)>0"
''
''If rs1.State = 1 Then rs1.Close
''
''If Me.txtCustomer.Text = "" Then
'' rs1.Open "SELECT DISTINCT remarks FROM invoicea where SUBLEDGER='" & Trim(cbofirm) & "'", con
''Else
'' rs1.Open "SELECT DISTINCT remarks FROM invoicea WHERE remarks='" & Me.txtCustomer.Text & "' and SUBLEDGER='" & Trim(cbofirm) & "' ", con
''End If
''
''While rs1.EOF = False
''
''    op = 0
''    rec = 0
''    payment = 0
''
''
''    If rs.State = 1 Then rs.Close
''    If Me.txtCustomer.Text <> "" Then
''    rs.Open "select sum(Amount) FROM receipt where Dates<datevalue('" & Me.fromDate.Value & "') AND CustomerName='" & rs1(0) & "' and firm='" & cbofirm & "'", con
''    Else
''    rs.Open "select sum(Amount) FROM receipt where Dates<datevalue('" & Me.fromDate.Value & "') and firm='" & cbofirm & "'", con
''    End If
''
''    If Not IsNull(rs(0)) Then
''    rec = rs(0)
''    Else
''    rec = 0
''    End If
''
''    If rs.State = 1 Then rs.Close
''    If Me.txtCustomer.Text <> "" Then
''    rs.Open "select sum(NETAMOUNT) FROM invoicea where INVOICEDATE<datevalue('" & Me.fromDate.Value & "') AND remarks='" & rs1(0) & "' and SUBLEDGER='" & Trim(cbofirm) & "'", con, adOpenKeyset
''    Else
''    rs.Open "select sum(NETAMOUNT) FROM invoicea where INVOICEDATE<datevalue('" & Me.fromDate.Value & "') and SUBLEDGER='" & Trim(cbofirm) & "'", con, adOpenKeyset
''    End If
''
''    If Not IsNull(rs(0)) Then
''    payment = rs(0)
''    Else
''    payment = 0
''    End If
''
''    op = payment - rec
''    rec = 0
''    payment = 0
''    If op > 0 Then
''        rec = op
''    Else
''        payment = op
''    End If
''    p = rs1(0)
''
''    If b12 = False Then
''       con.Execute "insert into winrpt(Narration,Receipt,Payment,Description,FromDate,ToDate,party) values( '" & "Opening Balance as on " & Me.todate.Value & "' ," & rec & "," & payment & ",'" & "SUB LEDGER ACCOUNT" & "','" & Me.fromDate.Value & "','" & Me.todate.Value & "','" & rs1(0) & "')"
''       b12 = True
''    End If
''
''    con.Execute "INSERT INTO WINRPT (date1,Narration,payment,party)  SELECT Dates,str(RecNo) + space(10-len(RecNo)) + iif(Bank_Cash='Bank','','Cash') + bank + space(10-len(bank)) + Cheque as rec1,Amount,CustomerName FROM receipt WHERE  firm='" & cbofirm & "' and CustomerName='" & rs1(0) & "' and (Dates>=datevalue('" & Me.fromDate.Value & "') and Dates<=datevalue('" & Me.todate.Value & "'))"
''
''    If txtCustomer = "" Then
''      con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party,MRVNo,opdes,dr,Qty,op,uid)  SELECT INVOICEDATE,(str(INVOICENO)+ ' - ' + str(INVOICEDATE)) + ' - ' + remarks ,NETAMOUNT,remarks,INVOICENO,'-','-',0,0,INVOICENO  FROM invoicea WHERE SUBLEDGER='" & Trim(cbofirm) & "' and remarks='" & rs1(0) & "' and (INVOICEDATE>=datevalue('" & Me.fromDate.Value & "') and INVOICEDATE<=datevalue('" & Me.todate.Value & "'))  group by INVOICEDATE,NETAMOUNT,remarks,INVOICENO"
''    Else
''     con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party,MRVNo,opdes,dr,Qty,op,uid)  SELECT INVOICEDATE,(str(INVOICENO)+ ' - ' + str(INVOICEDATE)) ,NETAMOUNT,remarks,INVOICENO,'-','-',0,0,INVOICENO  FROM invoicea WHERE SUBLEDGER='" & Trim(cbofirm) & "' and remarks='" & rs1(0) & "' and (INVOICEDATE>=datevalue('" & Me.fromDate.Value & "') and INVOICEDATE<=datevalue('" & Me.todate.Value & "'))  group by INVOICEDATE,NETAMOUNT,remarks,INVOICENO"
''    End If
''
''    op = 0
''
''rs1.MoveNext
''Wend
''
''
''Screen.MousePointer = vbDefault


End Sub

Private Sub CmdPrint_Click()


'If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
 
If cbofirm <> "" Then

If rs.State = 1 Then rs.Close
rs.Open "select cname from setup1", con
If rs.EOF = False Then
   If rs!CNAME = "ILYAS BOOK BINDING WORKS" Then
        cr.Reset
        cr.ReportFileName = App.Path & "\vatregister.rpt"
        'cr.ReplaceSelectionFormula "{INVOICEA.subledger}='" & cboFirm & "'"
        cr.ReplaceSelectionFormula "{INVOICEA.subledger}='" & cbofirm & "' and {invoicea.invoiceDate}>=datevalue('" & Format(fromDate.Value, "dd/MM/yyyy") & "') and {invoicea.invoiceDate}<=datevalue('" & Format(todate.Value, "dd/MM/yyyy") & "')"
        cr.WindowShowPrintSetupBtn = True
        cr.WindowState = crptMaximized
        cr.WindowShowPrintBtn = True
        cr.Formulas(0) = "fromDate='" & Me.fromDate.Value & "'"
        cr.Formulas(1) = "toDate='" & Me.todate.Value & "'"
        cr.Action = 1
   Else
        cr.Reset
        cr.ReportFileName = App.Path & "\vatregister.rpt"
        'cr.ReplaceSelectionFormula "{INVOICEA.subledger}='" & cboFirm & "'"
        cr.ReplaceSelectionFormula "{INVOICEA.subledger}='" & cbofirm & "' and {invoicea.invoiceDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {invoicea.invoiceDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
        cr.WindowShowPrintSetupBtn = True
        cr.WindowState = crptMaximized
        cr.WindowShowPrintBtn = True
        cr.Formulas(0) = "fromDate='" & Me.fromDate.Value & "'"
        cr.Formulas(1) = "toDate='" & Me.todate.Value & "'"
        cr.Action = 1
   
   End If
End If
 

End If




End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 7 Then Unload Me
   If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub Form_Load()
   On Error Resume Next
   fromDate.Value = Date
   todate.Value = Date
   pos Me
   
''   If Pay_rec = "rec" Then
''   Me.Caption = "Customer Ledger"
''   Else
''   Me.Caption = "Supplier Ledger"
''   End If
''
If rs.State = 1 Then rs.Close
rs.Open "select Name,Address from firm order by name", con
While rs.EOF = False
cbofirm.AddItem rs(0)
rs.MoveNext
Wend

cbofirm.ListIndex = 0


If rs.State = 1 Then rs.Close
rs.Open "select yarfrom,yarto from setup1", con
If rs.EOF = False Then
   fromDate.Value = rs(0)
   todate.Value = rs(1)
End If



   
End Sub
Private Sub txtCustomer_GotFocus()
''If PopUpValue1 <> "" Then
''Me.txtCustomer.Text = PopUpValue1
''PopUpValue1 = ""
''End If
End Sub
Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)

If (KeyCode = 113) Then

'If Pay_rec = "rec" Then
'popuplist2 "Select name from Customer order by Name", con
'Else
'popuplist2 "Select name from Supplier order by Name", con
'End If
popuplist2 "select Name,Address from customer order by name", con


End If

End Sub


