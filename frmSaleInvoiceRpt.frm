VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSaleInvoiceRpt 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Date Wise Purchase"
   ClientHeight    =   2445
   ClientLeft      =   4395
   ClientTop       =   2940
   ClientWidth     =   6585
   Icon            =   "frmSaleInvoiceRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   6585
   Begin VB.TextBox txtCustomer 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   900
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   300
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19660801
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   300
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19660801
      CurrentDate     =   39545
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3060
      TabIndex        =   6
      Top             =   1560
      Width           =   1710
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1485
   End
   Begin Crystal.CrystalReport cr 
      Left            =   120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Name :"
      Height          =   315
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   255
      Left            =   540
      TabIndex        =   5
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmSaleInvoiceRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
  Unload Me
End Sub
Sub customerLedger()
Screen.MousePointer = vbHourglass
Dim op, rec, payment As Double
Dim rs1 As New ADODB.Recordset
con.Execute "DELETE FROM winrpt"
If rs1.State = 1 Then rs1.Close

If Me.txtCustomer.Text = "" Then
rs1.Open "SELECT DISTINCT(Name) FROM Customer", con
Else
rs1.Open "SELECT DISTINCT(Name) FROM Customer WHERE Name='" & Me.txtCustomer.Text & "'", con
End If

While rs1.EOF = False
    op = 0
    rec = 0
    payment = 0
    If rs.State = 1 Then rs.Close
    rs.Open "select sum(Amount) FROM Receipt where Dates<datevalue('" & Me.fromDate.Value & "') AND CustomerName='" & rs1(0) & "'", con
    If Not IsNull(rs(0)) Then
    rec = rs(0)
    Else
    rec = 0
    End If

    If rs.State = 1 Then rs.Close
    rs.Open "select sum(NET) FROM IssueConsumeItem where IssueDate<datevalue('" & Me.fromDate.Value & "') AND CustomerName='" & rs1(0) & "' and ReturnQty=false", con, adOpenKeyset
    If Not IsNull(rs(0)) Then
    payment = rs(0)
    Else
    payment = 0
    End If
    
    If rs.State = 1 Then rs.Close
    rs.Open "select sum(NET) FROM IssueConsumeItem where IssueDate<datevalue('" & Me.fromDate.Value & "') AND CustomerName='" & rs1(0) & "' and ReturnQty=true", con, adOpenKeyset
    If Not IsNull(rs(0)) Then
    rec = rec + rs(0)
    End If


    op = rec - payment
    rec = 0
    payment = 0
    If op < 0 Then
        payment = op
    Else
        rec = op
    End If
    
    con.Execute "insert into winrpt(Narration,Payment,Receipt,Description,FromDate,ToDate,party) values( '" & "Opening Balance as on " & Me.todate.Value & "' ," & rec & "," & payment & ",'" & "SUB LEDGER ACCOUNT" & "','" & Me.fromDate.Value & "','" & Me.todate.Value & "','" & rs1(0) & "')"
    con.Execute "INSERT INTO WINRPT (date1,Narration,Payment,party)  SELECT Dates,str(RecNo) + space(10-len(RecNo)) + iif(Bank_Cash='Bank','','Cash') + bank + space(10-len(bank)) + Cheque as rec1,Amount,CustomerName FROM RECEIPT WHERE CustomerName='" & rs1(0) & "' and Dates>=datevalue('" & Me.fromDate.Value & "')"
    con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party)  SELECT Dates,str(RecNo) + space(10-len(RecNo)) + iif(Bank_Cash='Bank','','Cash') + bank + space(10-len(bank)) + Cheque as rec1,Amount,CustomerName FROM payment WHERE CustomerName='" & rs1(0) & "' and Dates>=datevalue('" & Me.fromDate.Value & "')"
  
    con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party,IssueNo)  SELECT IssueDate,IssueNo,Net,customerName,IssueNo FROM IssueConsumeItem WHERE CustomerName='" & rs1(0) & "' and IssueDate>=datevalue('" & Me.fromDate.Value & "') and ReturnQty=False  group by IssueDate,TempData,Net,customerName,IssueNo"
    con.Execute "INSERT INTO WINRPT (date1,Narration,Payment,party,IssueNo)  SELECT IssueDate,IssueNo,Net,customerName,IssueNo FROM IssueConsumeItem WHERE CustomerName='" & rs1(0) & "' and IssueDate>=datevalue('" & Me.fromDate.Value & "') and ReturnQty=true group by IssueDate,TempData,Net,customerName,IssueNo"
    
    op = 0
    
rs1.MoveNext
Wend



Screen.MousePointer = vbDefault


End Sub
Sub SupplierLedger()
Screen.MousePointer = vbHourglass
Dim op, rec, payment As Double
Dim rs1 As New ADODB.Recordset
con.Execute "DELETE FROM winrpt"
If rs1.State = 1 Then rs1.Close
If Me.txtCustomer.Text = "" Then
rs1.Open "SELECT DISTINCT(PartyName) FROM FoldingDetails", con
Else
rs1.Open "SELECT DISTINCT(PartyName) FROM FoldingDetails WHERE PartyName='" & Me.txtCustomer.Text & "'", con
End If
While rs1.EOF = False
    op = 0
    rec = 0
    payment = 0
    If rs.State = 1 Then rs.Close
    rs.Open "select sum(Amount) FROM payment where Dates<datevalue('" & Me.fromDate.Value & "') AND CustomerName='" & rs1(0) & "'", con
    If Not IsNull(rs(0)) Then
    rec = rs(0)
    Else
    rec = 0
    End If

    If rs.State = 1 Then rs.Close
    rs.Open "select sum(amount) FROM FoldingDetails where EntDate<datevalue('" & Me.fromDate.Value & "') AND PartyName='" & rs1(0) & "'", con, adOpenKeyset
    If Not IsNull(rs(0)) Then
    payment = rs(0)
    Else
    payment = 0
    End If

    op = payment - rec
    rec = 0
    payment = 0
    If op > 0 Then
        payment = op
    Else
        rec = op
    End If
    
    con.Execute "insert into winrpt(Narration,Receipt,Payment,Description,FromDate,ToDate,party) values( '" & "Opening Balance as on " & Me.todate.Value & "' ," & rec & "," & payment & ",'" & "SUB LEDGER ACCOUNT" & "','" & Me.fromDate.Value & "','" & Me.todate.Value & "','" & rs1(0) & "')"
    con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party)  SELECT Dates,str(RecNo) + space(10-len(RecNo)) + iif(Bank_Cash='Bank','','Cash') + bank + space(10-len(bank)) + Cheque as rec1,Amount,CustomerName FROM payment WHERE CustomerName='" & rs1(0) & "' and Dates>=datevalue('" & Me.fromDate.Value & "')"
    con.Execute "INSERT INTO WINRPT (date1,Narration,payment,party,MRVNo,opdes,dr,Qty,op)  SELECT entDate,Descr,amount,PartyName,EntryNo,str(Qty)+ ' X' +str(NoOfType),TypeOfFold,(Qty*NoOfType),rate  FROM FoldingDetails WHERE PartyName='" & rs1(0) & "' and entDate>=datevalue('" & Me.fromDate.Value & "')  group by entDate,amount,PartyName,EntryNo,Descr,Qty,NoOfType,TypeOfFold,Rate"
    
    op = 0
    
rs1.MoveNext
Wend

Screen.MousePointer = vbDefault


End Sub

Private Sub CmdPrint_Click()

''If Pay_rec = "rec" Then
''customerLedger
''If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
'' cr.Reset
'' cr.ReportFileName = App.Path & "\Al_SLAccount.rpt"
'' cr.WindowShowPrintSetupBtn = True
'' cr.WindowState = crptMaximized
'' cr.WindowShowPrintBtn = True
'' cr.Formulas(0) = "fromDate='" & Me.FromDate.Value & "'"
'' cr.Formulas(1) = "toDate='" & Me.ToDate.Value & "'"
''
'' cr.Action = 1
'' End If
''
''Else
SupplierLedger
If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
 cr.Reset
 cr.ReportFileName = App.Path & "\Al_SLAccount1.rpt"
 cr.WindowShowPrintSetupBtn = True
 
 cr.WindowState = crptMaximized
 cr.WindowShowPrintBtn = True
 
 cr.Formulas(0) = "fromDate='" & Me.fromDate.Value & "'"
 cr.Formulas(1) = "toDate='" & Me.todate.Value & "'"
 
 cr.Formulas(2) = "party='" & txtCustomer & "'"
 'cr.Formulas(3) = "firm_='" & COMPNAME & "'"
 
 cr.Action = 1
 End If

'End If



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
   
   If Pay_rec = "rec" Then
   Me.Caption = "Customer Ledger"
   Else
   Me.Caption = "Supplier Ledger"
   End If
   
End Sub
Private Sub txtCustomer_GotFocus()
If PopUpValue1 <> "" Then
Me.txtCustomer.Text = PopUpValue1
PopUpValue1 = ""
End If
End Sub
Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then

If Pay_rec = "rec" Then
popuplist2 "Select name from Customer order by Name", con
Else
popuplist2 "Select name from Supplier order by Name", con
End If

End If
End Sub
