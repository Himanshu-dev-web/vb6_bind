VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReceipt 
   Caption         =   "Receipt Form"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboCustomer 
      Height          =   315
      ItemData        =   "frmReceipt.frx":0000
      Left            =   5580
      List            =   "frmReceipt.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin Crystal.CrystalReport cr 
      Left            =   1620
      Top             =   4020
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtrecNo 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cdmDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1275
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3960
      Width           =   1275
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Top             =   2400
      Width           =   1515
   End
   Begin VB.OptionButton cash 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   1980
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton bank 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Top             =   1980
      Width           =   975
   End
   Begin VB.TextBox txtCheckno 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6720
      TabIndex        =   6
      Top             =   2460
      Width           =   2595
   End
   Begin VB.TextBox txtBank 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6720
      TabIndex        =   5
      Top             =   2040
      Width           =   2595
   End
   Begin VB.TextBox txtCustomer 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   1380
      Width           =   4035
   End
   Begin MSComCtl2.DTPicker recdate 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   39548
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2640
      TabIndex        =   21
      Top             =   60
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Customer Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2640
      TabIndex        =   20
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Receipt No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   19
      Top             =   360
      Width           =   1755
   End
   Begin VB.Label headBal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   18
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Balance Amount :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6780
      TabIndex        =   17
      Top             =   1380
      Width           =   2115
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Amount "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   13
      Top             =   2400
      Width           =   1755
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cheque No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   12
      Top             =   2460
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Bank Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   11
      Top             =   2040
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Customer Name "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   1380
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Receipt Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   9
      Top             =   720
      Width           =   1755
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bank_Click()

If bank.Value = True Then
    Me.txtBank.Enabled = True
    Me.txtCheckno.Enabled = True
    Me.txtCheckno.Text = "Cheque"
End If

End Sub
Private Sub cash_Click()
If cash.Value = True Then
Me.txtBank.Enabled = False
Me.txtCheckno.Enabled = False
Me.txtCheckno.Text = ""
End If
End Sub
Private Sub cdmDelete_Click()
If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
If Pay_rec = "rec" Then
con.Execute "delete from Receipt where recno=" & Me.txtrecNo.Text & ""
Else
con.Execute "delete from payment where recno=" & Me.txtrecNo.Text & ""
End If
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
   cr.ReplaceSelectionFormula "{Receipt.RecNo}=" & txtrecNo.Text & ""
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1
   
End Sub

Private Sub cmdRef_Click()
Me.txtAmt.Text = 0
Me.txtBank.Text = ""
Me.txtCheckno.Text = ""
Me.txtCustomer.Text = ""
headBal.Caption = ""

If Pay_rec = "rec" Then
 txtrecNo.Text = MaxSNo("receipt", "recno")
Else
 txtrecNo.Text = MaxSNo("payment", "recno")
End If

txtCustomer.SetFocus
End Sub

Private Sub cmdSave_Click()


If txtCustomer.Text = "" Then
MsgBox "Plz. Select Customer Name !!", vbInformation
Exit Sub
txtCustomer.SetFocus
End If
If Val(Me.txtAmt.Text) = 0 Then
MsgBox "Plz. Enter Amount !!", vbInformation
Exit Sub
Me.txtAmt.SetFocus
End If

If Pay_rec = "rec" Then
con.Execute "delete from Receipt where recno=" & Me.txtrecNo.Text & ""
Else
con.Execute "delete from payment where recno=" & Me.txtrecNo.Text & ""
End If

If MsgBox("Want to Save ?", vbQuestion + vbYesNo) = vbYes Then
save
End If



End Sub
Sub save()
If rs.State = 1 Then rs.Close

If Pay_rec = "rec" Then
rs.Open "select * from Receipt", con, adOpenDynamic, adLockOptimistic
Else
rs.Open "select * from payment", con, adOpenDynamic, adLockOptimistic
End If

rs.AddNew
rs!RecNo = txtrecNo.Text
rs!Dates = Me.recdate.Value
rs!CustomerName = Me.txtCustomer.Text
rs!bank = Me.txtBank.Text
rs!Cheque = Me.txtCheckno.Text
rs!Amount = Me.txtAmt.Text
If Me.cash.Value = True Then
rs!Bank_Cash = Me.cash.Caption
Else
rs!Bank_Cash = Me.bank.Caption
End If
rs.Update
Call cmdRef_Click
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub Form_Load()
pos Me




If Pay_rec = "rec" Then
txtrecNo.Text = MaxSNo("receipt", "recno")
Me.Caption = "Receipt Entry"
Label7.Caption = "Receipt No"
Label1.Caption = "Receipt Date"
Label2.Caption = "Customer Name"
cboCustomer.ListIndex = 0
Else
txtrecNo.Text = MaxSNo("payment", "recno")
Me.Caption = "Payment Entry"
Label7.Caption = "Payment No"
Label1.Caption = "Payment Date"
Label2.Caption = "Supplier Name"
cboCustomer.ListIndex = 0
End If

Me.recdate.Value = Date
Me.txtCheckno.Text = "Cheque"
End Sub

Private Sub txtCustomer_GotFocus()
If PopUpValue1 <> "" Then
txtCustomer.Text = PopUpValue1

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
    rs.Open "select sum(Amount) FROM payment where CustomerName='" & rs1(0) & "'", con
    If Not IsNull(rs(0)) Then
    rec = rs(0)
    Else
    rec = 0
    End If

    If rs.State = 1 Then rs.Close
    rs.Open "select sum(amount) FROM FoldingDetails where  PartyName='" & rs1(0) & "'", con, adOpenKeyset
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
    con.Execute "INSERT INTO WINRPT (date1,Narration,Receipt,party)  SELECT Dates,str(RecNo) + space(10-len(RecNo)) + iif(Bank_Cash='Bank','','Cash') + bank + space(10-len(bank)) + Cheque as rec1,Amount,CustomerName FROM payment WHERE CustomerName='" & txtCustomer & "'"
    con.Execute "INSERT INTO WINRPT (date1,Narration,payment,party,MRVNo,opdes,dr,Qty,op)  SELECT entDate,Descr,amount,PartyName,EntryNo,str(Qty)+ ' X' +str(NoOfType),TypeOfFold,(Qty*NoOfType),rate  FROM FoldingDetails WHERE PartyName='" & txtCustomer & "'  group by entDate,amount,PartyName,EntryNo,Descr,Qty,NoOfType,TypeOfFold,Rate"
    
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
If KeyCode = 113 Then
If Pay_rec = "rec" Then
popuplist2 "Select name from Customer order by Name", con
Else
If cboCustomer.Text = "Customer" Then
popuplist2 "Select name from Customer order by Name", con
Else
popuplist2 "Select name from Supplier order by Name", con
End If
End If
End If
End Sub
Private Sub txtrecNo_GotFocus()

If PopUpValue1 <> "" Then
txtrecNo.Text = PopUpValue1

If rs.State = 1 Then rs.Close
If Pay_rec = "rec" Then
rs.Open "select * from Receipt where recno=" & txtrecNo.Text & "", con
Else
rs.Open "select * from payment where recno=" & txtrecNo.Text & "", con
End If

If rs.EOF = False Then
Me.recdate.Value = rs!Dates
Me.txtCustomer.Text = rs!CustomerName & ""
Me.txtBank.Text = rs!bank & ""
Me.txtCheckno.Text = rs!Cheque & ""
Me.txtAmt.Text = rs!Amount & ""
If rs!Bank_Cash = "Cash" Then
Me.cash.Value = True
Me.bank.Value = False
Else
Me.cash.Value = False
Me.bank.Value = True
End If


ShowBalanace
PopUpValue1 = ""
End If

End If

End Sub

Private Sub txtrecNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
If Pay_rec = "rec" Then
popuplist2 "Select RecNo,Dates,CustomerName,Amount from receipt order by RecNo", con
Else
popuplist2 "Select RecNo,Dates,CustomerName,Amount from payment order by RecNo", con
End If
End If
End Sub
