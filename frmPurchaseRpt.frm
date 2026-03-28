VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPurchaseRpt 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Date Wise Purchase"
   ClientHeight    =   2445
   ClientLeft      =   4395
   ClientTop       =   2940
   ClientWidth     =   3450
   Icon            =   "frmPurchaseRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3450
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   1050
      TabIndex        =   5
      Top             =   1665
      Width           =   1890
   End
   Begin Crystal.CrystalReport cr 
      Left            =   105
      Top             =   1650
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   1050
      TabIndex        =   0
      Top             =   1140
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker FromDate 
      Height          =   300
      Left            =   1050
      TabIndex        =   1
      Top             =   300
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   59965443
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   59965443
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   255
      Left            =   225
      TabIndex        =   4
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      Height          =   255
      Left            =   225
      TabIndex        =   3
      Top             =   780
      Width           =   1215
   End
End
Attribute VB_Name = "frmPurchaseRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub CmdPrint_Click()
   
  If frmMain.k = 1 Then
   cr.Reset
   cr.ReportFileName = App.Path & "\PurchaseReg.rpt"
   cr.ReplaceSelectionFormula "{Purchase.RecDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {Purchase.RecDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "todate='" & todate.Value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
  ElseIf frmMain.k = 2 Then
   cr.Reset
   cr.ReportFileName = App.Path & "\consumvablePurchaseRpt.rpt"
   cr.ReplaceSelectionFormula "{ConsumvablePurchaseRegister.billdate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {ConsumvablePurchaseRegister.billdate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "todate='" & todate.Value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
  ElseIf frmMain.k = 3 Then
   cr.Reset
   cr.ReportFileName = App.Path & "\Dispatch.rpt"
   cr.ReplaceSelectionFormula "{invoice.recdate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {invoice.recdate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
  ElseIf frmMain.k = 4 Then
'   CR.Reset
'   CR.ReportFileName = App.Path & "\rawstockregister.rpt"
'   CR.ReplaceSelectionFormula "{RowStockSummaryRegister.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {RowStockSummaryRegister.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
'   CR.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
'   CR.Formulas(1) = "Todate='" & ToDate.Value & "'"
'   CR.WindowState = crptMaximized
'   CR.Action = 1
  
   ElseIf frmMain.k = 5 Then
   cr.Reset
   cr.ReportFileName = App.Path & "\saleinvoicesummery.rpt"
   cr.ReplaceSelectionFormula "{salesinvoice.invoicedate} >=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {salesinvoice.invoicedate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
   
   ElseIf frmMain.k = 6 Then
   cr.Reset
   cr.ReportFileName = App.Path & "\TaxInvoiceSummery.rpt"
   cr.ReplaceSelectionFormula "{taxinvoice.invoicedate} >=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {taxinvoice.invoicedate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
   
    ElseIf frmMain.k = 7 Then
   cr.Reset
   cr.ReportFileName = App.Path & "\challanInvoiceSummery.rpt"
   cr.ReplaceSelectionFormula "{ChallanTransferInvoice.invoicedate} >=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {ChallanTransferInvoice.invoicedate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
   
   
   ElseIf frmMain.k = 8 Then
   cr.Reset
   cr.ReportFileName = App.Path & "\taxSummery.rpt"
   cr.ReplaceSelectionFormula "{taxInvoice.invoicedate} >=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {taxInvoice.invoicedate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
   
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 7 Then Unload Me
End Sub

Private Sub Form_Load()
   On Error Resume Next
   fromDate.Value = Date
   todate.Value = Date
   pos Me
End Sub


