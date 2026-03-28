VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   Caption         =   "Book  Binding House"
   ClientHeight    =   9852
   ClientLeft      =   492
   ClientTop       =   456
   ClientWidth     =   16560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr 
      Left            =   45
      Top             =   90
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu menuMaster 
      Caption         =   "&Master"
      Begin VB.Menu menuCustomer 
         Caption         =   "Customer Details..."
      End
      Begin VB.Menu MenuUnit 
         Caption         =   "Folding Rate Master..."
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuBinder 
         Caption         =   "Printer Master ...."
      End
      Begin VB.Menu mnuMaster 
         Caption         =   "Group Master"
         Shortcut        =   ^G
      End
      Begin VB.Menu menuItem 
         Caption         =   "Book Master..."
         Shortcut        =   ^R
      End
      Begin VB.Menu menuSupplier 
         Caption         =   "Supplier Details"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu lineExit 
         Caption         =   "-"
      End
      Begin VB.Menu MENUTAX 
         Caption         =   "Tax Master..."
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVat 
         Caption         =   "VAT Master"
      End
      Begin VB.Menu mnuComp 
         Caption         =   "Company Master"
      End
      Begin VB.Menu mnureset 
         Caption         =   "Reset Database"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit..."
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menuTransiction 
      Caption         =   "Folding (Paper)"
      Begin VB.Menu mnuJobs 
         Caption         =   "Jobs Entry (Folding) ...."
      End
      Begin VB.Menu menuPay 
         Caption         =   "Payment Form.."
      End
      Begin VB.Menu mnuCust1 
         Caption         =   "Customer Ledger"
      End
      Begin VB.Menu mnurec 
         Caption         =   "Receipt Form..."
         Visible         =   0   'False
      End
      Begin VB.Menu menuIssueConsume 
         Caption         =   "Item Issue..."
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu menuIssueReturn 
         Caption         =   "Item Return..."
         Visible         =   0   'False
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu menuPurchase 
         Caption         =   "Raw Purchase Entry..."
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu menuRaw1 
         Caption         =   "Purchase Entry..."
         Visible         =   0   'False
      End
      Begin VB.Menu menuPurchaseReturn 
         Caption         =   "Purchase Return"
         Visible         =   0   'False
      End
      Begin VB.Menu lineIssue 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusalesinvoice 
         Caption         =   "Sales Invoice"
         Visible         =   0   'False
      End
      Begin VB.Menu mnutaxinvoice 
         Caption         =   "Tax Invoice"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPerformaInvoice 
         Caption         =   "Performa Invoice"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBill 
         Caption         =   "Bill"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCashMemo 
         Caption         =   "Cash Memo"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuChallanTransferInvoice 
         Caption         =   "Challan/Transfer Invoice"
         Visible         =   0   'False
      End
      Begin VB.Menu menuProcess 
         Caption         =   "Issue Raw I&tem ... "
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu lineSale 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuFinish 
         Caption         =   "Dispatch Form ..."
         Shortcut        =   ^F
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuNewBook1 
      Caption         =   "New Book"
      Begin VB.Menu mnuBookRec1 
         Caption         =   "Book Receied From Press ..."
      End
      Begin VB.Menu mnuBookDel 
         Caption         =   "Book Delivered ..."
      End
      Begin VB.Menu mnuNewBook 
         Caption         =   "New Book Ledger..."
      End
   End
   Begin VB.Menu mnuBook 
      Caption         =   "ITC Book"
      Begin VB.Menu mnuBookRec 
         Caption         =   "I.T.C.  Received ..."
      End
      Begin VB.Menu mnuConsumed 
         Caption         =   "I.T.C.  Delivered ..."
      End
      Begin VB.Menu mnuOldBookLedger 
         Caption         =   "I.T.C.  Book Ledger..."
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "Title ... "
      Begin VB.Menu mnuTitle1 
         Caption         =   "Title/Inner/Paper  Received ..."
      End
      Begin VB.Menu titleDelivered 
         Caption         =   "Title/Inner/Paper  Delivered ..."
      End
      Begin VB.Menu mnuTLedger1 
         Caption         =   "Title /Inner/Paper  Ledger ..."
      End
   End
   Begin VB.Menu mnuPTitle 
      Caption         =   "Print St. Of Title"
   End
   Begin VB.Menu mnuBilling 
      Caption         =   "Billing Option"
      Begin VB.Menu mnuBillent 
         Caption         =   "Bill Entry..."
      End
      Begin VB.Menu mnuChallan 
         Caption         =   "Challan Creation ..."
      End
      Begin VB.Menu mnuRec1 
         Caption         =   "Receipt/Payment Amount..."
      End
      Begin VB.Menu mnuCustomerL 
         Caption         =   "Customer Ledger..."
      End
      Begin VB.Menu mnuVat1 
         Caption         =   "GST Register..."
      End
      Begin VB.Menu mnuPendingBill 
         Caption         =   "Pending Bill..."
      End
   End
   Begin VB.Menu menurpt 
      Caption         =   "&Report"
      Visible         =   0   'False
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer Ledger ..."
         Visible         =   0   'False
      End
      Begin VB.Menu menuDateWisePurchase 
         Caption         =   "Raw Purchase Register..."
         Visible         =   0   'False
      End
      Begin VB.Menu menuconRegister 
         Caption         =   "Purchase Register..."
         Visible         =   0   'False
      End
      Begin VB.Menu menuDispatch 
         Caption         =   "Dispatch Register..."
         Visible         =   0   'False
      End
      Begin VB.Menu MENUTAXINVOICESUMMERY 
         Caption         =   "Tax Invoice Summery"
         Visible         =   0   'False
      End
      Begin VB.Menu menuchallaninvoicesummery 
         Caption         =   "Challan Invoice Summery"
         Visible         =   0   'False
      End
      Begin VB.Menu menuSummery 
         Caption         =   "Summery"
         Visible         =   0   'False
      End
      Begin VB.Menu menuStock1 
         Caption         =   "-"
      End
      Begin VB.Menu menuStockSummary 
         Caption         =   "Stock Summary..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSupp 
         Caption         =   "Supplier Ledger..."
         Visible         =   0   'False
      End
      Begin VB.Menu menuFinishSummary 
         Caption         =   "Stock Summary (Finish Item)..."
         Visible         =   0   'False
      End
      Begin VB.Menu menuasOn 
         Caption         =   "Summary of Raw Stock As On..."
         Visible         =   0   'False
      End
      Begin VB.Menu menusaleinvoicereport 
         Caption         =   "Sale Invoice Report"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Visible         =   0   'False
      Begin VB.Menu mnuCompany 
         Caption         =   "Company Master"
      End
      Begin VB.Menu mnuStudent 
         Caption         =   "student"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Data..."
      End
      Begin VB.Menu mnuRes 
         Caption         =   "Restore Data ..."
      End
   End
   Begin VB.Menu menuExit1 
      Caption         =   "E&xit Software"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K As Integer
Private Sub menuAdmin_Click()
frmAdmin.Show
End Sub
Private Sub MDIForm_Load()

    companyname
    pos Me
    Me.Caption = "Book Binding System  : " & frmPassword.cboyrs.text & "      -      " & " Developed By : 09997314681,7906524313"

End Sub
Private Sub menuasOn_Click()
    frmMonthlySummary.Show
End Sub
Private Sub menuchallaninvoicesummery_Click()
    
    K = 7
    frmPurchaseRpt.Caption = "CHALLAN INVOICE SUMMERY"
    frmPurchaseRpt.Show
    
End Sub
Private Sub menuconRegister_Click()
    K = 2
    frmPurchaseRpt.Show
End Sub
Private Sub menuCustomer_Click()
    Record = ""
    frmCustomer.Show
End Sub
Private Sub menuDateWisePurchase_Click()
K = 1
frmPurchaseRpt.Show
End Sub
Private Sub menuDispatch_Click()
   K = 3
   frmPurchaseRpt.Show

End Sub

Private Sub menuExit_Click()
  End
End Sub

Private Sub menuExit1_Click()
  End
End Sub

Private Sub menuFinish_Click()
  frmSales.Show
End Sub

Private Sub menuFinishSummary_Click()
   frmFinishItemStochSumary.Show
End Sub

Private Sub menuIssueConsume_Click()
  frmIssueConsumeItem.Show
End Sub

Private Sub menuIssueReturn_Click()
frmIssueReturn.Show
End Sub

Private Sub menuItem_Click()
  frmItemMaster.Show
End Sub

Private Sub menuPay_Click()
Pay_rec = "pay"
frmReceipt.Show
End Sub

Private Sub menuProcess_Click()
  frmIssue.Show
End Sub

Private Sub menuPurchase_Click()
   frmPurchase.Show
End Sub

Private Sub menuRaw_Click()
  'frmStockAson.Show
End Sub

Private Sub menuPurchaseReturn_Click()
frmPurchaseReturn.Show
End Sub

Private Sub menuRaw1_Click()
frmFinishPurchase.Show
End Sub

Private Sub menusaleinvoicereport_Click()
'frmsaleinvoicerpt.Show
End Sub

Private Sub menuSaleinvoicesummery_Click()
K = 5
frmPurchaseRpt.Show
End Sub

Private Sub menuStock_Click()
 K = 4
 'frmPurchaseRpt.Show
 frmRawStockSummary.Show
End Sub

Private Sub menuStockSummary_Click()
  K = 3
  'frmPurchaseRpt.Show
  frmConsumeItemSummary.Show
End Sub

Private Sub menuSummery_Click()
K = 8
frmPurchaseRpt.Caption = "SUMMERY"

frmPurchaseRpt.Show
End Sub

Private Sub menuSupplier_Click()
     Record = ""
     frmSupplier.Show
      
End Sub

Private Sub menuTax_Click()
   frmTaxMaster.Show
End Sub

Private Sub MENUTAXINVOICESUMMERY_Click()
K = 6
frmPurchaseRpt.Caption = "TAX INVOICE SUMMERY"
frmPurchaseRpt.Show
End Sub

Private Sub MenuUnit_Click()
   UnitMaster.Show
End Sub

Private Sub mnuBackup_Click()
frmbackup.Show
End Sub

Private Sub mnubill_Click()
Unload frmsaleinvoice
INVOICETYPE = "BILL"
frmsaleinvoice.Caption = "Bill"
frmsaleinvoice.LBLTAX.Visible = False
frmsaleinvoice.txttaxamount.Visible = False
frmsaleinvoice.txtTaxper.Visible = False
frmsaleinvoice.txtTotal.Visible = False
''''''''''''''''''''''''''''
frmsaleinvoice.Label15.Visible = True
frmsaleinvoice.txtRemarks.Visible = True

'frmsaleinvoice.LBLTAX.Visible = True
'frmsaleinvoice.txttaxper.Visible = True
'frmsaleinvoice.txttaxamount.Visible = True
frmsaleinvoice.Label16.Visible = True
frmsaleinvoice.txtSubTotal.Visible = True
'frmsaleinvoice.txtTotal.Visible = True
'frmsaleinvoice.Label20.Visible = True
frmsaleinvoice.txtLoadingUnloading.Visible = True
frmsaleinvoice.Label21.Visible = True
frmsaleinvoice.txtOtherCharges.Visible = True
frmsaleinvoice.Label18.Visible = True
frmsaleinvoice.txtTotalAmount.Visible = True



frmsaleinvoice.Label23.Visible = False
frmsaleinvoice.Label24.Visible = False
frmsaleinvoice.Label25.Visible = False
frmsaleinvoice.Label26.Visible = False
frmsaleinvoice.Label27.Visible = False
frmsaleinvoice.Label4.Visible = False
frmsaleinvoice.Label17.Visible = False
frmsaleinvoice.txtSaleStockTransfer.Visible = False
frmsaleinvoice.txtTimeOfRemovalOfGoods.Visible = False
frmsaleinvoice.txtTaxSaleInvoiceNo.Visible = False
frmsaleinvoice.txtModeOfTransportation.Visible = False
frmsaleinvoice.txtTransportationDate.Visible = False
frmsaleinvoice.txtRoadPermitNo.Visible = False
frmsaleinvoice.txtRoadPermitDate.Visible = False



frmsaleinvoice.chkLoading.Visible = True
frmsaleinvoice.chkUnloading.Visible = True
frmsaleinvoice.chkCartage.Visible = True



frmsaleinvoice.Label29.Visible = True
frmsaleinvoice.tot1.Visible = True
frmsaleinvoice.Label28.Visible = True
frmsaleinvoice.txtRoundoff.Visible = True
''''''''''''''''''''''''''''


frmsaleinvoice.Show
End Sub

Private Sub mnuBillent_Click()
frmIssue.Show
End Sub

Private Sub mnuBinder_Click()
frmSupplier.Show
End Sub

Private Sub mnuBookDel_Click()
ss = "D"
frmBookRecIssue.Show
End Sub

Private Sub mnuBookRec_Click()
ss = "R"
frmITC.Show
End Sub

Private Sub mnuBookRec1_Click()
ss = "R"
frmBookRecIssue.Show
End Sub

Private Sub mnuCashMemo_Click()
Unload frmsaleinvoice
INVOICETYPE = "CASH MEMO"
frmsaleinvoice.Caption = "Cash Memo"
frmsaleinvoice.LBLTAX.Visible = False
frmsaleinvoice.txttaxamount.Visible = False
frmsaleinvoice.txtTaxper.Visible = False
frmsaleinvoice.txtTotal.Visible = False
''''''''''''''''''''''''''''
frmsaleinvoice.Label15.Visible = True
frmsaleinvoice.txtRemarks.Visible = True

'frmsaleinvoice.LBLTAX.Visible = True
'frmsaleinvoice.txttaxper.Visible = True
'frmsaleinvoice.txttaxamount.Visible = True
frmsaleinvoice.Label16.Visible = True
frmsaleinvoice.txtSubTotal.Visible = True
'frmsaleinvoice.txtTotal.Visible = True
'frmsaleinvoice.Label20.Visible = True
frmsaleinvoice.txtLoadingUnloading.Visible = True
frmsaleinvoice.Label21.Visible = True
frmsaleinvoice.txtOtherCharges.Visible = True
frmsaleinvoice.Label18.Visible = True
frmsaleinvoice.txtTotalAmount.Visible = True


frmsaleinvoice.Label23.Visible = False
frmsaleinvoice.Label24.Visible = False
frmsaleinvoice.Label25.Visible = False
frmsaleinvoice.Label26.Visible = False
frmsaleinvoice.Label27.Visible = False
frmsaleinvoice.Label4.Visible = False
frmsaleinvoice.Label17.Visible = False
frmsaleinvoice.txtSaleStockTransfer.Visible = False
frmsaleinvoice.txtTimeOfRemovalOfGoods.Visible = False
frmsaleinvoice.txtTaxSaleInvoiceNo.Visible = False
frmsaleinvoice.txtModeOfTransportation.Visible = False
frmsaleinvoice.txtTransportationDate.Visible = False
frmsaleinvoice.txtRoadPermitNo.Visible = False
frmsaleinvoice.txtRoadPermitDate.Visible = False



frmsaleinvoice.chkLoading.Visible = True
frmsaleinvoice.chkUnloading.Visible = True
frmsaleinvoice.chkCartage.Visible = True




frmsaleinvoice.Label29.Visible = True
frmsaleinvoice.tot1.Visible = True
frmsaleinvoice.Label28.Visible = True
frmsaleinvoice.txtRoundoff.Visible = True
''''''''''''''''''''''''''''


frmsaleinvoice.Show
End Sub

Private Sub mnuChallan_Click()
frmChallan.Show
End Sub

Private Sub mnuChallanTransferInvoice_Click()
Unload frmsaleinvoice
INVOICETYPE = "CHALLAN/TRANSFER INVOICE"
frmsaleinvoice.Caption = "Challan/Transfer Invoice"
frmsaleinvoice.LBLTAX.Visible = False
frmsaleinvoice.txttaxamount.Visible = False
frmsaleinvoice.txtTaxper.Visible = False
'''''''''''''''''''''''''''''''''
'frmsaleinvoice.Label15.Visible = False
'frmsaleinvoice.txtRemarks.Visible = False

'frmsaleinvoice.LBLTAX.Visible = False
'frmsaleinvoice.txttaxper.Visible = False
'frmsaleinvoice.txttaxamount.Visible = False
frmsaleinvoice.Label16.Visible = False
frmsaleinvoice.txtSubTotal.Visible = False
frmsaleinvoice.txtTotal.Visible = False
'frmsaleinvoice.Label20.Visible = False
frmsaleinvoice.txtLoadingUnloading.Visible = False
frmsaleinvoice.Label21.Visible = False
frmsaleinvoice.txtOtherCharges.Visible = False
frmsaleinvoice.Label18.Visible = False
frmsaleinvoice.txtTotalAmount.Visible = False

frmsaleinvoice.Label23.Visible = True
frmsaleinvoice.Label24.Visible = True
frmsaleinvoice.Label25.Visible = True
frmsaleinvoice.Label26.Visible = True
frmsaleinvoice.Label27.Visible = True
frmsaleinvoice.Label4.Visible = True
frmsaleinvoice.Label17.Visible = True
frmsaleinvoice.txtSaleStockTransfer.Visible = True
frmsaleinvoice.txtTimeOfRemovalOfGoods.Visible = True
frmsaleinvoice.txtTaxSaleInvoiceNo.Visible = True
frmsaleinvoice.txtModeOfTransportation.Visible = True
frmsaleinvoice.txtTransportationDate.Visible = True
frmsaleinvoice.txtRoadPermitNo.Visible = True
frmsaleinvoice.txtRoadPermitDate.Visible = True



frmsaleinvoice.chkLoading.Visible = False
frmsaleinvoice.chkUnloading.Visible = False
frmsaleinvoice.chkCartage.Visible = False




frmsaleinvoice.Label29.Visible = False
frmsaleinvoice.tot1.Visible = False
frmsaleinvoice.Label28.Visible = False
frmsaleinvoice.txtRoundoff.Visible = False
 '''''''''''''''''''''''''''''''''''''''''''''''
frmsaleinvoice.Show
End Sub

Private Sub mnuComp_Click()
setup.Show
End Sub

Private Sub mnuCompany_Click()
  setup.Show
End Sub

Private Sub mnuConsumed_Click()
ss = "D"
frmITC.Show
End Sub

Private Sub mnuCust1_Click()
Pay_rec = "rec"
frmSaleInvoiceRpt.Show

End Sub

Private Sub mnuCustomer_Click()
Pay_rec = "rec"
frmSaleInvoiceRpt.Show
End Sub

Private Sub mnuCustomerL_Click()
 frmledger.Show 1
End Sub

Private Sub mnuJobs_Click()
 frmFolding.Show
End Sub

Private Sub mnuMaster_Click()
frmGp.Show

End Sub

Private Sub mnuNewBook_Click()
ss1 = "new"
frmFinishItemStochSumary.Show
End Sub

Private Sub mnuOldBookLedger_Click()
ss1 = "itc"
frmFinishItemStochSumary.Show
End Sub

Private Sub mnuPendingBill_Click()
frmPendingBill.Show
End Sub

Private Sub mnuPerformaInvoice_Click()
'Unload frmsaleinvoice
'INVOICETYPE = "PERFORMAINVOICE"
'frmsaleinvoice.LBLTAX.Visible = True
'frmsaleinvoice.txttaxper.Visible = True
'frmsaleinvoice.txttaxamount.Visible = True
'frmsaleinvoice.txtTotal.Visible = True
'frmsaleinvoice.Caption = "Performa Invoice"
'frmsaleinvoice.LBLTAX.Caption = "CST % "
'''''''''''''''''''''''''''''
'frmsaleinvoice.Label15.Visible = True
'frmsaleinvoice.txtRemarks.Visible = True
'
''frmsaleinvoice.LBLTAX.Visible = True
''frmsaleinvoice.txttaxper.Visible = True
''frmsaleinvoice.txttaxamount.Visible = True
'frmsaleinvoice.Label16.Visible = True
'frmsaleinvoice.txtSubTotal.Visible = True
'frmsaleinvoice.txtTotal.Visible = True
''frmsaleinvoice.Label20.Visible = True
'frmsaleinvoice.txtLoadingUnloading.Visible = True
'frmsaleinvoice.Label21.Visible = True
'frmsaleinvoice.txtOtherCharges.Visible = True
'frmsaleinvoice.Label18.Visible = True
'frmsaleinvoice.txtTotalAmount.Visible = True
'
'
'frmsaleinvoice.Label23.Visible = False
'frmsaleinvoice.Label24.Visible = False
'frmsaleinvoice.Label25.Visible = False
'frmsaleinvoice.Label26.Visible = False
'frmsaleinvoice.Label27.Visible = False
'frmsaleinvoice.Label4.Visible = False
'frmsaleinvoice.Label17.Visible = False
'frmsaleinvoice.txtSaleStockTransfer.Visible = False
'frmsaleinvoice.txtTimeOfRemovalOfGoods.Visible = False
'frmsaleinvoice.txtTaxSaleInvoiceNo.Visible = False
'frmsaleinvoice.txtModeOfTransportation.Visible = False
'frmsaleinvoice.txtTransportationDate.Visible = False
'frmsaleinvoice.txtRoadPermitNo.Visible = False
'frmsaleinvoice.txtRoadPermitDate.Visible = False
'
'
'
'frmsaleinvoice.chkLoading.Visible = True
'frmsaleinvoice.chkUnloading.Visible = True
'frmsaleinvoice.chkCartage.Visible = True
'
'frmsaleinvoice.Label29.Visible = True
'frmsaleinvoice.tot1.Visible = True
'frmsaleinvoice.Label28.Visible = True
'frmsaleinvoice.txtRoundoff.Visible = True
'
'''''''''''''''''''''''''''''
'frmsaleinvoice.Show

Unload frmsaleinvoice
INVOICETYPE = "PERFORMAINVOICE"
frmsaleinvoice.LBLTAX.Visible = True
frmsaleinvoice.txtTaxper.Visible = True
frmsaleinvoice.txttaxamount.Visible = True
frmsaleinvoice.txtTotal.Visible = True
frmsaleinvoice.Caption = "Performa Invoice"
frmsaleinvoice.LBLTAX.Caption = "VAT % "
''''''''''''''''''''''''''''
frmsaleinvoice.Label15.Visible = True
frmsaleinvoice.txtRemarks.Visible = True
frmsaleinvoice.Label16.Visible = True
frmsaleinvoice.txtSubTotal.Visible = True
frmsaleinvoice.txtTotal.Visible = True
frmsaleinvoice.txtLoadingUnloading.Visible = True
frmsaleinvoice.Label21.Visible = True
frmsaleinvoice.txtOtherCharges.Visible = True
frmsaleinvoice.Label18.Visible = True
frmsaleinvoice.txtTotalAmount.Visible = True


frmsaleinvoice.Label23.Visible = False
frmsaleinvoice.Label24.Visible = False
frmsaleinvoice.Label25.Visible = False
frmsaleinvoice.Label26.Visible = False
frmsaleinvoice.Label27.Visible = False
frmsaleinvoice.Label4.Visible = False
frmsaleinvoice.Label17.Visible = False
frmsaleinvoice.txtSaleStockTransfer.Visible = False
frmsaleinvoice.txtTimeOfRemovalOfGoods.Visible = False
frmsaleinvoice.txtTaxSaleInvoiceNo.Visible = False
frmsaleinvoice.txtModeOfTransportation.Visible = False
frmsaleinvoice.txtTransportationDate.Visible = False
frmsaleinvoice.txtRoadPermitNo.Visible = False
frmsaleinvoice.txtRoadPermitDate.Visible = False



frmsaleinvoice.chkLoading.Visible = True
frmsaleinvoice.chkUnloading.Visible = True
frmsaleinvoice.chkCartage.Visible = True



frmsaleinvoice.Label29.Visible = True
frmsaleinvoice.tot1.Visible = True
frmsaleinvoice.Label28.Visible = True
frmsaleinvoice.txtRoundoff.Visible = True


''''''''''''''''''''''''''''


frmsaleinvoice.Show
End Sub

Private Sub mnuPTitle_Click()
frmTitleST.Show
End Sub

Private Sub mnurec_Click()
Pay_rec = "rec"
frmReceipt.Show
End Sub

Private Sub mnuRec1_Click()
Pay_rec = "rec"
frmpayrec.Show
End Sub

Private Sub mnuRes_Click()
frmRestoredata.Show
End Sub

Private Sub mnureset_Click()
frmreset.Show 1
End Sub

Private Sub mnusalesinvoice_Click()
Unload frmsaleinvoice
INVOICETYPE = "SALEINVOICE"
frmsaleinvoice.LBLTAX.Visible = True
frmsaleinvoice.txtTaxper.Visible = True
frmsaleinvoice.txttaxamount.Visible = True
frmsaleinvoice.txtTotal.Visible = True
frmsaleinvoice.Caption = "Sale Invoice"
frmsaleinvoice.LBLTAX.Caption = "CST % "
''''''''''''''''''''''''''''
frmsaleinvoice.Label15.Visible = True
frmsaleinvoice.txtRemarks.Visible = True

'frmsaleinvoice.LBLTAX.Visible = True
'frmsaleinvoice.txttaxper.Visible = True
'frmsaleinvoice.txttaxamount.Visible = True
frmsaleinvoice.Label16.Visible = True
frmsaleinvoice.txtSubTotal.Visible = True
frmsaleinvoice.txtTotal.Visible = True
'frmsaleinvoice.Label20.Visible = True
frmsaleinvoice.txtLoadingUnloading.Visible = True
frmsaleinvoice.Label21.Visible = True
frmsaleinvoice.txtOtherCharges.Visible = True
frmsaleinvoice.Label18.Visible = True
frmsaleinvoice.txtTotalAmount.Visible = True


frmsaleinvoice.Label23.Visible = False
frmsaleinvoice.Label24.Visible = False
frmsaleinvoice.Label25.Visible = False
frmsaleinvoice.Label26.Visible = False
frmsaleinvoice.Label27.Visible = False
frmsaleinvoice.Label4.Visible = False
frmsaleinvoice.Label17.Visible = False
frmsaleinvoice.txtSaleStockTransfer.Visible = False
frmsaleinvoice.txtTimeOfRemovalOfGoods.Visible = False
frmsaleinvoice.txtTaxSaleInvoiceNo.Visible = False
frmsaleinvoice.txtModeOfTransportation.Visible = False
frmsaleinvoice.txtTransportationDate.Visible = False
frmsaleinvoice.txtRoadPermitNo.Visible = False
frmsaleinvoice.txtRoadPermitDate.Visible = False



frmsaleinvoice.chkLoading.Visible = True
frmsaleinvoice.chkUnloading.Visible = True
frmsaleinvoice.chkCartage.Visible = True

frmsaleinvoice.Label29.Visible = True
frmsaleinvoice.tot1.Visible = True
frmsaleinvoice.Label28.Visible = True
frmsaleinvoice.txtRoundoff.Visible = True

''''''''''''''''''''''''''''
frmsaleinvoice.Show
End Sub

Private Sub mnuStudent_Click()
frmStudentDetail.Show
End Sub

Private Sub mnuSupp_Click()
Pay_rec = "pay"
frmSaleInvoiceRpt.Show

End Sub

Private Sub mnutaxinvoice_Click()
Unload frmsaleinvoice
INVOICETYPE = "TAXINVOICE"
frmsaleinvoice.LBLTAX.Visible = True
frmsaleinvoice.txtTaxper.Visible = True
frmsaleinvoice.txttaxamount.Visible = True
frmsaleinvoice.txtTotal.Visible = True
frmsaleinvoice.Caption = "Tax Invoice"
frmsaleinvoice.LBLTAX.Caption = "VAT % "
''''''''''''''''''''''''''''
frmsaleinvoice.Label15.Visible = True
frmsaleinvoice.txtRemarks.Visible = True

'frmsaleinvoice.LBLTAX.Visible = True
'frmsaleinvoice.txttaxper.Visible = True
'frmsaleinvoice.txttaxamount.Visible = True
frmsaleinvoice.Label16.Visible = True
frmsaleinvoice.txtSubTotal.Visible = True
frmsaleinvoice.txtTotal.Visible = True
'frmsaleinvoice.Label20.Visible = True
frmsaleinvoice.txtLoadingUnloading.Visible = True
frmsaleinvoice.Label21.Visible = True
frmsaleinvoice.txtOtherCharges.Visible = True
frmsaleinvoice.Label18.Visible = True
frmsaleinvoice.txtTotalAmount.Visible = True






frmsaleinvoice.Label23.Visible = False
frmsaleinvoice.Label24.Visible = False
frmsaleinvoice.Label25.Visible = False
frmsaleinvoice.Label26.Visible = False
frmsaleinvoice.Label27.Visible = False
frmsaleinvoice.Label4.Visible = False
frmsaleinvoice.Label17.Visible = False
frmsaleinvoice.txtSaleStockTransfer.Visible = False
frmsaleinvoice.txtTimeOfRemovalOfGoods.Visible = False
frmsaleinvoice.txtTaxSaleInvoiceNo.Visible = False
frmsaleinvoice.txtModeOfTransportation.Visible = False
frmsaleinvoice.txtTransportationDate.Visible = False
frmsaleinvoice.txtRoadPermitNo.Visible = False
frmsaleinvoice.txtRoadPermitDate.Visible = False



frmsaleinvoice.chkLoading.Visible = True
frmsaleinvoice.chkUnloading.Visible = True
frmsaleinvoice.chkCartage.Visible = True



frmsaleinvoice.Label29.Visible = True
frmsaleinvoice.tot1.Visible = True
frmsaleinvoice.Label28.Visible = True
frmsaleinvoice.txtRoundoff.Visible = True


''''''''''''''''''''''''''''


frmsaleinvoice.Show
End Sub

Private Sub mnuTConsumed_Click()
ss = "D"
frmTitle.Show

End Sub

Private Sub mnuTitle1_Click()
ss = "R"
frmTitle.Show

End Sub

Private Sub mnuTLedger1_Click()
ss1 = "title"


frmFinishItemStochSumary.Show

End Sub
Private Sub mnuVAT_Click()
frmVAT.Show
End Sub

Private Sub mnuVat1_Click()
frmVatRegister.Show
End Sub

Private Sub titleDelivered_Click()
frmTitleDelivered.Show
End Sub
