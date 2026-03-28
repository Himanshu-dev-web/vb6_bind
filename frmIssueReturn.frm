VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmIssueReturn 
   Caption         =   "Issue Return"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Search Option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5250
      Left            =   10080
      TabIndex        =   25
      Top             =   1905
      Width           =   2700
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  &Search  "
         Height          =   345
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   810
         Width           =   1005
      End
      Begin VB.ListBox ListHeatingNo 
         Appearance      =   0  'Flat
         Height          =   3930
         Left            =   165
         TabIndex        =   27
         Top             =   1215
         Width           =   2445
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print Sale Tax"
         Height          =   345
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   450
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   150
         TabIndex        =   29
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   38679
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   150
         TabIndex        =   30
         Top             =   825
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   38679
      End
      Begin VB.Label Label7 
         Caption         =   "To "
         Height          =   300
         Left            =   225
         TabIndex        =   31
         Top             =   600
         Width           =   645
      End
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6330
      TabIndex        =   24
      Top             =   5100
      Width           =   1110
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1725
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1230
      Width           =   5820
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   750
      Left            =   120
      TabIndex        =   19
      Top             =   6720
      Width           =   4785
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   570
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   570
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   570
         Left            =   1905
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   570
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   570
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   915
      End
   End
   Begin VB.TextBox txtHeating 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   3915
   End
   Begin VB.ComboBox cboDeptt 
      Height          =   315
      Left            =   5820
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Top             =   840
      Width           =   4005
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8340
      TabIndex        =   17
      Top             =   5100
      Width           =   1605
   End
   Begin VB.TextBox tot1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   8340
      TabIndex        =   16
      Top             =   6525
      Width           =   1605
   End
   Begin VB.TextBox txtRoundoff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8340
      TabIndex        =   15
      Top             =   6810
      Width           =   1605
   End
   Begin VB.CheckBox chkCartage 
      BackColor       =   &H80000013&
      Caption         =   "Cartage"
      Height          =   255
      Left            =   7275
      TabIndex        =   14
      Top             =   5955
      Width           =   1035
   End
   Begin VB.CheckBox chkUnloading 
      BackColor       =   &H80000013&
      Caption         =   "Unloading"
      Height          =   255
      Left            =   6060
      TabIndex        =   13
      Top             =   5955
      Width           =   1215
   End
   Begin VB.CheckBox chkLoading 
      BackColor       =   &H80000013&
      Caption         =   "Loading"
      Height          =   255
      Left            =   5085
      TabIndex        =   12
      Top             =   5955
      Width           =   975
   End
   Begin VB.TextBox txttaxper 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7800
      TabIndex        =   11
      Top             =   5385
      Width           =   525
   End
   Begin VB.TextBox txttaxamount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8340
      TabIndex        =   10
      Top             =   5385
      Width           =   1605
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   8340
      TabIndex        =   9
      Top             =   7095
      Width           =   1605
   End
   Begin VB.TextBox txtLoadingUnloading 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   8340
      TabIndex        =   8
      Top             =   5955
      Width           =   1605
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   8340
      TabIndex        =   7
      Top             =   5670
      Width           =   1605
   End
   Begin VB.TextBox txtOtherCharges 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   8340
      TabIndex        =   6
      Top             =   6240
      Width           =   1605
   End
   Begin Crystal.CrystalReport cr 
      Left            =   8100
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3210
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   9840
      _cx             =   17357
      _cy             =   5662
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13888387
      ForeColorSel    =   16711680
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIssueReturn.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      Editable        =   2
      ShowComboButton =   1
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
   Begin MSComCtl2.DTPicker HeatingDate 
      Height          =   300
      Left            =   1725
      TabIndex        =   1
      Top             =   495
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   38679
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Height          =   60
      Left            =   60
      TabIndex        =   42
      Top             =   1665
      Width           =   12675
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Total :"
      Height          =   300
      Index           =   6
      Left            =   5730
      TabIndex        =   41
      Top             =   5100
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks "
      Height          =   300
      Index           =   4
      Left            =   285
      TabIndex        =   40
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date "
      Height          =   300
      Index           =   1
      Left            =   255
      TabIndex        =   39
      Top             =   525
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue No"
      Height          =   270
      Index           =   0
      Left            =   270
      TabIndex        =   38
      Top             =   195
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   300
      Index           =   2
      Left            =   255
      TabIndex        =   37
      Top             =   885
      Width           =   1215
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   7080
      TabIndex        =   36
      Top             =   6525
      Width           =   1140
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Roundoff"
      Height          =   195
      Left            =   7500
      TabIndex        =   35
      Top             =   6810
      Width           =   765
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   7020
      TabIndex        =   34
      Top             =   6240
      Width           =   1260
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      Height          =   195
      Left            =   7200
      TabIndex        =   33
      Top             =   7095
      Width           =   1005
   End
   Begin VB.Label LBLTAX 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "CST %"
      Height          =   195
      Left            =   6900
      TabIndex        =   32
      Top             =   5460
      Width           =   840
   End
End
Attribute VB_Name = "frmIssueReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemgp As String
Dim i As Integer

Private Sub cmdMain_Click()
Unload Me
End Sub
Sub AddItemInGrid()
    Dim m As String
    m = ""
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select ItemName from ItemMaster order by Itemname", con, adOpenDynamic, adLockOptimistic, adCmdText
    While rs.EOF = False
        If rs.EOF = False Then
            If m = "" Then
               m = rs!itemname
            Else
               m = m & "|" & rs!itemname
            End If
            vs.ColComboList(0) = m
       End If
        rs.MoveNext
    Wend

End Sub
Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQty.SetFocus
End Sub

Private Sub cdmModify_Click()

End Sub
Private Sub cbodeptt_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub cboName_Click()

If rs.State = 1 Then rs.Close
rs.Open "select Loading,Unloading,Cartage,TaxHead,TaxRate,OtherCharges from Customer where Name='" & cboName.Text & "'", con
If rs.EOF = False Then
  Me.LBLTAX.Caption = rs!TaxHead & ""
  Me.txtTaxper.Text = rs!TaxRate & ""
  Me.txtOtherCharges.Text = rs!OtherCharges & ""
  chkLoading.Value = IIf(rs!lOADING = True, 1, 0)
  chkUnloading.Value = IIf(rs!unloading = True, 1, 0)
  chkCartage.Value = IIf(rs!cartage = True, 1, 0)
   
End If

End Sub

Private Sub cboName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub chkCartage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtLoadingUnloading.SetFocus
End If
End Sub

Private Sub chkLoading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
chkUnloading.SetFocus
End If
End Sub

Private Sub chkUnloading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
chkCartage.SetFocus
End If
End Sub

Private Sub cmdadd_Click()
 If rs.State = 1 Then rs.Close
 rs.Open "select distinct(IssueNo) from IssueConsumeItem where IssueDate >=datevalue('" & FromDate.Value & "') and IssueDate <=datevalue('" & ToDate.Value & "') and  ReturnQty =true order by IssueNo", con
 ListHeatingNo.Clear
 If rs.EOF = False Then
    While rs.EOF = False
       ListHeatingNo.AddItem rs(0)
       rs.MoveNext
    Wend
 End If
End Sub

Private Sub cmdDelete_Click()
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
      DeleteStock
      DeleteRecord txtHeating.Text, "IssueNo", "IssueConsumeItem"
      Call cmdRef_Click
   End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()
    If cboName.Text = "" Then
    MsgBox "Plz. Select Customer Name !!", vbInformation
    Exit Sub
    End If

   If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
      DeleteStock
      DeleteRecord txtHeating.Text, "IssueNo", "IssueConsumeItem"
      SaveData
      Call cmdRef_Click
   End If
End Sub

Private Sub CmdPrint_Click()
   cr.Reset
   cr.ReportFileName = App.Path & "\SaleTaxReg.rpt"
   cr.ReplaceSelectionFormula "{Finishurchase.IssueDate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.IssueDate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & ToDate.Value & "'"
   
   If rs.State = 1 Then rs.Close
   rs.Open "select yarfrom,yarto from setup1", con
   If rs.EOF = False Then
   cr.Formulas(2) = "yrs='" & rs(0) & " To " & rs(1) & "'"
   End If
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1

End Sub

Private Sub cmdRef_Click()
      txtHeating.Text = ""
      txtRemarks.Text = ""
      
      txtTotal1.Text = 0
      
      vs.Clear
      SetWidth
      txtHeating.SetFocus
      cmdDelete.Enabled = False
      cmdModify.Enabled = False
      cmdSave.Enabled = True
      'txtHeating.Text = MaxSNo("IssueConsumeItem", "IssueNo")
      HeatingDate.Value = Date
      
      
      Dim O As Object
      
      For Each O In Me
      If TypeOf O Is TextBox Then
      O.Text = ""
      End If
      
      Next
      
      chkCartage.Value = 0
      chkLoading.Value = 0
      chkUnloading.Value = 0
      ' txtHeating.Text = MaxSNo("IssueConsumeItem", "IssueNo")

End Sub


Private Sub Command4_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
    
    If cboName.Text = "" Then
    MsgBox "Plz. Select Customer Name !!", vbInformation
    Exit Sub
    End If
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from IssueConsumeItem where issueNo='" & txtHeating.Text & "'", con
    If rs.EOF = True Then
       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
          SaveData
          cmdSave.Enabled = False
 '         addDeptt
    
       End If
    Else
        
       MsgBox "This Number Already Exist ...", vbCritical
       Exit Sub

    End If
End Sub
Sub addCustomer()
If rs.State = 1 Then rs.Close
rs.Open "SELECT distinct Name FROM customer order by Name", con
While rs.EOF = False
cboName.AddItem rs(0)
rs.MoveNext
Wend
End Sub
Sub SaveData()
         
        Dim str1 As String
        str1 = ""
        
        
        If txtHeating.Text = "" Then
           MsgBox "Please Verify Entries !!", vbCritical
           Exit Sub
        End If
    
       '-----------------------------------------
    
       Dim rs1 As New ADODB.Recordset
       If rs.State = 1 Then rs.Close
         rs.Open "select * from IssueConsumeItem", con, adOpenDynamic, adLockOptimistic
           For i = vs.FixedRows To vs.Rows - 1
             If vs.TextMatrix(i, 0) <> "" Then
                rs.AddNew
                rs.Fields("IssueNo").Value = txtHeating.Text
                rs.Fields("IssueDate").Value = HeatingDate.Value
                rs.Fields("Nerration").Value = vs.TextMatrix(i, 0)
                rs.Fields("Item").Value = vs.TextMatrix(i, 1)
                rs.Fields("NoOfPcs").Value = vs.TextMatrix(i, 2)
                rs.Fields("Unit").Value = vs.TextMatrix(i, 3)
                rs.Fields("Qty").Value = vs.TextMatrix(i, 4)
                rs.Fields("rate").Value = vs.TextMatrix(i, 5)
                rs.Fields("Remarks").Value = txtRemarks.Text
                rs.Fields("deptt").Value = cboDeptt.Text
                rs.Fields("customerName").Value = cboName.Text
                
                rs!lOADING = chkLoading.Value
                rs!unloading = chkUnloading.Value
                rs!cartage = chkCartage.Value
      
                rs.Fields("TaxPer").Value = Val(Me.txtTaxper.Text)
                rs.Fields("TaxAmount").Value = Val(Me.txttaxamount.Text)
                rs.Fields("Fright").Value = Val(Me.txtLoadingUnloading.Text)
                rs.Fields("OtherCharges").Value = Val(Me.txtOtherCharges.Text)
                rs.Fields("Rountoff").Value = Me.txtRoundoff.Text
                rs.Fields("Net").Value = Me.txtTotalAmount.Text
                rs.Fields("ReturnQty").Value = True
                rs.Update
                
                UpdatePurchase
              End If
              
              
              
           Next
           
           'con.Execute "update IssueConsumeItem set TempData='" & str1 & "' where issueNo=" & txtHeating.Text & ""
           str1 = ""
End Sub

Sub DeletePurchase()
Dim balance As Double
'----------------------------------------------
  balance = 0
  balance = Val(vs.TextMatrix(i, 2))
aa:
  
  If rs1.State = 1 Then rs1.Close
  rs1.Open "select BillNo,BillDate,qty,QtyRec from Finishurchase where ItemName='" & vs.TextMatrix(i, 0) & "' and ((Qty-qtyrec)<=0 or (issue)<>'0') order by BillDate desc ", con
  If rs1.EOF = False Then
  If Val(balance) >= (rs1.Fields("qtyrec").Value) Then
     balance = (Val(balance) - (rs1.Fields("qtyrec").Value))
     con.Execute "update Finishurchase set QtyRec=QtyRec-" & (rs1.Fields("qtyrec").Value) & "  where (ItemName='" & vs.TextMatrix(i, 0) & "' and  BillNo='" & rs1.Fields("billno").Value & "')"
     con.Execute "update Finishurchase set Issue=" & 0 & "  where (ItemName='" & vs.TextMatrix(i, 0) & "' and  BillNo='" & rs1.Fields("billno").Value & "') and QtyRec=0"
     If balance > 0 Then
        GoTo aa:
     End If
  ElseIf Val(balance) <= (rs1.Fields("qtyrec").Value) Then
     con.Execute "update Finishurchase set QtyRec=QtyRec-" & (balance) & " where (ItemName='" & vs.TextMatrix(i, 0) & "' and  BillNo='" & rs1.Fields("billno").Value & "')"
     con.Execute "update Finishurchase set Issue=" & 0 & "  where (ItemName='" & vs.TextMatrix(i, 0) & "' and  BillNo='" & rs1.Fields("billno").Value & "') and QtyRec=0"
  End If
  End If


'-----------------------------------------------
End Sub

Sub UpdatePurchase()
'----------------------------------------------
  Dim balance As Double
  balance = 0
  balance = Val(vs.TextMatrix(i, 2))
aa:
  
  If rs1.State = 1 Then rs1.Close
  rs1.Open "select BillNo,BillDate,qty,QtyRec from Finishurchase where ItemName='" & vs.TextMatrix(i, 0) & "' and (Qty-qtyrec)>0 order by BillDate", con
  If rs1.EOF = False Then
  If Val(balance) >= (rs1.Fields("qty").Value - rs1.Fields("qtyrec").Value) Then
     balance = (Val(balance) - (rs1.Fields("qty").Value - rs1.Fields("qtyrec").Value))
     con.Execute "update Finishurchase set QtyRec=QtyRec+" & (rs1.Fields("qty").Value - rs1.Fields("qtyrec").Value) & ",Issue=" & Me.txtHeating.Text & " where (ItemName='" & vs.TextMatrix(i, 0) & "' and  BillNo='" & rs1.Fields("billno").Value & "')"
     If balance > 0 Then
        GoTo aa:
     End If
  ElseIf Val(balance) <= (rs1.Fields("qty").Value - rs1.Fields("qtyrec").Value) Then
     con.Execute "update Finishurchase set QtyRec=QtyRec+" & (balance) & ",Issue='" & Me.txtHeating.Text & "' where (ItemName='" & vs.TextMatrix(i, 0) & "' and  BillNo='" & rs1.Fields("billno").Value & "')"
  End If
  End If
'-----------------------------------------------
End Sub
Sub SearchData()
    On Error Resume Next
    Dim n
    n = 0
    vs.Clear
    SetWidth
    If ListHeatingNo.Text = "" Then
    n = txtHeating.Text
    Else
    n = ListHeatingNo.Text
    End If
    
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from IssueConsumeItem where IssueNo='" & n & "' and ReturnQty = true", con
    If rs.EOF = True Then Exit Sub
    If rs.EOF = False Then
      txtHeating.Text = rs.Fields("IssueNo").Value
      HeatingDate.Value = rs.Fields("IssueDate").Value
      txtRemarks.Text = rs.Fields("Remarks").Value & ""
      cboDeptt.Text = rs.Fields("deptt").Value
      cboName.Text = rs.Fields("customerName").Value
      
      Me.txtTaxper.Text = rs.Fields("TaxPer").Value & ""
      Me.txttaxamount.Text = rs.Fields("TaxAmount").Value & ""
      Me.txtLoadingUnloading.Text = rs.Fields("Fright").Value & ""
      Me.txtOtherCharges.Text = rs.Fields("OtherCharges").Value & ""
      Me.txtRoundoff.Text = rs.Fields("Rountoff").Value & ""
      Me.txtTotalAmount.Text = rs.Fields("Net").Value & ""
      
      chkLoading.Value = IIf(rs!lOADING = True, 1, 0)
      chkUnloading.Value = IIf(rs!unloading = True, 1, 0)
      chkCartage.Value = IIf(rs!cartage = True, 1, 0)

    
      
      vs.Rows = rs.RecordCount + 10
      For i = vs.FixedRows To vs.Rows - 1
        If rs.EOF = False Then
          
          
         vs.TextMatrix(i, 0) = rs.Fields("Nerration").Value
         vs.TextMatrix(i, 1) = rs.Fields("Item").Value
         vs.TextMatrix(i, 2) = rs.Fields("NoOfPcs").Value
         vs.TextMatrix(i, 3) = rs.Fields("Unit").Value
         vs.TextMatrix(i, 4) = rs.Fields("Qty").Value
         vs.TextMatrix(i, 5) = rs.Fields("rate").Value
         vs.TextMatrix(i, 6) = (rs.Fields("rate").Value * rs.Fields("Qty").Value)
          
         rs.MoveNext
        End If
     Next
    End If
    Total
    cmdDelete.Enabled = True
    cmdModify.Enabled = True
    cmdSave.Enabled = False
End Sub
Sub IssueStock()
    
'    Dim rs_u As New ADODB.Recordset
'
'    For i = vs.FixedRows To vs.Rows - 1
'
'     If vs.TextMatrix(i, 0) <> "" Then
'        If rs_u.State = 1 Then rs_u.Close
'        rs_u.Open "select * from Stock where ItemName='" & vs.TextMatrix(i, 0) & "' and status='Consume'", con, adOpenDynamic, adLockOptimistic
'        If rs_u.EOF = False Then
'
'            rs_u!qty = rs_u!qty - vs.TextMatrix(i, 2)
'            rs_u.Update
'
'        End If
'    End If
'
'    Next
    
End Sub
Sub DeleteStock()
    Dim rs_u As New ADODB.Recordset
    For i = vs.FixedRows To vs.Rows - 1
    If vs.TextMatrix(i, 0) <> "" Then
        DeletePurchase
    End If
    Next
End Sub
Sub ItemGpSearch(Str As String)
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ItemGp from ItemMaster where ItemName='" & Str & "'", con
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).Value
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 113 Then
       
      tbl = "Customer"
      popuplist11 "Select name from Customer order by Name", con, , True, colmn
      popuplist1.Show

 End If
' If KeyCode = 13 Then
' SendKeys "{tab}"
' End If
End Sub

Private Sub Form_Load()
 txtTaxper.Text = ""
If rs.State = 1 Then rs.Close
rs.Open "select cstper from setup1", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
txtTaxper.Text = rs!cstper
End If

 SetWidth
 
 'AddItemInGrid
 SetWidth
 'txtHeating.Text = MaxSNo("IssueConsumeItem", "IssueNo")
 
 HeatingDate.Value = Date
 FromDate.Value = Date
 ToDate.Value = Date
 
 addCustomer
       
 pos Me
 

 '===============================================
 Dim m As String
m = ""
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select ItemGp from ItemMaster group by ItemGp", con
While rs.EOF = False
If rs.EOF = False Then
If m = "" Then
m = rs!itemgp
Else
m = m & "|" & rs!itemgp
End If
vs.ColComboList(0) = m
End If
rs.MoveNext
Wend
 
End Sub

Sub addGridData()
Dim m As String
m = ""
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
If vs.TextMatrix(vs.RowSel, 0) <> "" Then
   rs.Open "select ItemName,Unit,ItemGp from ItemMaster where ItemGp='" & vs.TextMatrix(vs.RowSel, 0) & "' group by ItemName,Unit,ItemGp", con
   
   While rs.EOF = False
If rs.EOF = False Then
If m = "" Then
m = rs!itemname
Else
m = m & "|" & rs!itemname
End If
vs.ColComboList(1) = m
End If
rs.MoveNext
Wend

End If

End Sub
Sub SetWidth()
    vs.Cols = 7
    vs.FormatString = "<Group|Item Name|No of Pcs.|Unit|>Qty|Rate|>Amount"
    vs.ColWidth(0) = 1600
    vs.ColWidth(1) = 3000
    vs.ColWidth(2) = 1100
    vs.ColWidth(3) = 1000
    vs.ColWidth(4) = 900
    vs.ColWidth(5) = 900
    vs.ColWidth(6) = 1200
    
    
End Sub

Private Sub HeatingDate_KeyDown(KeyCode As Integer, Shift As Integer)
 '  If KeyCode = 13 Then cboDeptt.SetFocus
 If KeyCode = 13 Then
 SendKeys "{tab}"
 End If
End Sub
Private Sub ListHeatingNo_Click()
  SearchData
End Sub

Private Sub tot1_Change()
Total
End Sub

Private Sub tot1_GotFocus()
Total
End Sub

Private Sub txtHeating_GotFocus()

If PopUpValue1 <> "" Then
    txtHeating.Text = PopUpValue1
    SearchData
    PopUpValue1 = ""
End If

End Sub
Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then

popuplist2 "Select IssueNo,IssueDate from IssueConsumeItem where ReturnQty = true group by IssueNo,IssueDate", con

End If

If KeyCode = 13 Then
If txtHeating.Text <> "" Then
    SearchData
End If
End If

End Sub

Private Sub txtHeating_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then HeatingDate.SetFocus
End Sub
Private Sub txtParty_GotFocus()
  txtParty.Text = Record
End Sub

Private Sub txtParty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cboFinishItem.SetFocus
End Sub
Private Sub txtParty_LostFocus()
 txtParty.Text = Record
End Sub
Private Sub txtQty_GotFocus()
     txtQty.SelLength = 10
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtLoadingUnloading_Change()
Total
End Sub

Private Sub txtLoadingUnloading_GotFocus()
Total
End Sub

Private Sub txtLoadingUnloading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtOtherCharges.SetFocus
End If
End Sub

Private Sub txtOtherCharges_Change()
Total
End Sub

Private Sub txtOtherCharges_GotFocus()
Total
End Sub

Private Sub txtOtherCharges_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Me.cmdSave.Enabled = True Then
cmdSave.SetFocus
Else
cmdModify.SetFocus
End If
End If
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txttaxper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
chkLoading.SetFocus
End If
End Sub

Private Sub txttaxper_LostFocus()
Total
End Sub

Private Sub txtTotal_Change()
Total
End Sub

Private Sub txtTotal_GotFocus()
Total
End Sub

Private Sub txtTotalAmount_Change()
Total
End Sub

Private Sub txtTotalAmount_GotFocus()
Total
End Sub

Private Sub vs_GotFocus()
   If PopUpValue1 <> "" Then
      vs.TextMatrix(vs.RowSel, 0) = PopUpValue1
      vs.TextMatrix(vs.RowSel, 1) = PopUpValue2
       
      If rs.State = 1 Then rs.Close
      rs.Open "select Rate from Finishurchase where ItemName='" & PopUpValue1 & "' and (Qty-qtyrec)>0 order by BillDate", con
      If rs.EOF = False Then
         vs.TextMatrix(vs.RowSel, 3) = rs.Fields(0).Value
      End If
       PopUpValue1 = ""
      PopUpValue2 = ""
      PopUpValue3 = ""
   End If
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
    vs.RemoveItem (vs.RowSel)
    Total
  End If
    
  If KeyCode = 114 Then
     If rs1.State = 1 Then rs1.Close
     popuplist2 "select ItemName,Unit from Finishurchase group by ItemName,Unit", con
 End If
        
 
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
   If vs.Col = 0 Then
        
      addGridData
      SendKeys "{right}"
   ElseIf vs.Col = 1 Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where ItemName='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
        If rs.EOF = False Then
           vs.TextMatrix(vs.RowSel, 3) = rs.Fields("Unit").Value
            SendKeys "{right}"
           
        End If
    ElseIf vs.Col = 2 Then
        SendKeys "{right}"
        SendKeys "{right}"
    ElseIf vs.Col = 3 Then
        SendKeys "{right}"
    ElseIf vs.Col = 4 Then
        SendKeys "{right}"
   
    ElseIf vs.Col = 5 Then
           vs.TextMatrix(vs.RowSel, 6) = (Val(vs.TextMatrix(vs.RowSel, 4)) * Val(vs.TextMatrix(vs.RowSel, 5)))
           SendKeys "{home}"
           SendKeys "{down}"
           
    End If
    
    

End If

End Sub

Sub Total()
    
    On Error Resume Next
    txtTotal1.Text = 0
    txtAmt.Text = 0
    
    For i = 1 To vs.Rows - 1
        txtTotal1.Text = Format((CDbl(txtTotal1.Text) + CDbl(vs.TextMatrix(i, 4))), "0.000")
        txtAmt.Text = Format((CDbl(txtAmt.Text) + CDbl(vs.TextMatrix(i, 6))), "0.00")
    Next
    If Val(txtTaxper.Text) > 0 Then txttaxamount.Text = Format(((Val(txtAmt.Text) * Val(txtTaxper.Text) / 100)), "0.00")

    txtTotal.Text = Format(Val(txtAmt.Text) + Val(txttaxamount.Text), "0.00")
    tot1.Text = Format(Val(txtTotal.Text) + Val(txtOtherCharges.Text) + Val(txtLoadingUnloading.Text), "0.00")
    
    txtTotalAmount.Text = Round(Format(Val(tot1.Text), "0.00"))

    txtRoundoff.Text = Format((Val(txtTotalAmount.Text) - Val(tot1.Text)), "0.00")

End Sub
Sub Total1()
    On Error Resume Next
    txtTotal2.Text = 0
    For i = 1 To vs1.Rows - 1
        txtTotal2.Text = Format((CDbl(txtTotal2.Text) + CDbl(vs1.TextMatrix(i, 2))), "#,###.00")
    Next
End Sub
Sub Total2()
    
    On Error Resume Next
    
    txtTotal3.Text = 0
    For i = 1 To vs2.Rows - 1
        txtTotal3.Text = Format((CDbl(txtTotal3.Text) + CDbl(vs2.TextMatrix(i, 2))), "#,###.00")
    Next
End Sub
Private Sub vs_LeaveCell()
  Total
End Sub

Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
    vs1.RemoveItem (vs1.RowSel)
    Total1
  End If
End Sub

Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
   If vs1.Col = 0 Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", con
        If rs.EOF = False Then
           vs1.TextMatrix(vs1.RowSel, 1) = rs.Fields("Unit").Value
           SendKeys "{right}"
           SendKeys "{right}"
        End If
    
    ElseIf vs1.Col = 2 Then
           SendKeys "{home}"
           SendKeys "{down}"
    End If
    
    Total1

End If

End Sub
Private Sub vs1_LeaveCell()
   Total1
End Sub

Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 46 Then
    vs2.RemoveItem (vs2.RowSel)
    Total2
 End If
End Sub

Private Sub vs2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
   If vs2.Col = 0 Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", con
        If rs.EOF = False Then
           vs2.TextMatrix(vs2.RowSel, 1) = rs.Fields("Unit").Value
           SendKeys "{right}"
           SendKeys "{right}"
        End If
    
    ElseIf vs2.Col = 2 Then
           SendKeys "{home}"
           SendKeys "{down}"
    End If
    
    Total2

End If

End Sub
Private Sub vs2_LeaveCell()
   Total2
End Sub


