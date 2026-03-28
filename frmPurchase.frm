VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchase 
   Caption         =   "Raw Purchase"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPurchase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTaxper 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6825
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   6225
      Width           =   375
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7665
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0"
      Top             =   6840
      Width           =   1230
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7665
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "0"
      Top             =   6240
      Width           =   1245
   End
   Begin VB.TextBox txtFright 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7665
      TabIndex        =   6
      Text            =   "0"
      Top             =   6540
      Width           =   1230
   End
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
      Height          =   5625
      Left            =   9225
      TabIndex        =   21
      Top             =   1110
      Width           =   2535
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  &Search"
         Height          =   405
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ListBox ListBillNo 
         Appearance      =   0  'Flat
         Height          =   3930
         Left            =   135
         TabIndex        =   22
         Top             =   1530
         Width           =   2280
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   1035
         TabIndex        =   24
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   38679
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   1035
         TabIndex        =   25
         Top             =   705
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   38679
      End
      Begin VB.Label Label7 
         Caption         =   "To Date :"
         Height          =   300
         Left            =   105
         TabIndex        =   27
         Top             =   690
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "From Date :"
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   390
         Width           =   1320
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7665
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   5940
      Width           =   1245
   End
   Begin VB.TextBox txtMrvNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1395
      TabIndex        =   0
      Top             =   270
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   705
      Left            =   3090
      TabIndex        =   15
      Top             =   7185
      Width           =   5805
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   495
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   165
         Width           =   1170
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   165
         Width           =   1125
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   165
         Width           =   1170
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   495
         Left            =   4635
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   165
         Width           =   1110
      End
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   165
         Width           =   1140
      End
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1395
      TabIndex        =   1
      Top             =   645
      Width           =   4170
   End
   Begin MSComCtl2.DTPicker BillDate 
      Height          =   315
      Left            =   6510
      TabIndex        =   4
      Top             =   615
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38679
   End
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6510
      TabIndex        =   2
      Top             =   255
      Width           =   1800
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4875
      Left            =   90
      TabIndex        =   5
      Top             =   990
      Width           =   9045
      _cx             =   15954
      _cy             =   8599
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
      BackColorFixed  =   12640511
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
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPurchase.frx":0442
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
   Begin MSComCtl2.DTPicker RecDate 
      Height          =   315
      Left            =   4050
      TabIndex        =   3
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38679
   End
   Begin VB.Label Tax 
      Caption         =   "Tax :"
      Height          =   300
      Left            =   5445
      TabIndex        =   33
      Top             =   6225
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      Height          =   225
      Left            =   7230
      TabIndex        =   32
      Top             =   6255
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "Net Amount :"
      Height          =   300
      Index           =   8
      Left            =   5445
      TabIndex        =   31
      Top             =   6840
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Fright :"
      Height          =   300
      Index           =   6
      Left            =   5460
      TabIndex        =   28
      Top             =   6510
      Width           =   1260
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
      Left            =   1395
      TabIndex        =   20
      Top             =   30
      Width           =   2565
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   300
      Index           =   5
      Left            =   5445
      TabIndex        =   19
      Top             =   5925
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MRV No. "
      Height          =   300
      Index           =   4
      Left            =   90
      TabIndex        =   17
      Top             =   270
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "MRV Date "
      Height          =   300
      Index           =   3
      Left            =   3195
      TabIndex        =   16
      Top             =   285
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name "
      Height          =   300
      Index           =   2
      Left            =   75
      TabIndex        =   14
      Top             =   675
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date "
      Height          =   300
      Index           =   1
      Left            =   5610
      TabIndex        =   13
      Top             =   660
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No. "
      Height          =   300
      Index           =   0
      Left            =   5595
      TabIndex        =   12
      Top             =   285
      Width           =   870
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dd As Boolean
Dim itemgp As String
Dim i As Integer
Dim m As String
Dim billno As String
Private Sub BillDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub cmdadd_Click()
   If rs.State = 1 Then rs.Close
   rs.Open "select distinct(MRVNo) from Purchase where (RecDate>=datevalue('" & fromDate.Value & "') and RecDate<=datevalue('" & todate.Value & "'))", con, adOpenDynamic, adLockOptimistic
   ListBillNo.Clear
   If rs.EOF = False Then
      While rs.EOF = False
          ListBillNo.AddItem rs.Fields("MRVNo").Value
          rs.MoveNext
      Wend
   End If
   ListBillNo.SetFocus
End Sub

Private Sub cmdDel_Click()
   
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
      DeleteData
      Call cmdRef_Click
      Call cmdadd_Click
   End If
  
End Sub
Sub DeleteData()


    Dim cmd As New ADODB.Command
    
    DeleteStock
    
    Set cmd.ActiveConnection = con
    cmd.CommandText = "delete from Purchase where MRVNo='" & txtMrvNo.Text & "'"
    cmd.Execute
    
End Sub
Private Sub cmdMain_Click()
   Unload Me
End Sub
Sub SaveData()
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Purchase", con, adOpenDynamic, adLockOptimistic
    
    
    For i = 1 To vs.Rows - 1
    
      If vs.TextMatrix(i, 0) <> "" Then
        
        rs.AddNew
        rs.Fields("BillNo").Value = txtBillNo.Text
        rs.Fields("BillDate").Value = BillDate.Value
        rs.Fields("RecDate").Value = recdate.Value
        rs.Fields("PartyName").Value = txtParty.Text
        rs.Fields("MRVNo").Value = txtMrvNo.Text
        rs.Fields("ItemName").Value = vs.TextMatrix(i, 0)
        rs.Fields("Qty").Value = vs.TextMatrix(i, 1)
        rs.Fields("Unit").Value = vs.TextMatrix(i, 2)
        rs.Fields("Rate").Value = vs.TextMatrix(i, 3)
        rs.Fields("Amount").Value = vs.TextMatrix(i, 4)
        
        rs.Fields("TaxPer").Value = txttaxper.Text
        rs.Fields("TaxAmount").Value = txttax.Text
        rs.Fields("NetAmount").Value = txtNet.Text
        rs.Fields("Fright").Value = txtFright.Text
        rs.Update
        
        UpdateStock
       End If
   Next
    
    
End Sub

Private Sub cmdModify_Click()
     
     
    If txtBillNo.Text = "" Or txtParty.Text = "" Or txtTotal.Text = "" Then
       MsgBox "Please Verify Entries !!", vbCritical
       Exit Sub
    End If
    

     
     If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
        DeleteData
        SaveData
        Call cmdRef_Click
        Call cmdadd_Click
   End If
End Sub

Private Sub cmdRef_Click()
     txtBillNo.Text = ""
     cmdModify.Enabled = False
     cmdDel.Enabled = False
     txtParty.Text = ""
     txtMrvNo.Text = ""
     txtTotal.Text = ""
     txttaxper.Text = 0
     txttax.Text = 0
     txtNet.Text = 0
     txtFright.Text = 0
     txtTotal.Text = 0
     
     
     vs.Clear
     SetWidth
     vs.Rows = 20
     txtMrvNo.SetFocus
     
     
     
     cmdSave.Enabled = True
     cmdModify.Enabled = False
     cmdDel.Enabled = False
     Record = ""
     
     
   BillDate.Value = Date
   recdate.Value = Date
   'FromDate.Value = Date
   'ToDate.Value = Date

     
End Sub

Private Sub cmdSave_Click()
    
    
    If txtBillNo.Text = "" Or txtParty.Text = "" Or txtTotal.Text = "" Then
       MsgBox "Please Verify Entries !!", vbCritical
       Exit Sub
    End If
    

    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Purchase where MRVNo='" & txtMrvNo.Text & "'", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = True Then
      
      If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
         SaveData
         Call cmdRef_Click
         Call cmdadd_Click
      End If
    Else
        
      MsgBox "Dulicate MRV No !!", vbCritical
      
     
    End If
    
End Sub
Sub UpdateStock()
    
    Dim rs_u As New ADODB.Recordset
    If rs_u.State = 1 Then rs_u.Close
    rs_u.Open "select * from Stock where ItemName='" & vs.TextMatrix(i, 0) & "' and Status='Raw'", con, adOpenDynamic, adLockOptimistic
    If rs_u.EOF = True Then
        rs_u.AddNew
        rs_u!itemname = vs.TextMatrix(i, 0)
        ItemGpSearch vs.TextMatrix(i, 0)
        rs_u!itemgp = itemgp
        rs_u!qty = vs.TextMatrix(i, 1)
        rs_u!unit = vs.TextMatrix(i, 2)
        rs_u!Rate = vs.TextMatrix(i, 3)
        rs_u!Status = "Raw"
        rs_u.Update
     Else
        rs_u!qty = rs_u!qty + vs.TextMatrix(i, 1)
        rs_u.Update
    End If
    
End Sub

Sub ItemGpSearch(Str As String)
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ItemGp from ItemMaster where ItemName='" & Str & "'", con
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).Value
    End If
End Sub
Sub DeleteStock()
    
'    Dim rs_u As New ADODB.Recordset
'
'    For i = vs.FixedRows To vs.Rows - 1
'
'     If vs.TextMatrix(i, 0) <> "" Then
'        If rs_u.State = 1 Then rs_u.Close
'        rs_u.Open "select * from Stock where ItemName='" & vs.TextMatrix(i, 0) & "' and status='Raw'", con, adOpenDynamic, adLockOptimistic
'        If rs_u.EOF = False Then
'
'            If rs1.State = 1 Then rs1.Close
'            rs1.Open "select * from Purchase where ItemName='" & vs.TextMatrix(i, 0) & "' and BillNo='" & txtBillNo.Text & "'", con, adOpenDynamic, adLockOptimistic
'            rs_u!qty = rs_u!qty - rs1.Fields("qty").Value
'            rs_u.Update
'
'        End If
'    End If
'
'    Next
    
End Sub
Sub CaculateTaxAmount()
    If Val(txtTotal.Text) = 0 Then
       txtTotal.Text = 0
    End If
    
    If CDbl(txtFright.Text) = 0 Then
       txtFright.Text = 0
    End If
    
    
    txttax.Text = (CDbl(txtTotal.Text) * CDbl(txttaxper.Text)) / 100
    txttax.Text = Format(txttax.Text, "#,###.00")
    
    txtNet.Text = Format((CDbl(txttax.Text) + CDbl(txtTotal.Text)) + CDbl(txtFright.Text), "#,###.00")
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
       Unload Me
    ElseIf KeyCode = 113 Then
       tbl = "Supplier"
       popuplist11 "Select name from Supplier order by Name", con, , True, colmn
       popuplist1.Show
    ElseIf KeyCode = 13 Then
       'SendKeys "{tab}"
    End If
End Sub

Sub Search()
    
    On Error Resume Next
    
    vs.Clear
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Purchase where MRVNo='" & billno & "'", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        txtBillNo.Text = rs.Fields("BillNo").Value
        BillDate.Value = rs.Fields("BillDate").Value
        recdate.Value = rs.Fields("RecDate").Value
        txtParty.Text = rs.Fields("PartyName").Value
        txtMrvNo.Text = rs.Fields("MRVNo").Value
        txttaxper.Text = rs.Fields("TaxPer").Value
        txttax.Text = rs.Fields("TaxAmount").Value
        txtNet.Text = rs.Fields("NetAmount").Value
        
          
        vs.Rows = rs.RecordCount + 10
        For i = 1 To rs.RecordCount
          vs.TextMatrix(i, 0) = rs.Fields("ItemName").Value
          vs.TextMatrix(i, 1) = rs.Fields("Qty").Value
          vs.TextMatrix(i, 2) = rs.Fields("Unit").Value
          vs.TextMatrix(i, 3) = rs.Fields("Rate").Value
          vs.TextMatrix(i, 4) = rs.Fields("Amount").Value
          txtFright.Text = rs.Fields("Fright").Value
          rs.MoveNext
        Next
        
        
        
     Total
        
     
     
     
     
     
     cmdModify.Enabled = True
     cmdDel.Enabled = True
     cmdSave.Enabled = False
     
     End If
     
     

     
End Sub
Sub Total()
    
    txtTotal.Text = 0
    For i = 1 To vs.Rows - 1
       If vs.TextMatrix(i, 0) <> "" Then
        txtTotal.Text = Format((CDbl(txtTotal.Text) + CDbl(vs.TextMatrix(i, 4))), "#,###.00")
       End If
    Next
    
    CaculateTaxAmount
    
End Sub
Private Sub Form_Load()
    SetWidth
    
    colmn = 1
    popuplist11 "Select name from Supplier order by Name", con, , True, colmn
    tbl = "Supplier"
    
    m = ""
    If rs.State = 1 Then rs.Close
    rs.Open "select ItemName from ItemMaster where ItemGp='Raw Item' order by Itemname", con, adOpenDynamic, adLockOptimistic, adCmdText
    While rs.EOF = False
     
        If rs.EOF = False Then
            
            If m = "" Then
               m = rs!itemname
            Else
               m = m & "|" & rs!itemname
               'Rs.MoveNext
            End If
            vs.ColComboList(0) = m
            
       End If
        
        rs.MoveNext
     Wend
     
     
     
    '----------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from TaxHead", con
    If rs.EOF = False Then
       txttaxper.Text = rs(1)
       Tax.Caption = rs(0)
    End If
     
     
     
   BillDate.Value = Date
   recdate.Value = Date
   fromDate.Value = Date
   todate.Value = Date
     
     

End Sub
Sub SetWidth()
    
    vs.FormatString = "Item Name|>Qty|^Unit|>Rate|>Amount"
    vs.ColWidth(0) = 2500
    vs.ColWidth(1) = 1500
    vs.ColWidth(2) = 2200
    vs.ColWidth(3) = 1000
    vs.ColWidth(4) = 1000
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then todate.SetFocus
End Sub

Private Sub ListBillNo_Click()
   billno = ListBillNo
   Search
End Sub

Private Sub RecDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtParty.SetFocus
   End If
End Sub
Private Sub ToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       Call cmdadd_Click
    End If
End Sub

Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      BillDate.SetFocus
      billno = txtBillNo.Text
      Search
      SetWidth
   End If
End Sub
Private Sub txtFright_Change()
    
    'txtTax.Text = (CDbl(txtTotal.Text) * CDbl(txtTaxper.Text)) / 100
    'txtTax.Text = Format(txtTax.Text, "#,###.00")
    'txtNet.Text = Format((CDbl(txtTax.Text) + CDbl(txtTotal.Text)) + CDbl(txtFright.Text), "#,###.00")

End Sub

Private Sub txtFright_GotFocus()
   txtFright.SelLength = 15
End Sub
Private Sub txtFright_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      CaculateTaxAmount
      Call cmdSave_Click
   End If
End Sub
Private Sub txtFright_LostFocus()
   
   txtFright.Text = Format(txtFright.Text, "#,###.00")
End Sub

Private Sub txtMrvNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then recdate.SetFocus
End Sub

Private Sub txtParty_GotFocus()
   dd = False
   txtParty.Text = Record
   tbl = "Supplier"
End Sub
Private Sub txtParty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtBillNo.SetFocus
    End If
End Sub
Private Sub vs_GotFocus()
    dd = True
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
  
  On Error Resume Next
     
  If KeyCode = 13 Then
     
   If vs.Col = 0 Then
     If rs.State = 1 Then rs.Close
     rs.Open "select * from ItemMaster where ItemName='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
     If rs.EOF = False Then
        vs.TextMatrix(vs.RowSel, 2) = rs.Fields("Unit").Value
        vs.TextMatrix(vs.RowSel, 3) = rs.Fields("Rate").Value
        SendKeys "{right}"
     End If
     
     
   End If
   
   If vs.Col = 1 Then
     SendKeys "{home}"
     SendKeys "{down}"
     vs.TextMatrix(vs.RowSel, 4) = (CDbl(vs.TextMatrix(vs.RowSel, 3)) * CDbl(vs.TextMatrix(vs.RowSel, 1)))
     Total
   End If
     
   
   
  
  End If

End Sub
