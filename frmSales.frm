VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSales 
   Caption         =   "Dispatch Form "
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   Icon            =   "frmSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Dispatch From Opening"
      Height          =   240
      Left            =   1365
      TabIndex        =   3
      Top             =   915
      Width           =   2520
   End
   Begin VB.TextBox txtHeatingNo 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5670
      TabIndex        =   33
      Top             =   255
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search Finish Item"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5685
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   570
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      Top             =   270
      Width           =   1800
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1365
      TabIndex        =   2
      Top             =   585
      Width           =   4260
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   705
      Left            =   6030
      TabIndex        =   21
      Top             =   6960
      Width           =   5805
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         Width           =   1140
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   495
         Left            =   4635
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   150
         Width           =   1110
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   1170
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   150
         Width           =   1125
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   495
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   1170
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   5430
      Width           =   1245
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
      Height          =   2565
      Left            =   30
      TabIndex        =   13
      Top             =   5385
      Width           =   4875
      Begin VB.ListBox ListBillNo 
         Appearance      =   0  'Flat
         Height          =   2370
         Left            =   2685
         TabIndex        =   15
         Top             =   150
         Width           =   2130
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  &Search"
         Height          =   405
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   2145
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   975
         TabIndex        =   16
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   38679
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   975
         TabIndex        =   17
         Top             =   705
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   38679
      End
      Begin VB.Label Label6 
         Caption         =   "From Date :"
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   390
         Width           =   1320
      End
      Begin VB.Label Label7 
         Caption         =   "To Date :"
         Height          =   300
         Left            =   105
         TabIndex        =   18
         Top             =   690
         Width           =   1320
      End
   End
   Begin VB.TextBox txtFright 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10425
      TabIndex        =   5
      Text            =   "0"
      Top             =   6030
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   5730
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   6330
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtTaxper 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9555
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5715
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker RecDate 
      Height          =   315
      Left            =   4230
      TabIndex        =   1
      Top             =   255
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20774913
      CurrentDate     =   38679
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3900
      Left            =   15
      TabIndex        =   4
      Top             =   1455
      Width           =   11925
      _cx             =   21034
      _cy             =   6879
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   100
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSales.frx":0442
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "F3 For Search Item"
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
      Left            =   1365
      TabIndex        =   34
      Top             =   1200
      Width           =   2565
   End
   Begin VB.Label Label1 
      Caption         =   "Dispatch No"
      Height          =   300
      Index           =   0
      Left            =   45
      TabIndex        =   31
      Top             =   285
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Dispatch Date "
      Height          =   300
      Index           =   1
      Left            =   3180
      TabIndex        =   30
      Top             =   315
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Name "
      Height          =   270
      Index           =   2
      Left            =   60
      TabIndex        =   29
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   300
      Index           =   5
      Left            =   8205
      TabIndex        =   28
      Top             =   5415
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
      Left            =   1335
      TabIndex        =   27
      Top             =   15
      Width           =   2565
   End
   Begin VB.Label Label1 
      Caption         =   "Fright :"
      Height          =   300
      Index           =   6
      Left            =   8220
      TabIndex        =   26
      Top             =   6000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Net Amount :"
      Height          =   300
      Index           =   8
      Left            =   8205
      TabIndex        =   25
      Top             =   6330
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      Height          =   225
      Left            =   9990
      TabIndex        =   24
      Top             =   5745
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Tax 
      Caption         =   "Tax :"
      Height          =   300
      Left            =   8205
      TabIndex        =   23
      Top             =   5715
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dd As Boolean
Dim itemgp As String
Dim i As Integer
Dim m As String
Dim billno As String

Private Sub Check1_Click()
''' If Check1.Value = 1 Then
'''
'''    m = ""
'''    If rs.State = 1 Then rs.Close
'''    rs.Open "select ItemName from ItemMaster where ItemGp='Finish Item' order by Itemname", con, adOpenDynamic, adLockOptimistic, adCmdText
'''    While rs.EOF = False
'''        If rs.EOF = False Then
'''            If m = "" Then
'''               m = rs!itemname
'''            Else
'''               m = m & "|" & rs!itemname
'''            End If
'''       End If
'''        rs.MoveNext
'''     Wend
'''    vs.ColComboList(0) = m
'''
'''    '----------------
''' Else
'''    vs.ColComboList(0) = ""
''' End If
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     vs.SetFocus
  End If
End Sub

Private Sub cmdSearch_Click()
    frmSearchDispatche.Show
    
    
   
    
    frmSearchDispatche.ListView1.ListItems.Clear
    
    If rs.State = 1 Then rs.Close
    rs.Open "select HeatingNo,ItemName,Grade,Size,Balance,Unit from searchSales where balance>0", con
    For i = 1 To rs.RecordCount
      If rs.EOF = False Then
            
            Set litem = frmSearchDispatche.ListView1.ListItems.Add(, , rs.Fields(1).Value)
            
            
            
            litem.ListSubItems.Add , , rs.Fields(0).Value
            litem.ListSubItems.Add , , rs.Fields(2).Value
            litem.ListSubItems.Add , , rs.Fields(3).Value
            litem.ListSubItems.Add , , Format(rs.Fields(4).Value, ".000")
            litem.ListSubItems.Add , , rs.Fields(5).Value
      
      End If
      rs.MoveNext
    Next
    
End Sub

Private Sub Form_Activate()
txtBillNo.SetFocus
End Sub

Private Sub RecDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtParty.SetFocus
End If
End Sub

Private Sub cmdadd_Click()
   If rs.State = 1 Then rs.Close
   rs.Open "select distinct(BillNo) from Invoice where (RecDate>=datevalue('" & FromDate.Value & "') and RecDate<=datevalue('" & ToDate.Value & "'))", con, adOpenDynamic, adLockOptimistic
   ListBillNo.Clear
   If rs.EOF = False Then
      While rs.EOF = False
         If Not IsNull(rs.Fields("BillNo").Value) Then
          ListBillNo.AddItem rs.Fields("BillNo").Value
         End If
         rs.MoveNext
      Wend
   End If
End Sub

Private Sub cmdDel_Click()
   'cmdSave.Enabled = True
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
      
      deleteIssue
      
      DeleteData
      Call cmdRef_Click
      'Call cmdadd_Click
   End If
  
End Sub
Sub deleteIssue()


Dim rss As New ADODB.Recordset
Dim Search As New ADODB.Recordset

 
For i = 1 To vs.Rows - 1

 If vs.TextMatrix(i, 0) <> "" Then
 If vs.TextMatrix(i, 9) <> "" Then
    Set rss = New ADODB.Recordset
    rss.Open "select * from IssueRawMetrial where HeatingNo=" & vs.TextMatrix(i, 9) & " and ItemName='" & vs.TextMatrix(i, 0) & "'", con, adOpenDynamic, adLockOptimistic
    If rss.EOF = False Then
    If rss.Fields("issue").Value > 0 Then
    If Search.State = 1 Then Search.Close
    Search.Open "select * from Invoice where HeatNo='" & vs.TextMatrix(i, 9) & "' and ItemName='" & vs.TextMatrix(i, 0) & "' and BillNo=" & txtBillNo.Text & "", con, adOpenDynamic, adLockOptimistic
    If Search.EOF = False Then
       rss.Fields("Issue").Value = (CDbl(rss.Fields("Issue").Value) - CDbl(Search!qty))
       rss.Update
    End If
       
    End If
    End If
 End If
 End If

Next

End Sub
Sub DeleteData()
    Dim cmd As New ADODB.Command
    'DeleteStock
    Set cmd.ActiveConnection = con
    cmd.CommandText = "delete from Invoice where BillNo=" & txtBillNo.Text & ""
    cmd.Execute
End Sub
Private Sub cmdMain_Click()
   Unload Me
End Sub
Sub SaveData()
    
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Invoice", con, adOpenDynamic, adLockOptimistic
    
    
    For i = 1 To vs.Rows - 1
    
      If vs.TextMatrix(i, 0) <> "" Then
        
        rs.AddNew
        rs.Fields("BillNo").Value = txtBillNo.Text
        rs.Fields("RecDate").Value = RecDate.Value
        rs.Fields("PartyName").Value = txtParty.Text
        rs.Fields("ItemName").Value = vs.TextMatrix(i, 0)
        rs.Fields("Size").Value = vs.TextMatrix(i, 1)
        rs.Fields("Grade").Value = vs.TextMatrix(i, 2)
        
        
        rs.Fields("NoofRode").Value = vs.TextMatrix(i, 3)
        rs.Fields("NoofBunde").Value = vs.TextMatrix(i, 4)
        
        rs.Fields("Qty").Value = vs.TextMatrix(i, 5)
        rs.Fields("Unit").Value = vs.TextMatrix(i, 6)
        rs.Fields("Rate").Value = vs.TextMatrix(i, 7)
        If vs.TextMatrix(i, 8) <> "" Then
          rs.Fields("Amount").Value = vs.TextMatrix(i, 8)
        End If
        rs.Fields("TaxPer").Value = txttaxper.Text
        rs.Fields("TaxAmount").Value = txttax.Text
        rs.Fields("NetAmount").Value = txtNet.Text
        rs.Fields("Fright").Value = txtFright.Text
        rs.Fields("HeatNo").Value = vs.TextMatrix(i, 9)
        rs.Update
        
        UpdateIssue
        
       End If
   Next
    
    
End Sub

Private Sub cmdModify_Click()
     
     
    If txtBillNo.Text = "" Or txtParty.Text = "" Or txtTotal.Text = "" Then
       MsgBox "Please Verify Entries !!", vbCritical
       Exit Sub
    End If
    

     
     If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
        
        deleteIssue
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
     
     txtTotal.Text = ""
     txttaxper.Text = 0
     txttax.Text = 0
     txtNet.Text = 0
     txtFright.Text = 0
     txtTotal.Text = 0
     
     
     vs.Clear
     SetWidth
     vs.Rows = 100
     
     
     
     
     cmdsave.Enabled = True
     cmdModify.Enabled = False
     cmdDel.Enabled = False
     Record = ""
     
     txtBillNo.Text = MaxSNo("Invoice", "BillNo")
     
     RecDate.Value = Date
     
     
     
End Sub
Sub UpdateIssue()

Dim rss As New ADODB.Recordset
Dim Search As New ADODB.Recordset
   
If vs.TextMatrix(i, 0) <> "" Then
    
 If vs.TextMatrix(i, 9) <> "" Then
    If rss.State = 1 Then rss.Close
    rss.Open "select * from IssueRawMetrial where HeatingNo=" & vs.TextMatrix(i, 9) & " and ItemName='" & vs.TextMatrix(i, 0) & "'", con, adOpenDynamic, adLockOptimistic
    If rss.EOF = False Then
       rss.Fields("Issue").Value = (CDbl(rss.Fields("Issue").Value) + CDbl(vs.TextMatrix(i, 5)))
       rss.Update
    End If
 End If
   
End If
  
  
End Sub
Private Sub cmdSave_Click()
    
    
    If txtBillNo.Text = "" Or txtParty.Text = "" Or txtTotal.Text = "" Then
       MsgBox "Please Verify Entries !!", vbCritical
       Exit Sub
    End If
    

    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Invoice where BillNo=" & txtBillNo.Text & "", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = True Then
      
      If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
         SaveData
         Call cmdRef_Click
         Call cmdadd_Click
      End If
    Else
        
      MsgBox "Bill No Already Exist !!", vbCritical
      
     
    End If
    
End Sub
Sub UpdateStock()
    
''    Dim rs_u As New ADODB.Recordset
''    Set rs_u = New ADODB.Recordset
''    If rs_u.State = 1 Then rs_u.Close
''    rs_u.Open "select * from Stock where ItemName='" & vs.TextMatrix(i, 0) & "' and Status='Finish'", con, adOpenDynamic, adLockOptimistic
''    If rs_u.EOF = True Then
''        rs_u.AddNew
''        rs_u!itemname = vs.TextMatrix(i, 0)
''        ItemGpSearch vs.TextMatrix(i, 0)
''        rs_u!itemgp = itemgp
''        rs_u!qty = vs.TextMatrix(i, 5)
''        rs_u!unit = vs.TextMatrix(i, 6)
''        rs_u!Rate = vs.TextMatrix(i, 7)
''        rs_u!Status = "Finish"
''        rs_u.Update
''     Else
''        rs_u!qty = rs_u!qty + vs.TextMatrix(i, 5)
''        rs_u.Update
''    End If
    
End Sub

Sub ItemGpSearch(Str As String)
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ItemGp from ItemMaster where ItemName='" & Str & "'", con
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).Value
    End If
End Sub
Sub DeleteStock()
    
''    Dim rs_u As New ADODB.Recordset
''
''    For i = vs.FixedRows To vs.Rows - 1
''
''     If vs.TextMatrix(i, 0) <> "" Then
''        If rs_u.State = 1 Then rs_u.Close
''        rs_u.Open "select * from Stock where ItemName='" & vs.TextMatrix(i, 0) & "' and status='Finish'", con, adOpenDynamic, adLockOptimistic
''        If rs_u.EOF = False Then
''
''            If rs1.State = 1 Then rs1.Close
''            rs1.Open "select * from Invoice where ItemName='" & vs.TextMatrix(i, 0) & "' and BillNo='" & txtBillNo.Text & "'", con, adOpenDynamic, adLockOptimistic
''
''            rs_u!qty = rs_u!qty - rs1.Fields("qty").Value
''            rs_u.Update
''
''        End If
''    End If
''
''    Next
    
End Sub
Sub CaculateTaxAmount()
    If Val(txtTotal.Text) = 0 Then
       txtTotal.Text = 0
    End If
    
    If Val(txtFright.Text) = 0 Then
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
       tbl = "Customer"
       popuplist11 "Select name from Customer order by Name", con, , True, colmn
       popuplist1.Show
    ElseIf KeyCode = 13 Then
       'SendKeys "{tab}"
    End If
End Sub
Sub Search()
    
    On Error Resume Next
    
    vs.Clear
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Invoice where BillNo=" & billno & "", con, adOpenDynamic, adLockOptimistic
    
    
      
    If rs.EOF = False Then
        
        txtBillNo.Text = rs.Fields("BillNo").Value
        RecDate.Value = rs.Fields("RecDate").Value
        txtParty.Text = rs.Fields("PartyName").Value
        txttaxper.Text = rs.Fields("TaxPer").Value
        txttax.Text = rs.Fields("TaxAmount").Value
        txtNet.Text = rs.Fields("NetAmount").Value
        txtFright.Text = rs.Fields("Fright").Value
        txtHeatingNo.Text = rs.Fields("HeatNo").Value
        
    For i = 1 To rs.RecordCount + 2
    
      If rs.EOF = False Then
      
        vs.TextMatrix(i, 0) = rs.Fields("ItemName").Value
        vs.TextMatrix(i, 1) = rs.Fields("Size").Value
        vs.TextMatrix(i, 2) = rs.Fields("Grade").Value
        
        
        vs.TextMatrix(i, 3) = rs.Fields("NoofRode").Value
        vs.TextMatrix(i, 4) = rs.Fields("NoofBunde").Value
        
        
        vs.TextMatrix(i, 5) = rs.Fields("Qty").Value
        vs.TextMatrix(i, 6) = rs.Fields("Unit").Value
        vs.TextMatrix(i, 7) = rs.Fields("Rate").Value
        vs.TextMatrix(i, 8) = rs.Fields("Amount").Value
        vs.TextMatrix(i, 9) = rs.Fields("HeatNo").Value
        
      End If
        
      rs.MoveNext
        
    Next
    
    
  End If
    
  vs.Rows = vs.Rows + 5
        
  Total
        
  SetWidth
     
     
     
     
     cmdModify.Enabled = True
     cmdDel.Enabled = True
     cmdsave.Enabled = False
     
     
     
     

     
End Sub
Sub Total()
    
    txtTotal.Text = 0
    For i = 1 To vs.Rows - 1
       If vs.TextMatrix(i, 0) <> "" Then
        txtTotal.Text = Format((CDbl(txtTotal.Text) + CDbl(vs.TextMatrix(i, 8))), "#,###.00")
       End If
    Next
    
    CaculateTaxAmount
    
End Sub
Private Sub Form_Load()
    SetWidth
    
    colmn = 1
    popuplist11 "Select name from Supplier order by Name", con, , True, colmn
    tbl = "Supplier"
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from TaxHead", con
    If rs.EOF = False Then
       txttaxper.Text = rs(1)
       Tax.Caption = rs(0)
    End If
     
    vs.Rows = 18
      
    txtBillNo.Text = MaxSNo("Invoice", "BillNo")
    Call cmdRef_Click
    
    FromDate.Value = Date
    ToDate.Value = Date

   
End Sub
Sub SetWidth()
    
    vs.FormatString = "Item Name|Size|Grade|No of Rode|No of Bundle|>Qty|^Unit|>Rate|>Amount"
    vs.ColWidth(0) = 2500
    vs.ColWidth(1) = 1100
    vs.ColWidth(2) = 1100
    vs.ColWidth(3) = 1100
    vs.ColWidth(4) = 1200
    vs.ColWidth(5) = 1200
    vs.ColWidth(6) = 1100
    vs.ColWidth(7) = 1000
    vs.ColWidth(8) = 1000
    vs.ColWidth(9) = 0
End Sub

Private Sub ListBillNo_Click()
   billno = ListBillNo
   Search
End Sub
Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      RecDate.SetFocus
      billno = txtBillNo.Text
      Search
      SetWidth
   End If
End Sub
Private Sub txtFright_Change()
    On Error Resume Next
    txttax.Text = (CDbl(txtTotal.Text) * CDbl(txttaxper.Text)) / 100
    txttax.Text = Format(txttax.Text, "#,###.00")
    txtNet.Text = Format((CDbl(txttax.Text) + CDbl(txtTotal.Text)) + CDbl(txtFright.Text), "#,###.00")
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
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtParty_GotFocus()
   dd = False
   txtParty.Text = Record
   tbl = "Supplier"
End Sub
Private Sub txtParty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub txtParty_LostFocus()
   Record = ""
End Sub

Private Sub vs_GotFocus()
    dd = True
    
    If vs.Col = 0 Then
       
    If Check1.Value = 0 Then
       
        'For i = 1 To frmSales.vs.Rows - 1
        'If frmSales.vs.TextMatrix(i, 0) <> "" Then
        '   If frmSales.vs.TextMatrix(i, 0) = PopUpValue2 Then
        '      MsgBox "Item Already Exist !!", vbCritical
         '     Exit Sub
         '  End If
           
        'End If
        'Next
       
       
       
       
       vs.TextMatrix(vs.RowSel, 0) = PopUpValue2
       vs.TextMatrix(vs.RowSel, 1) = PopUpValue4
       vs.TextMatrix(vs.RowSel, 2) = PopUpValue3
       vs.TextMatrix(vs.RowSel, 5) = PopUpValue5
       vs.TextMatrix(vs.RowSel, 9) = PopUpValue1
       
       
       
       PopUpValue2 = ""
       PopUpValue4 = ""
       PopUpValue3 = ""
       PopUpValue5 = ""
       PopUpValue9 = ""
    
    Else
    
        For i = 1 To frmSales.vs.Rows - 1
        If frmSales.vs.TextMatrix(i, 0) <> "" Then
           If frmSales.vs.TextMatrix(i, 0) = PopUpValue1 Then
              MsgBox "Item Already Exist !!", vbCritical
              Exit Sub
           End If
           
        End If
        Next
       
       
       
       
       vs.TextMatrix(vs.RowSel, 0) = PopUpValue1
       vs.TextMatrix(vs.RowSel, 5) = PopUpValue2
       
       
       
       PopUpValue1 = ""
       PopUpValue2 = ""
       
    End If
       
    End If
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
       Dim rs1 As New ADODB.Recordset
       If KeyCode = 46 Then
          vs.RemoveItem (vs.RowSel)
       End If
            
       If vs.Col = 0 Then
          
       If Check1.Value = 0 Then
        
         If KeyCode = 114 Then
           If rs1.State = 1 Then rs1.Close
           popuplist2 "select HeatingNo,ItemName,Grade,Size,Balance,Unit from searchSales where balance>0", con
         End If
        
       Else
        
        If KeyCode = 114 Then
          If rs1.State = 1 Then rs1.Close
          popuplist2 "select ItemName,OpeningStock from ItemMaster where OpeningStock>0 and ItemGp='Finish Item'", con
        End If
       
       
       End If
       
       
       End If
       
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
  
  On Error Resume Next
     
  If KeyCode = 13 Then
     
  
  
  If vs.Col = 0 Then
     
     If rs.State = 1 Then rs.Close
     rs.Open "select * from ItemMaster where ItemName='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
     If rs.EOF = False Then
        vs.TextMatrix(vs.RowSel, 6) = rs.Fields("Unit").Value
        vs.TextMatrix(vs.RowSel, 7) = rs.Fields("Rate").Value
        SendKeys "{right}"
     End If
     
  End If
  
  If vs.Col = 1 Or vs.Col = 2 Or vs.Col = 3 Or vs.Col = 4 Then
        SendKeys "{right}"
  ElseIf vs.Col = 3 Or vs.Col = 5 Then
        SendKeys "{right}"
        SendKeys "{right}"
  ElseIf vs.Col = 7 Then
        vs.TextMatrix(vs.RowSel, 8) = (CDbl(vs.TextMatrix(vs.RowSel, 5)) * CDbl(vs.TextMatrix(vs.RowSel, 7)))
        SendKeys "{home}"
        SendKeys "{down}"
        SendKeys "{right}": SendKeys "{right}": SendKeys "{right}"
        Total
  End If
      
    
    
   
     
     
   
   
  
End If

End Sub
Private Sub vs_LostFocus()
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   PopUpValue4 = ""
   
End Sub
