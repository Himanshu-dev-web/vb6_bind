VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmIssueConsumeItem 
   Caption         =   "Issue Item "
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIssueConsumeItem.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5700
      TabIndex        =   29
      Top             =   6120
      Width           =   1770
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   4005
   End
   Begin Crystal.CrystalReport cr 
      Left            =   8040
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboDeptt 
      Height          =   315
      Left            =   4020
      TabIndex        =   28
      Top             =   180
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.TextBox txtHeating 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1665
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   690
      Left            =   330
      TabIndex        =   18
      Top             =   6675
      Width           =   5385
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   450
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   450
         Left            =   4230
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   450
         Left            =   2145
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   450
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   450
         Left            =   1110
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1665
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1350
      Width           =   5820
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3540
      TabIndex        =   17
      Top             =   6120
      Width           =   1410
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
      Height          =   5250
      Left            =   7725
      TabIndex        =   13
      Top             =   2160
      Width           =   3810
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   345
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   450
         Width           =   1170
      End
      Begin VB.ListBox ListHeatingNo 
         Appearance      =   0  'Flat
         Height          =   3930
         Left            =   1140
         TabIndex        =   14
         Top             =   1215
         Width           =   2595
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  &Search  "
         Height          =   345
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   810
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   1170
         TabIndex        =   9
         Top             =   450
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   38679
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   1170
         TabIndex        =   10
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   38679
      End
      Begin VB.Label Label6 
         Caption         =   "From Date :"
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   495
         Width           =   1320
      End
      Begin VB.Label Label7 
         Caption         =   "To Date :"
         Height          =   300
         Left            =   105
         TabIndex        =   15
         Top             =   750
         Width           =   1320
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3750
      Left            =   240
      TabIndex        =   4
      Top             =   2340
      Width           =   7395
      _cx             =   13044
      _cy             =   6615
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
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIssueConsumeItem.frx":0442
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
   Begin MSComCtl2.DTPicker HeatingDate 
      Height          =   300
      Left            =   1665
      TabIndex        =   1
      Top             =   615
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   529
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38679
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "F3 For Search Record"
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
      Index           =   1
      Left            =   4080
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   300
      Index           =   2
      Left            =   195
      TabIndex        =   25
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue No"
      Height          =   270
      Index           =   0
      Left            =   210
      TabIndex        =   24
      Top             =   315
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date "
      Height          =   300
      Index           =   1
      Left            =   195
      TabIndex        =   23
      Top             =   645
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "Raw Metrial Issue"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   300
      TabIndex        =   22
      Top             =   1950
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks "
      Height          =   300
      Index           =   4
      Left            =   225
      TabIndex        =   21
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Total :"
      Height          =   300
      Index           =   6
      Left            =   2940
      TabIndex        =   20
      Top             =   6120
      Width           =   600
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Height          =   60
      Left            =   0
      TabIndex        =   19
      Top             =   1785
      Width           =   12075
   End
End
Attribute VB_Name = "frmIssueConsumeItem"
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

Private Sub cboName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub cmdadd_Click()
 If rs.State = 1 Then rs.Close
 rs.Open "select distinct(IssueNo) from IssueConsumeItem where IssueDate >=datevalue('" & fromDate.Value & "') and IssueDate <=datevalue('" & todate.Value & "') order by IssueNo", con
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
   cr.ReportFileName = App.Path & "\DepttWise.rpt"
   If cboDeptt.Text <> "" Then
      cr.ReplaceSelectionFormula "{IssueConsumeItem.IssueDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {IssueConsumeItem.IssueDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {IssueConsumeItem.Deptt}= '" & cboDeptt.Text & "'"
   Else
      cr.ReplaceSelectionFormula "{IssueConsumeItem.IssueDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {IssueConsumeItem.IssueDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   End If
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
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
      txtHeating.Text = MaxSNo("IssueConsumeItem", "IssueNo")
      HeatingDate.Value = Date

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
    rs.Open "select * from IssueConsumeItem where issueNo=" & txtHeating.Text & "", con
    If rs.EOF = True Then
       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
          SaveData
       End If
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
                rs.Fields("Item").Value = vs.TextMatrix(i, 0)
                rs.Fields("Unit").Value = vs.TextMatrix(i, 1)
                rs.Fields("Qty").Value = vs.TextMatrix(i, 2)
                rs.Fields("Remarks").Value = txtRemarks.Text
                rs.Fields("deptt").Value = cboDeptt.Text
                rs.Fields("customerName").Value = cboName.Text
                rs.Fields("rate").Value = vs.TextMatrix(i, 3)
                rs.Update
                
                UpdatePurchase
              End If
              
              
              
           Next
          
End Sub

Sub DeletePurchase()
Dim balance As Double
'----------------------------------------------
  balance = 0
  balance = Val(vs.TextMatrix(i, 2))
aa:
  
  If rs1.State = 1 Then rs1.Close
  rs1.Open "select BillNo,BillDate,qty,QtyRec from Finishurchase where ItemName='" & vs.TextMatrix(i, 0) & "' and ((Qty-qtyrec)<=0 or (issue)<>0) order by BillDate desc ", con
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
     con.Execute "update Finishurchase set QtyRec=QtyRec+" & (balance) & ",Issue=" & Me.txtHeating.Text & " where (ItemName='" & vs.TextMatrix(i, 0) & "' and  BillNo='" & rs1.Fields("billno").Value & "')"
  End If
  End If
'-----------------------------------------------
End Sub
Sub SearchData()
    On Error Resume Next
    vs.Clear
    cmdRef_Click
    If rs.State = 1 Then rs.Close
    rs.Open "select * from IssueConsumeItem where IssueNo=" & ListHeatingNo.Text & "", con
    If rs.EOF = False Then
      txtHeating.Text = rs.Fields("IssueNo").Value
      HeatingDate.Value = rs.Fields("IssueDate").Value
      txtRemarks.Text = rs.Fields("Remarks").Value & ""
      cboDeptt.Text = rs.Fields("deptt").Value
      cboName.Text = rs.Fields("customerName").Value
      vs.Rows = rs.RecordCount + 10
      For i = vs.FixedRows To vs.Rows - 1
        If rs.EOF = False Then
          vs.TextMatrix(i, 0) = rs.Fields("Item").Value
          vs.TextMatrix(i, 1) = rs.Fields("Unit").Value
          vs.TextMatrix(i, 2) = rs.Fields("Qty").Value
          vs.TextMatrix(i, 3) = rs.Fields("rate").Value
          vs.TextMatrix(i, 4) = (rs.Fields("rate").Value * rs.Fields("Qty").Value)
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
End Sub
Private Sub Form_Load()
 
 SetWidth
 txtHeating.Text = MaxSNo("IssueConsumeItem", "IssueNo")
 
 HeatingDate.Value = Date
 fromDate.Value = Date
 todate.Value = Date
 
 addCustomer
       
 pos Me
 
 Dim s As String
 s = ""
 If rs.State = 1 Then rs.Close
 rs.Open "select ItemName from ItemMaster order by ItemName", con, adOpenKeyset, adLockReadOnly
 While rs.EOF = False
 If s = "" Then
 s = rs(0)
 Else
 s = s & "|" & rs(0)
 End If
 rs.MoveNext
 Wend
 
 vs.ColComboList(0) = s
 
End Sub

Sub SetWidth()
    vs.Cols = 5
    vs.FormatString = "Item Name|Unit|>Qty|Rate|>Amount"
    vs.ColWidth(0) = 2500
    vs.ColWidth(1) = 100
    vs.ColWidth(2) = 1300
    vs.ColWidth(3) = 1300
    vs.ColWidth(4) = 1300
    
    
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
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub vs_GotFocus()
   If PopUpValue1 <> "" Then
      vs.TextMatrix(vs.RowSel, 0) = PopUpValue1
      vs.TextMatrix(vs.RowSel, 1) = PopUpValue2
      'vs.TextMatrix(vs.RowSel, 3) = PopUpValue3
      
      If rs.State = 1 Then rs.Close
      rs.Open "select Rate from Finishurchase where ItemName='" & PopUpValue1 & "' and (Qty-qtyrec)>0 order by BillDate", con
      If rs.EOF = False Then
         vs.TextMatrix(vs.RowSel, 3) = rs.Fields(0).Value
      End If
      
'  If rs.State = 1 Then rs.Close
'  rs.Open "select rate from itemmaster where ItemName='" & PopUpValue1 & "'", con
'
'  If rs.EOF = False Then
'   vs.TextMatrix(vs.RowSel, 3) = rs.Fields(0).Value
'   End If
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
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where ItemName='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
        If rs.EOF = False Then
           vs.TextMatrix(vs.RowSel, 1) = rs.Fields("Unit").Value
           SendKeys "{right}"
           SendKeys "{right}"
        End If
    ElseIf vs.Col = 2 Then
        SendKeys "{right}"
    ElseIf vs.Col = 3 Then
           vs.TextMatrix(vs.RowSel, 4) = (CDbl(vs.TextMatrix(vs.RowSel, 2)) * CDbl(vs.TextMatrix(vs.RowSel, 3)))
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
        txtTotal1.Text = Format((CDbl(txtTotal1.Text) + CDbl(vs.TextMatrix(i, 2))), "#,###.000")
        txtAmt.Text = Format((CDbl(txtAmt.Text) + CDbl(vs.TextMatrix(i, 4))), "#,###.00")
    Next
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

