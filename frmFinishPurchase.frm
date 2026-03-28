VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmFinishPurchase 
   Caption         =   "Item Purchase"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmFinishPurchase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox tot1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6735
      Width           =   1320
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   7305
      Width           =   1320
   End
   Begin VB.TextBox txtOtherCharges 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      TabIndex        =   15
      Top             =   6165
      Width           =   1320
   End
   Begin VB.TextBox txtRoundoff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7020
      Width           =   1320
   End
   Begin VB.CheckBox chkCst 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CST"
      Height          =   255
      Left            =   6135
      TabIndex        =   10
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox chkVat 
      BackColor       =   &H00C0E0FF&
      Caption         =   "VAT"
      Height          =   255
      Left            =   5460
      TabIndex        =   9
      Top             =   5640
      Width           =   675
   End
   Begin VB.CheckBox chkLoading 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Loading"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   5880
      Width           =   915
   End
   Begin VB.CheckBox chkUnloading 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Unloading"
      Height          =   255
      Left            =   5475
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CheckBox chkCartage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cartage"
      Height          =   255
      Left            =   6570
      TabIndex        =   13
      Top             =   5880
      Width           =   855
   End
   Begin Crystal.CrystalReport cr 
      Left            =   180
      Top             =   4860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtdis 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      TabIndex        =   16
      Top             =   6450
      Width           =   1320
   End
   Begin VB.TextBox txteduper 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0"
      Top             =   5310
      Width           =   420
   End
   Begin VB.TextBox txtexiseper 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0"
      Top             =   5025
      Width           =   420
   End
   Begin VB.TextBox txtEduAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5310
      Width           =   1320
   End
   Begin VB.TextBox txtExcise 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5025
      Width           =   1320
   End
   Begin VB.TextBox txtTaxper 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      Top             =   5595
      Width           =   420
   End
   Begin VB.TextBox txtFright 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      TabIndex        =   14
      Top             =   5880
      Width           =   1320
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5595
      Width           =   1320
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
      Height          =   7290
      Left            =   9150
      TabIndex        =   37
      Top             =   105
      Width           =   2535
      Begin VB.CommandButton cmdBill 
         BackColor       =   &H00FFFFFF&
         Caption         =   " &Print Bill Wise"
         Height          =   405
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1500
         Width           =   75
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   " &Print List of purchase made against tax invoice"
         Height          =   585
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1540
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.ListBox ListBillNo 
         Appearance      =   0  'Flat
         Height          =   4515
         Left            =   135
         TabIndex        =   43
         Top             =   2610
         Width           =   2280
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  &Search"
         Height          =   405
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1080
         Width           =   2220
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   1035
         TabIndex        =   39
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   66060289
         CurrentDate     =   38679
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   1035
         TabIndex        =   40
         Top             =   705
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   66060289
         CurrentDate     =   38679
      End
      Begin VB.Label Label6 
         Caption         =   "From Date :"
         Height          =   300
         Left            =   120
         TabIndex        =   42
         Top             =   390
         Width           =   1320
      End
      Begin VB.Label Label7 
         Caption         =   "To Date :"
         Height          =   300
         Left            =   105
         TabIndex        =   41
         Top             =   690
         Width           =   1320
      End
   End
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6180
      TabIndex        =   3
      Top             =   210
      Width           =   1800
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Top             =   570
      Width           =   4155
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   705
      Left            =   240
      TabIndex        =   29
      Top             =   6540
      Width           =   5625
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   495
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   165
         Width           =   1110
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   165
         Width           =   1125
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   165
         Width           =   1125
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   165
         Width           =   1125
      End
   End
   Begin VB.TextBox txtMrvNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1350
      TabIndex        =   0
      Top             =   255
      Width           =   1725
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4740
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker BillDate 
      Height          =   315
      Left            =   6195
      TabIndex        =   4
      Top             =   525
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      Format          =   66060289
      CurrentDate     =   38679
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3795
      Left            =   0
      TabIndex        =   5
      Top             =   900
      Width           =   9030
      _cx             =   15928
      _cy             =   6694
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmFinishPurchase.frx":0442
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
      Left            =   3945
      TabIndex        =   1
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      Format          =   66060289
      CurrentDate     =   38679
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6225
      TabIndex        =   55
      Top             =   6780
      Width           =   1140
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Roundoff"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6600
      TabIndex        =   54
      Top             =   7020
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      Height          =   195
      Index           =   11
      Left            =   6345
      TabIndex        =   53
      Top             =   6210
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Index           =   10
      Left            =   6735
      TabIndex        =   51
      Top             =   6480
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   7245
      TabIndex        =   50
      Top             =   5340
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   7245
      TabIndex        =   49
      Top             =   5025
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Education Cess"
      Height          =   300
      Index           =   9
      Left            =   5610
      TabIndex        =   48
      Top             =   5340
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Excixe"
      Height          =   300
      Index           =   7
      Left            =   5670
      TabIndex        =   47
      Top             =   5025
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   7245
      TabIndex        =   45
      Top             =   5640
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount "
      Height          =   195
      Index           =   8
      Left            =   6480
      TabIndex        =   44
      Top             =   7320
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No. "
      Height          =   300
      Index           =   0
      Left            =   5535
      TabIndex        =   36
      Top             =   255
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date "
      Height          =   300
      Index           =   1
      Left            =   5550
      TabIndex        =   35
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name "
      Height          =   300
      Index           =   2
      Left            =   360
      TabIndex        =   34
      Top             =   585
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MRV Date "
      Height          =   300
      Index           =   3
      Left            =   3150
      TabIndex        =   33
      Top             =   270
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MRV No. "
      Height          =   300
      Index           =   4
      Left            =   345
      TabIndex        =   32
      Top             =   255
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total "
      Height          =   300
      Index           =   5
      Left            =   6840
      TabIndex        =   31
      Top             =   4740
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
      TabIndex        =   30
      Top             =   0
      Width           =   2565
   End
End
Attribute VB_Name = "frmFinishPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dd As Boolean
Dim itemgp As String
Dim i As Integer
Dim m As String
Dim billno
Private Sub BillDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub chkCartage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub chkCst_Click()
CaculateTaxAmount
If chkCst.Value = 1 Then
chkVat.Value = 0
End If
CaculateTaxAmount
End Sub

Private Sub chkCst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub chkCst_LostFocus()
CaculateTaxAmount
End Sub

Private Sub chkLoading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub chkUnloading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub chkVat_Click()
CaculateTaxAmount
If chkVat.Value = 1 Then
chkCst.Value = 0
End If

CaculateTaxAmount
End Sub

Private Sub chkVat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub chkVat_LostFocus()
CaculateTaxAmount
End Sub

Private Sub cmdadd_Click()
   If rs.State = 1 Then rs.Close
   rs.Open "select distinct(MRVNo) from Finishurchase where (RecDate>=datevalue('" & FromDate.Value & "') and RecDate<=datevalue('" & ToDate.Value & "')) order by MRVNo", con, adOpenDynamic, adLockOptimistic
   ListBillNo.Clear
   If rs.EOF = False Then
      While rs.EOF = False
          ListBillNo.AddItem rs.Fields("MRVNo").Value
          rs.MoveNext
      Wend
   End If

End Sub

Private Sub cmdBill_Click()
   
   cr.Reset
   cr.ReportFileName = App.Path & "\BillWisePurchase.rpt"
   If txtParty.Text <> "" Then
      cr.ReplaceSelectionFormula "{Finishurchase.RecDate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.RecDate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.PartyName}= '" & txtParty.Text & "'"
   Else
      cr.ReplaceSelectionFormula "{Finishurchase.RecDate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.RecDate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
   End If
   cr.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & ToDate.Value & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1

End Sub

Private Sub cmdDel_Click()
   
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
      
      DeleteData
      Call cmdRef_Click
      'Call cmdadd_Click
      txtMrvNo.SetFocus
   End If
  
End Sub
Sub DeleteData()
    Dim cmd As New ADODB.Command
    
    'DeleteStock
    
    Set cmd.ActiveConnection = con
    cmd.CommandText = "delete from Finishurchase where BillNo='" & txtBillNo.Text & "'"
    cmd.Execute
    
    
End Sub

Private Sub cmdItem_Click()
   cr.Reset
   cr.ReportFileName = App.Path & "\PartyWisePurchase.rpt"
   If txtParty.Text <> "" Then
      cr.ReplaceSelectionFormula "{Finishurchase.RecDate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.RecDate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.PartyName}= '" & txtParty.Text & "'"
   Else
      cr.ReplaceSelectionFormula "{Finishurchase.RecDate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.RecDate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
   End If
   cr.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & ToDate.Value & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1

End Sub

Private Sub cmdMain_Click()
   Unload Me
End Sub
Sub SaveData()

    Dim str1 As String
    str1 = ""
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Finishurchase", con, adOpenDynamic, adLockOptimistic
    For i = 1 To vs.Rows - 1
    
      If vs.TextMatrix(i, 0) <> "" Then
        rs.AddNew
        rs.Fields("BillNo").Value = txtBillNo.Text
        rs.Fields("BillDate").Value = BillDate.Value
        rs.Fields("RecDate").Value = RecDate.Value
        rs.Fields("PartyName").Value = txtParty.Text
        rs.Fields("MRVNo").Value = txtMrvNo.Text
        rs.Fields("ItemName").Value = vs.TextMatrix(i, 0)
        rs.Fields("Narr").Value = vs.TextMatrix(i, 1)
        rs.Fields("Qty").Value = vs.TextMatrix(i, 2)
        rs.Fields("Unit").Value = vs.TextMatrix(i, 3)
        rs.Fields("Rate").Value = vs.TextMatrix(i, 4)
        rs.Fields("Amount").Value = vs.TextMatrix(i, 5)
        rs.Fields("TaxPer").Value = txtTaxper.Text
        
        If str1 = "" Then
'        a = Len(Mid(Trim(vs.TextMatrix(i, 1)), 1, 25))
'        str1 = Trim(vs.TextMatrix(i, 1)) & Space(26 - a) & Trim(vs.TextMatrix(i, 2)) & Space(7 - Len(Str(Trim(vs.TextMatrix(i, 2))))) & Trim(vs.TextMatrix(i, 3)) & Space(7 - Len(vs.TextMatrix(i, 3))) & vs.TextMatrix(i, 4) & Space(7 - Len(vs.TextMatrix(i, 4))) & vs.TextMatrix(i, 5)
'        Else
'        a = Len(Mid(Trim(vs.TextMatrix(i, 1)), 1, 25))
'        str1 = Trim(vs.TextMatrix(i, 1)) & Space(26 - a) & Trim(vs.TextMatrix(i, 2)) & Space(7 - Len(Str(Trim(vs.TextMatrix(i, 2))))) & Trim(vs.TextMatrix(i, 3)) & Space(7 - Len(vs.TextMatrix(i, 3))) & vs.TextMatrix(i, 4) & Space(7 - Len(vs.TextMatrix(i, 4))) & vs.TextMatrix(i, 5)
        End If

        
        
        rs.Fields("TaxAmount").Value = txtTax.Text
        rs.Fields("NetAmount").Value = txtNet.Text
        
        rs.Fields("Fright").Value = txtFright.Text
        rs.Fields("othecharges").Value = Val(txtOtherCharges.Text)
        rs.Fields("totalwro").Value = tot1.Text
        rs.Fields("amtroundoff").Value = txtRoundoff.Text
        
        rs.Fields("ExcisePer").Value = txtexiseper.Text
        rs.Fields("Excise").Value = txtExcise.Text
        rs.Fields("EducessPer").Value = txteduper.Text
        rs.Fields("Educess").Value = txtEduAmt.Text
        rs.Fields("AddLess").Value = txtdis.Text
    
        rs!cstYN = chkCst.Value
        rs!vatYN = chkVat.Value
        rs!loadingYN = chkLoading.Value
        rs!UnloadingYN = chkUnloading.Value
        rs!CartageYN = chkCartage.Value
    
    
    
    If chkVat.Value = 0 Then
    
     End If
        
        rs.Update
        UpdateStock
       End If
   Next
   
   
    'con.Execute "update Finishurchase set TempData='" & str1 & "' where MRVNo='" & Me.txtMrvNo.Text & "'"
    str1 = ""

   
End Sub
Sub ItemGpSearch(Str As String)
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ItemGp from ItemMaster where ItemName='" & Str & "'", con
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).Value
    End If
End Sub

Sub UpdateStock()
    Dim rs_u As New ADODB.Recordset
    'If rs_u.State = 1 Then rs_u.Close
    'rs_u.Open "select * from Stock where ItemName='" & vs.TextMatrix(i, 0) & "' and ItemGp='" & itemgp & "' and Status='Consume'", con, adOpenDynamic, adLockOptimistic
    'If rs_u.EOF = True Then
    '    rs_u.AddNew
    '    rs_u!itemname = vs.TextMatrix(i, 0)
    '    ItemGpSearch vs.TextMatrix(i, 0)
    '    rs_u!itemgp = itemgp
    '    rs_u!unit = vs.TextMatrix(i, 2)
    '    rs_u!Rate = vs.TextMatrix(i, 3)
    '    rs_u!qty = vs.TextMatrix(i, 1)
    '    rs_u!Status = "Consume"
    '    rs_u.Update
    ' Else
     '   rs_u!qty = rs_u!qty + vs.TextMatrix(i, 1)
     '   rs_u.Update
    'End If
    
End Sub
Sub DeleteStock()
    
'    Dim rs_u As New ADODB.Recordset
'
'    For i = vs.FixedRows To vs.Rows - 1
'
'     If vs.TextMatrix(i, 0) <> "" Then
'        If rs_u.State = 1 Then rs_u.Close
'        rs_u.Open "select * from Stock where ItemName='" & vs.TextMatrix(i, 0) & "' and status='Consume'", con, adOpenDynamic, adLockOptimistic
'        If rs_u.EOF = False Then
'
'            If rs1.State = 1 Then rs1.Close
'            rs1.Open "select * from Finishurchase where ItemName='" & vs.TextMatrix(i, 0) & "' and BillNo='" & txtBillNo.Text & "'", con, adOpenDynamic, adLockOptimistic
'
'            rs_u!qty = rs_u!qty - rs1.Fields("qty").Value
'            rs_u.Update
'
'        End If
'    End If
'
'    Next
    
End Sub

Private Sub cmdModify_Click()
     
    If txtBillNo.Text = "" Or txtParty.Text = "" Or txtTotal.Text = "" Then
       MsgBox "Please Verify Entries !!", vbCritical
       Exit Sub
    End If
    

     
     If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
        
        DeleteStock
        DeleteData
        SaveData
        Call cmdRef_Click
        Call cmdadd_Click
        
   End If
End Sub

Private Sub cmdParty_Click()
   cr.Reset
   cr.ReportFileName = App.Path & "\PartyWisePurchase1.rpt"
   If txtParty.Text <> "" Then
      cr.ReplaceSelectionFormula "{Finishurchase.RecDate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.RecDate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.PartyName}= '" & txtParty.Text & "'"
   Else
      cr.ReplaceSelectionFormula "{Finishurchase.RecDate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.RecDate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
   End If
   cr.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & ToDate.Value & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1

End Sub

Private Sub CmdPrint_Click()
   cr.Reset
   cr.ReportFileName = App.Path & "\PartyWisePurchase.rpt"
   cr.ReplaceSelectionFormula "{Finishurchase.RecDate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {Finishurchase.RecDate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
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
     txtBillNo.Text = ""
     cmdModify.Enabled = False
     cmdDel.Enabled = False
     txtParty.Text = ""
     txtMrvNo.Text = ""
     txtTotal.Text = ""
     vs.Clear
     SetWidth
     vs.Rows = 20
     txtMrvNo.SetFocus
     Record = ""
     
     txtTaxper.Text = 0
     txtTax.Text = 0
     txtNet.Text = 0
     txtFright.Text = 0
     txtTotal.Text = 0
     txtdis.Text = 0
     
     txtExcise.Text = 0
     txtEduAmt.Text = 0
     
     chkVat.Value = 0
     chkCst.Value = 0
     chkLoading.Value = 0
     chkUnloading.Value = 0
     chkCartage.Value = 0
     txtEduAmt.Text = 0
     txtExcise.Text = 0

     
     cmdSave.Enabled = True
     cmdModify.Enabled = False
     cmdDel.Enabled = False
     Record = ""
     
     
    RecDate.Value = Date
    BillDate.Value = Date
    'FromDate.Value = Date
    'ToDate.Value = Date
     
     
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select max(val(mrvno)) from Finishurchase", con
    If Not IsNull(rs1(0)) Then
    Me.txtMrvNo.Text = rs1(0) + 1
    Else
    Me.txtMrvNo.Text = 1
    End If
    
    
     
     
     Call Form_Load
End Sub
Private Sub cmdSave_Click()
    
    
    If txtBillNo.Text = "" Or txtParty.Text = "" Or txtTotal.Text = "" Then
       MsgBox "Please Verify Entries !!", vbCritical
       Exit Sub
    End If
    
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Finishurchase where MRVNo='" & txtMrvNo.Text & "'", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = True Then
      
      If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
         SaveData
         Call cmdRef_Click
         Call cmdadd_Click
      End If
    
    Else
     
     MsgBox "This MRV No Already Exist !!", vbCritical
     Exit Sub
        
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
       Unload Me
    ElseIf KeyCode = 113 Then
       'tbl = "Supplier"
       'popuplist11 "Select name from Supplier order by Name", con, , True, colmn
       'popuplist1.Show
    ElseIf KeyCode = 13 Then
       'SendKeys "{tab}"
    End If
End Sub

Sub Search()
    
    On Error Resume Next
    Dim rs_search As New ADODB.Recordset
    
    vs.Clear
    
    If rs_search.State = 1 Then rs_search.Close
    rs_search.Open "select * from Finishurchase where MRVNo='" & txtMrvNo.Text & "'", con, adOpenDynamic, adLockOptimistic
    If rs_search.EOF = False Then
        
        txtBillNo.Text = rs_search.Fields("BillNo").Value
        BillDate.Value = rs_search.Fields("BillDate").Value
        RecDate.Value = rs_search.Fields("RecDate").Value
        txtParty.Text = rs_search.Fields("PartyName").Value
        txtMrvNo.Text = rs_search.Fields("MRVNo").Value
        txtTaxper.Text = rs_search.Fields("TaxPer").Value
        txtTax.Text = rs_search.Fields("TaxAmount").Value
        txtNet.Text = rs_search.Fields("NetAmount").Value
        txtFright.Text = rs_search.Fields("Fright").Value
              
        txtexiseper.Text = rs_search.Fields("ExcisePer").Value
        txtExcise.Text = rs_search.Fields("Excise").Value
        txteduper.Text = rs_search.Fields("EducessPer").Value
        txtEduAmt.Text = rs_search.Fields("Educess").Value
        
        txtdis.Text = rs_search.Fields("AddLess").Value
        
        txtFright.Text = rs_search.Fields("Fright").Value
        txtOtherCharges.Text = rs_search.Fields("othercharges").Value
        tot1.Text = rs_search.Fields("totalwro").Value
        txtRoundoff.Text = rs_search.Fields("amtroundoff").Value
        
        
        If rs_search.Fields("cstYN") = True Then
        chkCst.Value = 1
        Else: chkCst.Value = 0
        End If
        If rs_search.Fields("vatYN") = True Then
         chkVat.Value = 1
         Else: chkVat.Value = 0
        End If
        If rs_search.Fields("loadingYN") = True Then
        chkLoading.Value = rs_search!loadingYN
        End If
        If rs_search.Fields("UnloadingYN") = True Then
        chkUnloading.Value = 1
        Else: chkUnloading.Value = 0
        End If
        If rs_search.Fields("CartageYN") = True Then
        chkCartage.Value = 1
        Else: chkCartage.Value = 0
         End If
        
        vs.Rows = rs_search.RecordCount + 10
        
        For i = 1 To rs_search.RecordCount
          vs.TextMatrix(i, 0) = rs_search.Fields("ItemName").Value
          vs.TextMatrix(i, 1) = rs_search.Fields("Narr").Value
          vs.TextMatrix(i, 2) = rs_search.Fields("Qty").Value
          vs.TextMatrix(i, 3) = rs_search.Fields("Unit").Value
          vs.TextMatrix(i, 4) = rs_search.Fields("Rate").Value
          vs.TextMatrix(i, 5) = rs_search.Fields("Amount").Value
          rs_search.MoveNext
        Next
        
     Total
        
     
     cmdModify.Enabled = True
     cmdDel.Enabled = True
     cmdSave.Enabled = False
     
     End If
     
     SetWidth

     
End Sub
Sub Total()
    
    txtTotal.Text = 0
    For i = 1 To vs.Rows - 1
       If vs.TextMatrix(i, 0) <> "" Then
        txtTotal.Text = Format((CDbl(txtTotal.Text) + CDbl(vs.TextMatrix(i, 5))), "#,###.00")
       End If
    Next
    
   
   CaculateTaxAmount
    
End Sub

Sub CaculateTaxAmount()

  

    If Val(txtTotal.Text) = 0 Then
       txtTotal.Text = 0
    End If

    If Val(txtFright.Text) = 0 Then
       txtFright.Text = "0.00"

    End If

    If Val(txtdis.Text) = 0 Then
       txtdis.Text = "0.00"
    End If


    If Val(txtTax.Text) = 0 Then
       txtTax.Text = 0
    End If

    If Val(txtExcise.Text) = 0 Then
       txtExcise.Text = 0
    End If

    If Val(txtEduAmt.Text) = 0 Then
       txtEduAmt.Text = 0
    End If
    
    If Val(txtOtherCharges.Text) = 0 Then
       txtOtherCharges.Text = 0
    End If
    If txtexiseper.Text = "" Then
    txtexiseper.Text = 0
    End If
    
    If txteduper.Text = "" Then
    txteduper.Text = 0
    End If
    txtExcise.Text = ((txtexiseper.Text * txtTotal.Text) / 100)
    txtExcise.Text = Format(txtExcise.Text, "0.00")

    txtEduAmt.Text = ((txteduper.Text) * (txtTotal.Text)) / 100
    txtEduAmt.Text = Format(txtEduAmt.Text, "0.00")


    txtTax.Text = ((CDbl(txtTotal.Text) * CDbl(txtTaxper.Text)) / 100)
    txtTax.Text = Format(txtTax.Text, "0.00")

    txtFright.Text = Format(txtFright.Text, "0.00")
    txtdis.Text = Format(txtdis.Text, "0.00")



    tot1.Text = Format((CDbl(txtTax.Text) + CDbl(txtTotal.Text)) + CDbl(txtFright.Text) + CDbl(txtExcise.Text) + CDbl(txtEduAmt.Text) + CDbl(txtOtherCharges.Text))
    tot1.Text = Format(((tot1.Text) - (txtdis.Text)), "0.00")

    txtRoundoff.Text = Format((Val(txtNet.Text) - Val(tot1.Text)), "0.00")

    txtNet.Text = Round(Format(Val(tot1.Text), "0.00"))
    txtRoundoff = Format((Val(txtNet.Text) - Val(tot1.Text)), "0.00")

    
    If chkCst.Value = 1 Then
                If rs.State = 1 Then rs.Close
                rs.Open "select cstper from setup1", con, adOpenDynamic, adLockOptimistic
                txtTaxper.Text = rs!cstper

    ElseIf chkVat.Value = 1 Then
            If rs.State = 1 Then rs.Close
                rs.Open "select vatper from setup1", con, adOpenDynamic, adLockOptimistic
                txtTaxper.Text = rs!vatper

    
    
    
    End If
    
    
End Sub

Private Sub Form_Load()
    SetWidth
    
    colmn = 1
    popuplist11 "Select name from Supplier order by Name", con, , True, colmn
    tbl = "Supplier"
    
    m = ""
    
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


    RecDate.Value = Date
    BillDate.Value = Date
    FromDate.Value = Date
    ToDate.Value = Date
    
    Call cmdadd_Click

    '----------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from TaxHead", con
    If rs.EOF = False Then
       txtTaxper.Text = rs(1)
       'Tax.Caption = rs(0)
    End If

    
pos Me

        If rs.State = 1 Then rs.Close
        rs.Open "select EducationCessPer from setup1", con, adOpenDynamic, adLockOptimistic
        txteduper.Text = rs!EducationCessPer
        
          If rs.State = 1 Then rs.Close
        rs.Open "select ExcisePer from setup1", con, adOpenDynamic, adLockOptimistic
        txtexiseper.Text = rs!ExcisePer


If rs1.State = 1 Then rs1.Close
    rs1.Open "select max(val(mrvno)) from Finishurchase", con
    If Not IsNull(rs1(0)) Then
    Me.txtMrvNo.Text = rs1(0) + 1
    Else
    Me.txtMrvNo.Text = 1
    End If
    
End Sub
Sub SetWidth()
    vs.Cols = 6
    vs.FormatString = "Item Name|Group|>Qty|^Unit|>Rate|>Amount"
    vs.ColWidth(0) = 1800
    vs.ColWidth(1) = 2500
    vs.ColWidth(2) = 1100
    vs.ColWidth(3) = 1100
    vs.ColWidth(4) = 1100
    vs.ColWidth(5) = 1100
End Sub

Private Sub ListBillNo_Click()
  ' billno = ListBillNo
  txtMrvNo.Text = ListBillNo
   Search
End Sub

Private Sub RecDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtParty.SetFocus
   End If
End Sub

Private Sub tot1_GotFocus()
CaculateTaxAmount
End Sub

Private Sub tot1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      BillDate.SetFocus
      billno = txtBillNo.Text
   '   Search
      
   End If
End Sub
Private Sub txtdis_GotFocus()
   txtdis.SelLength = 15
   CaculateTaxAmount
End Sub

Private Sub txtdis_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtdis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
'     txtNet.SetFocus
   End If
End Sub

Private Sub txtdis_LostFocus()
   CaculateTaxAmount
End Sub

Private Sub txtEduAmt_Change()
'CaculateTaxAmount
End Sub

Private Sub txtEduAmt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txteduper_Change()
CaculateTaxAmount
End Sub

Private Sub txteduper_GotFocus()
If rs.State = 1 Then rs.Close
rs.Open "select EducationCessPer from setup1", con, adOpenDynamic, adLockOptimistic
txteduper.Text = rs!EducationCessPer

End Sub

Private Sub txteduper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
If KeyCode = 46 Then
txteduper.Text = 0
End If


End Sub

Private Sub txtExcise_Change()
'CaculateTaxAmount
End Sub

Private Sub txtExcise_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtExcise_LostFocus()
CaculateTaxAmount
End Sub

Private Sub txtexiseper_Change()
CaculateTaxAmount
End Sub

Private Sub txtexiseper_GotFocus()
 If rs.State = 1 Then rs.Close
rs.Open "select ExcisePer from setup1", con, adOpenDynamic, adLockOptimistic
txtexiseper.Text = rs!ExcisePer

End Sub

Private Sub txtexiseper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
If KeyCode = 46 Then
txtexiseper.Text = ""
End If
End Sub

Private Sub txtFright_GotFocus()
  CaculateTaxAmount
  txtFright.SelLength = 15
 
End Sub

Private Sub txtFright_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtFright_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'        txtdis.SetFocus
'   End If
End Sub
Private Sub txtFright_LostFocus()
  CaculateTaxAmount
  txtFright.SelLength = 15
End Sub

Private Sub txtMrvNo_GotFocus()
If PopUpValue1 <> "" Then
txtMrvNo.Text = PopUpValue1
Search
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
PopUpValue3 = ""
End If
End Sub

Private Sub txtMrvNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist2 "SELECT DISTINCT mrvno,RecDate as MRVDate,Partyname,billno from Finishurchase order by RecDate", con
End If
End Sub

Private Sub txtMrvNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then RecDate.SetFocus
End Sub

Private Sub txtNet_GotFocus()
CaculateTaxAmount
End Sub

Private Sub txtNet_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtOtherCharges_GotFocus()
txtOtherCharges.SelLength = 20
End Sub

Private Sub txtOtherCharges_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtParty_GotFocus()
   dd = False
   txtParty.Text = Record
   'tbl = "Supplier"
   
   txtParty.Text = PopUpValue1
   
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   popuplist2 "Select name from Supplier order by Name", con
End If

End Sub

Private Sub txtParty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtBillNo.SetFocus
    End If
End Sub
Private Sub txtParty_LostFocus()
   PopUpValue1 = ""
End Sub

Private Sub txtRoundoff_GotFocus()
CaculateTaxAmount
End Sub

Private Sub txtRoundoff_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtTax_Change()
'CaculateTaxAmount
End Sub

Private Sub txttax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txttaxper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtTaxper_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then txtFright.SetFocus
End Sub

Private Sub txttaxper_LostFocus()
'CaculateTaxAmount
End Sub

Private Sub txtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub vs_GotFocus()
    dd = True
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    For i = 0 To 5
        vs.TextMatrix(vs.RowSel, i) = ""
       Next
End If
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
  
  On Error Resume Next
     
  If KeyCode = 13 Then
     
   If vs.Col = 0 Then
     If rs.State = 1 Then rs.Close
     rs.Open "select * from ItemMaster where ItemName='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
     If rs.EOF = False Then
        vs.TextMatrix(vs.RowSel, 1) = rs.Fields("ItemGp").Value
        vs.TextMatrix(vs.RowSel, 3) = rs.Fields("Unit").Value
        vs.TextMatrix(vs.RowSel, 4) = rs.Fields("Rate").Value
        SendKeys "{right}"
     End If
   ElseIf vs.Col = 1 Then
       SendKeys "{right}"
   End If
   
   If vs.Col = 2 Then
       SendKeys "{right}"
       SendKeys "{right}"
   End If
   
   If vs.Col = 4 Then
     SendKeys "{home}"
     SendKeys "{down}"
     vs.TextMatrix(vs.RowSel, 5) = (CDbl(vs.TextMatrix(vs.RowSel, 4)) * CDbl(vs.TextMatrix(vs.RowSel, 2)))
     Total
   End If
     
   
   
  
  End If

End Sub


