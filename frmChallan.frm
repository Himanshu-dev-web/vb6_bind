VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmChallan 
   Caption         =   "Challan Creation"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtAdd2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7425
      TabIndex        =   48
      Top             =   1125
      Width           =   5505
   End
   Begin VB.TextBox txtfrmName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7440
      TabIndex        =   46
      Top             =   450
      Width           =   5505
   End
   Begin VB.TextBox txtAdd1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7440
      TabIndex        =   45
      Top             =   780
      Width           =   5505
   End
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   9540
      TabIndex        =   24
      Top             =   8520
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtTaxRebate_Per 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9870
      TabIndex        =   23
      Text            =   "0"
      Top             =   9120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtTaxRebate_Amt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10500
      TabIndex        =   22
      Top             =   9120
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox txtEntryTax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10500
      TabIndex        =   21
      Top             =   8820
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox txtEntry_PTax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9870
      TabIndex        =   20
      Text            =   "0"
      Top             =   8820
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtHighEduAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   8895
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtHighereduper 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10140
      TabIndex        =   18
      Text            =   "0"
      Top             =   8760
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CheckBox chkCst 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CST"
      Height          =   255
      Left            =   10140
      TabIndex        =   17
      Top             =   8760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkVat 
      BackColor       =   &H00C0E0FF&
      Caption         =   "VAT"
      Height          =   255
      Left            =   10080
      TabIndex        =   16
      Top             =   8760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txteduper 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10260
      TabIndex        =   15
      Text            =   "0"
      Top             =   8700
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtexiseper 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10140
      TabIndex        =   14
      Text            =   "0"
      Top             =   8700
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtEduAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   8610
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtExcise 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   8745
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtTaxper 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10200
      TabIndex        =   11
      Text            =   "0"
      Top             =   8700
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   9480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1845
      TabIndex        =   9
      Top             =   975
      Width           =   5370
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   885
      Left            =   360
      TabIndex        =   2
      Top             =   6060
      Width           =   8325
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   735
         Left            =   60
         Picture         =   "frmChallan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1350
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   735
         Left            =   6960
         Picture         =   "frmChallan.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1350
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4230
         Picture         =   "frmChallan.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1350
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   735
         Left            =   2820
         Picture         =   "frmChallan.frx":1468
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1350
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   735
         Left            =   1440
         Picture         =   "frmChallan.frx":1875
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1350
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   735
         Left            =   5640
         Picture         =   "frmChallan.frx":2459
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   1260
      End
   End
   Begin VB.TextBox txtMrvNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   540
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   " &Print List of purchase made against tax invoice"
      Height          =   585
      Left            =   9660
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Visible         =   0   'False
      Width           =   2190
   End
   Begin Crystal.CrystalReport cr 
      Left            =   11700
      Top             =   6300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3975
      Left            =   300
      TabIndex        =   25
      Top             =   1845
      Width           =   13710
      _cx             =   24183
      _cy             =   7011
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
      BackColorFixed  =   8569087
      ForeColorFixed  =   -2147483630
      BackColorSel    =   11461627
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmChallan.frx":298B
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
      WordWrap        =   -1  'True
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
      Left            =   3180
      TabIndex        =   26
      Top             =   540
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   64487425
      CurrentDate     =   38679
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Firm Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   6300
      TabIndex        =   47
      Top             =   450
      Width           =   1125
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sub Total (2) :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   9780
      TabIndex        =   44
      Top             =   8700
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   10365
      TabIndex        =   43
      Top             =   9120
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rebate "
      Height          =   300
      Index           =   8
      Left            =   9960
      TabIndex        =   42
      Top             =   8700
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Tax"
      Height          =   300
      Index           =   6
      Left            =   9960
      TabIndex        =   41
      Top             =   8700
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   10335
      TabIndex        =   40
      Top             =   8820
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblsubtotal2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   11055
      TabIndex        =   39
      Top             =   9885
      Width           =   165
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sub Total (1) :"
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
      Left            =   9840
      TabIndex        =   38
      Top             =   8760
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Label lblsubtotal1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   10830
      TabIndex        =   37
      Top             =   9195
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Higher Education Cess"
      Height          =   300
      Index           =   4
      Left            =   9900
      TabIndex        =   36
      Top             =   8700
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   10335
      TabIndex        =   35
      Top             =   8880
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Case Memo  No. "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   360
      TabIndex        =   34
      Top             =   540
      Width           =   1635
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   10365
      TabIndex        =   33
      Top             =   8940
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   10860
      TabIndex        =   32
      Top             =   9840
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Education Cess"
      Height          =   300
      Index           =   9
      Left            =   9840
      TabIndex        =   31
      Top             =   8760
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Excixe"
      Height          =   300
      Index           =   7
      Left            =   9900
      TabIndex        =   30
      Top             =   8760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "%"
      Height          =   195
      Left            =   10320
      TabIndex        =   29
      Top             =   8940
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   360
      TabIndex        =   28
      Top             =   1005
      Width           =   1275
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
      TabIndex        =   27
      Top             =   180
      Width           =   2565
   End
End
Attribute VB_Name = "frmChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dd As Boolean
Dim itemgp As String
Dim i As Integer
Dim m As String
Dim rs As New ADODB.Recordset
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
'   If rs.State = 1 Then rs.Close
'   rs.Open "select distinct(CH_No) from challan where (RecDate>=datevalue('" & FromDate.Value & "') and RecDate<=datevalue('" & ToDate.Value & "')) order by CH_No", con, adOpenDynamic, adLockOptimistic
'   ListBillNo.Clear
'   If rs.EOF = False Then
'      While rs.EOF = False
'          ListBillNo.AddItem rs.Fields("CH_No").Value & ""
'          rs.MoveNext
'      Wend
'   End If

End Sub

Private Sub cmdBill_Click()
   
   cr.Reset
   cr.ReportFileName = App.Path & "\BillWisePurchase.rpt"
   If txtParty.Text <> "" Then
      cr.ReplaceSelectionFormula "{challan.RecDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {challan.RecDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {challan.PartyName}= '" & txtParty.Text & "'"
   Else
      cr.ReplaceSelectionFormula "{challan.RecDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {challan.RecDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
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

Private Sub cmdDel_Click()
   
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
      
      DeleteData
      Call cmdRef_Click
      ' txtBillNo.SetFocus
   End If
  
End Sub
Sub DeleteData()
    Dim cmd As New ADODB.Command
    
     
    Set cmd.ActiveConnection = con
    cmd.CommandText = "delete from challan where ch_No='" & txtMrvNo.Text & "' and FirmName='" & txtfrmName & "'"
    cmd.Execute
    
    
End Sub

Private Sub cmdItem_Click()
   cr.Reset
   cr.ReportFileName = App.Path & "\PartyWisePurchase.rpt"
   If txtParty.Text <> "" Then
      cr.ReplaceSelectionFormula "{challan.RecDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {challan.RecDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {challan.PartyName}= '" & txtParty.Text & "'"
   Else
      cr.ReplaceSelectionFormula "{challan.RecDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {challan.RecDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
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

Private Sub cmdMain_Click()
   Unload Me
End Sub
Sub SaveData()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from challan", con, adOpenDynamic, adLockOptimistic
    For i = 1 To vs.Rows - 1
    
      If vs.TextMatrix(i, 0) <> "" Then
        rs.AddNew
        rs.Fields("ch_No").Value = txtMrvNo.Text
        rs.Fields("ch_Date").Value = RecDate.Value
        rs.Fields("PartyName").Value = txtParty.Text
        rs.Fields("ItemName").Value = vs.TextMatrix(i, 0)
        rs.Fields("desc1").Value = vs.TextMatrix(i, 1)
        rs.Fields("desc2").Value = vs.TextMatrix(i, 2)
        rs.Fields("desc3").Value = vs.TextMatrix(i, 3)
        rs.Fields("FirmName").Value = txtfrmName.Text
        rs.Fields("add1").Value = txtAdd1.Text
        rs.Fields("add2").Value = txtAdd2.Text
        
        rs.Update
       
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

Sub UpdateStock()
'    Dim rs_u As New ADODB.Recordset
'    If rs_u.State = 1 Then rs_u.Close
'    rs_u.Open "select * from Stock where ItemName='" & vs.TextMatrix(i, 0) & "' and ItemGp='" & itemgp & "' and Status='Consume'", con, adOpenDynamic, adLockOptimistic
'    If rs_u.EOF = True Then
'        rs_u.AddNew
'        rs_u!itemname = vs.TextMatrix(i, 0)
'        ItemGpSearch vs.TextMatrix(i, 0)
'        rs_u!itemgp = itemgp
'        rs_u!unit = vs.TextMatrix(i, 2)
'        rs_u!Rate = vs.TextMatrix(i, 3)
'        rs_u!qty = vs.TextMatrix(i, 1)
'        rs_u!Status = "Consume"
'        rs_u.Update
'     Else
'        rs_u!qty = rs_u!qty + vs.TextMatrix(i, 1)
'        rs_u.Update
'    End If
    
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
'            rs1.Open "select * from challan where ItemName='" & vs.TextMatrix(i, 0) & "' and BillNo='" & txtBillNo.Text & "'", con, adOpenDynamic, adLockOptimistic
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
     
    

     
     If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
        
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
      cr.ReplaceSelectionFormula "{challan.RecDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {challan.RecDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "') and {challan.PartyName}= '" & txtParty.Text & "'"
   Else
      cr.ReplaceSelectionFormula "{challan.RecDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {challan.RecDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
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

Private Sub CmdPrint_Click()
   cr.Reset
   cr.ReportFileName = App.Path & "\PartyWisePurchase.rpt"
   cr.ReplaceSelectionFormula "{challan.RecDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {challan.RecDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   
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

     'txtFormNo.Text = ""
     'txtBillNo.Text = ""
     cmdModify.Enabled = False
     cmdDel.Enabled = False
     txtParty.Text = ""
     txtMrvNo.Text = ""
     vs.clear
     SetWidth
     vs.Rows = 20
     'txtBillNo.SetFocus
     Record = ""
     
     txtfrmName.Text = ""
     txtAdd1.Text = ""
     txtAdd2.Text = ""
     
     txtTaxper.Text = 0
     txtTax.Text = 0
     
     txtEntryTax.Text = 0
     txtHighEduAmt.Text = 0
     
     txtExcise.Text = 0
     txtEduAmt.Text = 0
     maxmrvno
     chkVat.Value = 0
     chkCst.Value = 0
     'chkLoading.Value = 0
     'chkUnloading.Value = 0
     'chkCartage.Value = 0
     txtEduAmt.Text = 0
     txtExcise.Text = 0
     
     'txtRoundoff.Text = 0
     
     
     
     cmdSave.Enabled = True
     cmdModify.Enabled = False
     cmdDel.Enabled = False
     Record = ""
     
     
    RecDate.Value = Date
    'BillDate.Value = Date
    'FromDate.Value = Date
    'ToDate.Value = Date
     Call Form_Load
     
     txtMrvNo.SetFocus
     'RecDate.SetFocus
     
End Sub

Private Sub cmdSave_Click()
    
    If txtParty.Text = "" Then
       txtParty.Text = "-"
    End If
    
    
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from challan where ch_No='" & txtMrvNo.Text & "' and FirmName='" & txtfrmName & "'", con, adOpenDynamic, adLockOptimistic
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

Private Sub Command1_Click()
 
 If txtMrvNo = "" Then
    MsgBox "Enter Challan No", vbCritical
    Exit Sub
 End If
 
 cr.Reset
 cr.ReportFileName = App.Path & "\challan1.rpt"
 'cr.DataFiles(0) = App.Path & "\" + Trim(Main.Directory) + "\data.mdb"
 'cr.DataFiles(1) = App.Path & "\" + Trim(Main.Directory) + "\data.mdb"
 cr.ReplaceSelectionFormula "{challan.ch_no} = '" & txtMrvNo & "' and {challan.firmname} = '" & txtfrmName & "'"
 cr.WindowShowPrintSetupBtn = True
 cr.WindowShowPrintBtn = True
 cr.WindowShowRefreshBtn = True
 cr.Action = 1
 

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
Sub maxmrvno()
 
 If txtfrmName.Text = "" Then Exit Sub
 
 If rs.State = 1 Then rs.Close
 rs.Open "select max(val(CH_No)) from challan where FirmName='" & txtfrmName.Text & "'", con, adOpenDynamic, adLockOptimistic
 If Not IsNull(rs(0)) Then
 Me.txtMrvNo.Text = rs(0) + 1
 Else
 Me.txtMrvNo.Text = 1
 End If
  
End Sub
Sub updateSetup()
con.Execute "update setup1 set ExcisePer='" & Val(txtexiseper) & "',EducationCessPer='" & Val(txteduper.Text) & "',HighEduCess=" & Val(txtHighereduper.Text) & ",vatper=" & Val(txtTaxper.Text) & ",entryTax=" & Val(txtEntry_PTax) & ",Rebateper=" & Val(txtTaxRebate_Per) & ""
CaculateTaxAmount
End Sub
Sub Search()
    
'On Error Resume Next
    
On Error GoTo abc:
    
    
    Dim rs_search As New ADODB.Recordset
    
    vs.clear
    
    If rs_search.State = 1 Then rs_search.Close
    rs_search.Open "select * from challan where ch_No='" & txtMrvNo.Text & "' and FirmName='" & txtfrmName & "'", con, adOpenDynamic, adLockOptimistic
    If rs_search.EOF = False Then
        
        RecDate.Value = rs_search.Fields("ch_Date").Value
        txtParty.Text = rs_search.Fields("PartyName").Value
        txtMrvNo.Text = rs_search.Fields("CH_No").Value
        txtfrmName.Text = rs_search.Fields("FirmName").Value & ""
        txtAdd1.Text = rs_search.Fields("add1").Value & ""
        txtAdd2.Text = rs_search.Fields("add2").Value & ""
        
        
        vs.Rows = rs_search.RecordCount + 10
        
        For i = 1 To rs_search.RecordCount
          vs.TextMatrix(i, 0) = rs_search.Fields("ItemName").Value
          vs.TextMatrix(i, 1) = rs_search.Fields("desc1").Value
          vs.TextMatrix(i, 2) = rs_search.Fields("desc2").Value
          vs.TextMatrix(i, 3) = rs_search.Fields("desc3").Value
          rs_search.MoveNext
        Next
        
     
        cmdModify.Enabled = True
        cmdDel.Enabled = True
        cmdSave.Enabled = False
     
     End If
     
     SetWidth

Exit Sub


abc:

MsgBox "" & Err.Description
     
End Sub

Sub Total()
    
    txtTotal.Text = 0
    For i = 1 To vs.Rows - 1
       If vs.TextMatrix(i, 0) <> "" Then
       If Val(vs.TextMatrix(i, 6)) > 0 Then
        txtTotal.Text = Format((CDbl(txtTotal.Text) + CDbl(vs.TextMatrix(i, 6))), "#,###.00")
       End If
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
    
    If Val(txtTaxRebate_Amt.Text) = 0 Then
       txtTaxRebate_Amt.Text = 0
    End If
    
    
    txtExcise.Text = Round(((txtexiseper.Text * txtTotal.Text) / 100) + 0.0001, 0)
    txtExcise.Text = Format(txtExcise.Text, "0.00")

    Dim amt_exise, rebAmt, sum1 As Double
    amt_exise = txtExcise.Text
    rebAmt = 0
    sum1 = 0
    
    txtEduAmt.Text = Round((((txteduper.Text) * (amt_exise)) / 100) + 0.0001, 0)
    
    txtEduAmt.Text = Format(txtEduAmt.Text, "0.00")

    
    
    txtTax.Text = Format(txtTax.Text, "0.00")

    
    txtFright.Text = Format(txtFright.Text, "0.00")
    txtdis.Text = Format(txtdis.Text, "0.00")

    '-------------------------------------------------

    'txtHighEduAmt.Text = ((txtHighereduper.Text) * Val(txtEduAmt.Text)) / 100
    
    txtHighEduAmt.Text = Round((Val(txtEduAmt.Text) / 2) + 0.0001, 0)
    
    txtHighEduAmt.Text = Format(txtHighEduAmt.Text, "0.00")

    'txtEntryTax.Text = Round((((txtEntry_PTax.Text) * (txtTotal.Text)) / 100) + 0.0001, 0)
    'txtEntryTax.Text = Format(txtEntryTax.Text, "0.00")

    
   '----------------------------------------------------------------------
    '-------------------------------------------------
    lblsubtotal1.Caption = (CDbl(txtTotal.Text) + CDbl(txtEduAmt.Text) + CDbl(txtExcise.Text) + CDbl(txtHighEduAmt.Text))
    lblsubtotal1.Caption = Format(lblsubtotal1.Caption, "0.00")
    
    '-------------------------------------------------
    
    'txttax.Text = ((CDbl(txtTotal.Text) * CDbl(txttaxper.Text)) / 100)
    txtTax.Text = Round(((CDbl(lblsubtotal1.Caption) * CDbl(txtTaxper.Text)) / 100), 0)
 
    
    lblsubtotal2.Caption = Round((Val(lblsubtotal1.Caption) + Val(txtTax.Text)) + 0.0001, 0)
    lblsubtotal2.Caption = Format(lblsubtotal2.Caption, "0.00")
    
   '-------------Entry Tax----------------------------------------------------------
   
   If Val(txtEntry_PTax.Text) > 0 Then
    txtEntryTax.Text = Round(((CDbl(lblsubtotal2.Caption) * CDbl(txtEntry_PTax.Text)) / 100) + 0.0001, 0)
    txtEntryTax.Text = Format(txtEntryTax.Text, "0.00")
   End If
   '-----------------------------------------------------------------------
    
    sum1 = (Val(lblsubtotal2) + Val(txtEntryTax.Text))
    If Val(txtTaxRebate_Per.Text) > 0 Then
       rebAmt = Round(((sum1 * Val(txtTaxRebate_Per.Text)) / 100) + 0.0001, 0)
       txtTaxRebate_Amt.Text = Format(rebAmt, "0.00")
    End If
 
    

    tot1.Text = Format((CDbl(txtTax.Text) + CDbl(txtTotal.Text)) + CDbl(txtFright.Text) + CDbl(txtExcise.Text) + CDbl(txtEduAmt.Text) + CDbl(txtOtherCharges.Text) + Val(txtHighEduAmt.Text) + Val(txtEntryTax.Text))
    tot1.Text = Format(((tot1.Text) - (Val(txtdis.Text) + Val(txtTaxRebate_Amt.Text))), "0.00")

    
    txtNet.Text = Round(Val(tot1.Text), 0)
    txtNet.Text = Format(txtNet.Text, "0.00")
    
   
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
    
    RecDate.Value = Date
    
    SetWidth
    
    
'    m = ""
'
'    If rs.State = 1 Then rs.close
'    rs.Open "select papermaker_name from papermakeMaster order by papermaker_name", CON, adOpenDynamic, adLockOptimistic, adCmdText
'    While rs.EOF = False
'
'        If rs.EOF = False Then
'
'            If m = "" Then
'               m = rs!papermaker_name & ""
'            Else
'               m = m & "|" & rs!papermaker_name
'            End If
'
'
'       End If
'
'       rs.MoveNext
'     Wend
'
'
'    vs.ColComboList(0) = m
    
    
  

    
    Call cmdadd_Click

 
    maxmrvno


End Sub
Sub SetWidth()
    
    vs.FormatString = "Product Name|<Description-1|<Description-2|<Description-3"
    vs.ColWidth(0) = 3500
    vs.ColWidth(1) = 3000
    vs.ColWidth(2) = 3000
    vs.ColWidth(3) = 3000
    
   For i = 0 To 3
     vs.Cell(flexcpFontSize, 0, i) = 11
  Next

    
End Sub

Private Sub ListBillNo_Click()
  ' billno = ListBillNo
  txtMrvNo.Text = ListBillNo
   Search
End Sub

Private Sub RecDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtfrmName.SetFocus
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
Private Sub txtBillNo_LostFocus()
   
   If Len(txtBillNo) = 0 Then
      txtBillNo = "-"
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

Private Sub txteduper_LostFocus()
updateSetup
End Sub

Private Sub txtEntry_PTax_LostFocus()
updateSetup
End Sub

Private Sub txtEntryTax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   updateSetup
   txtTaxRebate_Amt.SetFocus
End If
End Sub

Private Sub txtEntryTax_LostFocus()
updateSetup
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

Private Sub txtexiseper_LostFocus()
updateSetup
End Sub

Private Sub txtFright_GotFocus()
  
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


Private Sub txtfrmName_GotFocus()

If PopUpValue1 <> "" Then

    txtfrmName.Text = PopUpValue1
    txtAdd1.Text = PopUpValue2
    txtAdd2.Text = PopUpValue3
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    
    maxmrvno
    
End If

End Sub

Private Sub txtfrmName_KeyDown(KeyCode As Integer, Shift As Integer)

If (KeyCode = 113 Or KeyCode = 255) Then
   popuplist2 "select Name,Address,TIN from firm order by name", con
End If

If KeyCode = 13 Then
   If txtfrmName = "" Then Exit Sub
   txtParty.SetFocus
End If

End Sub

Private Sub txtHighereduper_LostFocus()
updateSetup
End Sub

Private Sub txtMrvNo_GotFocus()
txtMrvNo.SelLength = 10
If PopUpValue1 <> "" Then
 txtMrvNo.Text = PopUpValue1
 txtfrmName.Text = PopUpValue4
Search
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
PopUpValue4 = ""
End If
End Sub

Private Sub txtMrvNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
  popuplist2 "SELECT DISTINCT CH_No as ChallanNo,ch_Date as Dates,partyname as Customer,FirmName from challan", con
End If

If KeyCode = 13 Then
   Search
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
txtOtherCharges.SelLength = 15
End Sub

Private Sub txtOtherCharges_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub txtParty_GotFocus()
   dd = False
   If PopUpValue1 <> "" Then
   
   txtParty.Text = PopUpValue1
   PopUpValue1 = ""
   
   End If
   
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   popuplist2 "select Name from customer order by name", con
End If

End Sub

Private Sub txtParty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       vs.Col = 0
       vs.Row = 1
       vs.SetFocus
    End If
End Sub
Private Sub txtParty_LostFocus()
   If Len(txtParty) = 0 Then
      txtParty = "-"
   End If
   
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

Private Sub txttax_GotFocus()
txtTax.SelLength = 15
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

Private Sub txtTaxper_LostFocus()
'CaculateTaxAmount
updateSetup
End Sub


Private Sub txtTaxRebate_Amt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFright.SetFocus
   updateSetup

End If
End Sub

Private Sub txtTaxRebate_Per_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtTaxRebate_Amt.SetFocus
End If
End Sub

Private Sub txtTaxRebate_Per_LostFocus()
updateSetup
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
    For i = 0 To 4
        vs.TextMatrix(vs.RowSel, i) = ""
       Next
End If
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
  
  On Error Resume Next
     
  If KeyCode = 13 Then
  
    If vs.Col = 0 Then
          
        If vs.TextMatrix(vs.RowSel, 0) <> "" Then
           SendKeys "{right}"
       End If
       
     ElseIf vs.Col = 1 Then
         SendKeys "{right}"
     ElseIf vs.Col = 2 Then
         SendKeys "{right}"
     ElseIf vs.Col = 3 Then
         SendKeys "{home}"
         SendKeys "{down}"
     End If
  
  End If

End Sub




