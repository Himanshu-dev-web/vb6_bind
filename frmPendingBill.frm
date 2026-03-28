VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPendingBill 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Pending Bill"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11496
   Icon            =   "frmPendingBill.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11496
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CboPName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   1524
      TabIndex        =   10
      Top             =   1140
      Width           =   6444
   End
   Begin VB.ComboBox cbogp 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1524
      TabIndex        =   9
      Top             =   780
      Width           =   6444
   End
   Begin Crystal.CrystalReport cr 
      Left            =   420
      Top             =   8040
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   495
      Left            =   9528
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   48
      Width           =   990
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   495
      Left            =   8448
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   48
      Width           =   1005
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   495
      Left            =   7368
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   48
      Width           =   1035
   End
   Begin VB.ComboBox cboBook 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmPendingBill.frx":000C
      Left            =   1512
      List            =   "frmPendingBill.frx":0016
      TabIndex        =   0
      Top             =   144
      Width           =   2496
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5970
      Left            =   180
      TabIndex        =   5
      Top             =   1800
      Width           =   11085
      _cx             =   19553
      _cy             =   10530
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
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
      Rows            =   500
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPendingBill.frx":0029
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
      Editable        =   0
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
   Begin MSComCtl2.DTPicker todate 
      Height          =   312
      Left            =   5844
      TabIndex        =   2
      Top             =   120
      Width           =   1392
      _ExtentX        =   2455
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155189249
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   312
      Left            =   4044
      TabIndex        =   1
      Top             =   120
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   550
      _Version        =   393216
      Format          =   155189249
      CurrentDate     =   39545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Group :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   252
      Left            =   5544
      TabIndex        =   6
      Top             =   180
      Width           =   432
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Book Type  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   144
      TabIndex        =   4
      Top             =   180
      Width           =   1512
   End
End
Attribute VB_Name = "frmPendingBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboGp_Click()

CboPName.clear
If rs.State = 1 Then rs.Close
rs.Open "select ItemName from  ItemMaster where ItemGp='" & cboGp.text & "'", con
While rs.EOF = False
CboPName.AddItem rs(0)
rs.MoveNext
Wend

End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub

Private Sub CmdPrint_Click()

Dim str1 As String

str1 = ""

If (cboBook.text <> "") Then
   str1 = "{pendingbill.type}='" & cboBook.text & "'"
End If


If str1 <> "" Then
    If (cboGp.text <> "") Then
      str1 = str1 & " and {pendingbill.bgp}='" & cboGp.text & "'"
    End If
Else
     If (cboGp.text <> "") Then
       str1 = "{pendingbill.bgp}='" & cboGp.text & "'"
     End If
End If


If str1 <> "" Then

    If (CboPName.text <> "") Then
            str1 = str1 & " and {pendingbill.BOOKName}='" & CboPName.text & "'"
    End If

Else
    If (CboPName.text <> "") Then
      str1 = "{pendingbill.BOOKName}='" & CboPName.text & "'"
    End If
End If


If str1 <> "" Then
 str1 = str1 & " and {pendingbill.EntDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {pendingbill.entDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
Else
 str1 = "{pendingbill.EntDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {pendingbill.entDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
End If




cr.ReportFileName = App.Path & "\PendingBill.rpt"
 cr.WindowShowPrintSetupBtn = True
 
 cr.WindowState = crptMaximized
 cr.WindowShowPrintBtn = True
 cr.Formulas(0) = "fromDate='" & Me.fromDate.Value & "'"
 cr.Formulas(1) = "toDate='" & Me.todate.Value & "'"
' If (CboPName.text = "" And cboBook.text <> "" And cboGp <> "") Then
'   cr.ReplaceSelectionFormula "{pendingbill.type}='" & cboBook.text & "' and {pendingbill.bgp}='" & cboGp.text & "' and {pendingbill.EntDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {pendingbill.entDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
'  ElseIf (CboPName.text <> "" And cboBook.text = "" And cboGp <> "") Then
'    cr.ReplaceSelectionFormula "{pendingbill.type}='" & cboBook.text & "' and {pendingbill.bgp}='" & cboGp.text & "' and {pendingbill.EntDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {pendingbill.entDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
' End If
If str1 <> "" Then
 cr.ReplaceSelectionFormula str1
End If
cr.Action = 1
 
 
End Sub
Private Sub cmdView_Click()

vs.clear
vs.Cols = 4
vs.Rows = 2

Dim str1 As String
str1 = ""

str1 = "entdate>=datevalue('" & fromDate.Value & "') and entdate<=datevalue('" & todate.Value & "')"
'==============
If rs.State = 1 Then rs.Close
rs.Open "select  ItemName,Itemgp from ItemMaster  where Itemgp='" & cboGp.text & "' order by ItemName", con
While rs.EOF = False
     con.Execute "update ITC set bgp='" & rs!itemgp & "' where BookName= '" & rs!itemname & "'"
     con.Execute "update BookReceive set bgp= '" & rs!itemgp & "' where BookName= '" & rs!itemname & "'"
rs.MoveNext
Wend






If CboPName.text <> "" Then
   str1 = str1 & " and BOOKName='" & CboPName.text & "'"
End If

If (cboBook.text <> "") Then
   str1 = str1 & " and type='" & cboBook.text & "'"
End If

If (cboGp.text <> "") Then
   str1 = str1 & " and bgp='" & cboGp.text & "'"
End If



  
If rs.State = 1 Then rs.Close
'If (CboPName.Text <> "" And cboBook.Text <> "") Then
'   rs.Open "select EntDate,BOOKName,qty,status from PendingBill  where BOOKName='" & CboPName.Text & "' and type='" & cboBook.Text & "' and entdate>=datevalue('" & fromDate.Value & "') and entdate<=datevalue('" & todate.Value & "') order by BOOKName,entdate", con
'ElseIf (cboBook.Text <> "") Then
'   rs.Open "select EntDate,BOOKName,qty,status from PendingBill  where type='" & cboBook.Text & "' and entdate>=datevalue('" & fromDate.Value & "') and entdate<=datevalue('" & todate.Value & "') order by BOOKName,entdate", con
'Else
   rs.Open "select EntDate,BOOKName,qty,status from PendingBill  where " & str1 & " order by BOOKName,entdate", con
'End If



For k1 = 1 To rs.RecordCount
 vs.Rows = vs.Rows + 1
 vs.TextMatrix(k1, 0) = rs!entdate
 vs.TextMatrix(k1, 1) = rs!BookName
 If rs!Status = "R" Then
    vs.TextMatrix(k1, 2) = rs!qty
 Else
    vs.TextMatrix(k1, 3) = rs!qty
 End If

 
 rs.MoveNext

Next





vs.FormatString = "Date|Book Name|Receive Qty|Billing Qty"

vs.ColWidth(0) = 1200
vs.ColWidth(1) = 5800
vs.ColWidth(2) = 1500
vs.ColWidth(3) = 1500




End Sub
Private Sub Form_Load()
  cboBook.ListIndex = 0
     
  fromDate.Value = Date
  todate.Value = Date
     
  fillcombo cboGp, "Name", "ConsumeGp", con
End Sub

