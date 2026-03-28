VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFolding 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Customer Wise Folding Entry"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15165
   Icon            =   "frmFolding.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   15165
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtdesc 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox txtTAmt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15300
      TabIndex        =   21
      Top             =   8220
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   825
      Left            =   180
      TabIndex        =   17
      Top             =   3420
      Width           =   5265
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   555
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   555
         Left            =   3855
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   165
         Width           =   1290
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   555
         Left            =   2550
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   165
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   555
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   165
         Width           =   1245
      End
   End
   Begin VB.TextBox txtEntryNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtRate 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2040
      Width           =   1320
   End
   Begin VB.ComboBox cboTypeMode 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   2100
      Width           =   1425
   End
   Begin VB.ComboBox cboNumber 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmFolding.frx":000C
      Left            =   4920
      List            =   "frmFolding.frx":0019
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtQty 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   1380
   End
   Begin VB.ComboBox cboCustomer 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   1140
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker EnDate 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   660
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   71041025
      CurrentDate     =   40805
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7890
      Left            =   6300
      TabIndex        =   16
      Top             =   120
      Width           =   11085
      _cx             =   19553
      _cy             =   13917
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
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmFolding.frx":0026
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
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   23
      Top             =   2520
      Width           =   1920
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Total Amount(Rs.):"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12960
      TabIndex        =   22
      Top             =   8220
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Entry No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   15
      Top             =   180
      Width           =   1920
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Rate :"
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
      Left            =   3480
      TabIndex        =   14
      Top             =   2100
      Width           =   1395
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Type Of Fold :"
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
      Left            =   75
      TabIndex        =   13
      Top             =   2100
      Width           =   1920
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "No. Of Type :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Qty :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   11
      Top             =   1560
      Width           =   1920
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Customer Name :"
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
      Left            =   75
      TabIndex        =   10
      Top             =   1140
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Entry Date :"
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
      Left            =   75
      TabIndex        =   9
      Top             =   660
      Width           =   1920
   End
End
Attribute VB_Name = "frmFolding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboCustomer_Click()
fillgrid
End Sub

Private Sub cboCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtQty.SetFocus
End Sub


Private Sub cboNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboTypeMode.SetFocus
End Sub

Private Sub cboTypeMode_Click()

If rs.State = 1 Then rs.Close
rs.Open "select rate from unitmaster where name='" & cboTypeMode.Text & "'", con
If rs.EOF = False Then
  txtRate.Text = Format(rs(0), ".00")
End If

End Sub
Private Sub cboTypeMode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then txtdesc.SetFocus

End Sub

Private Sub cboTypeMode_LostFocus()
If rs.State = 1 Then rs.Close
rs.Open "select rate from unitmaster where name='" & cboTypeMode.Text & "'", con
If rs.EOF = False Then
  txtRate.Text = Format(rs(0), ".00")
End If

End Sub

Private Sub cmdDel_Click()

If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
con.Execute "delete from FoldingDetails where EntryNo =" & vs.TextMatrix(vs.RowSel, 0) & ""
Call cmdRef_Click
End If

End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub
Private Sub cmdRef_Click()

txtEntryNo.Text = MaxSNo("FoldingDetails", "EntryNo")

cboCustomer.Text = ""
cboNumber.Text = ""
cboTypeMode.Text = ""

txtQty = ""
txtRate = ""
txtdesc.Text = ""

EnDate.SetFocus

fillgrid


End Sub
Private Sub cmdSave_Click()



If cboCustomer.Text = "" Then
    MsgBox "Select Customer ...", vbCritical
    cboCustomer.SetFocus
    Exit Sub
End If



If cboNumber.Text = "" Then
    MsgBox "Select Number of Type ...", vbCritical
    cboNumber.SetFocus
    Exit Sub
End If


If cboTypeMode.Text = "" Then
    MsgBox "Select Type Of Fold ...", vbCritical
    cboTypeMode.SetFocus
    Exit Sub
End If





If rs.State = 1 Then rs.Close
rs.Open "select * from FoldingDetails where EntryNo=" & txtEntryNo.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then
   rs.AddNew
End If


rs!EntryNo = txtEntryNo.Text
rs!EntDate = EnDate.Value
rs!PartyName = cboCustomer.Text
rs!Qty = Val(txtQty)
rs!NoOfType = Val(cboNumber.Text)
rs!TypeOfFold = Trim(cboTypeMode.Text)
rs!Rate = Val(txtRate)
rs!Amount = ((Val(txtQty) * Val(cboNumber.Text) * Val(txtRate)) / 1000)
rs!descr = txtdesc.Text

rs.Update

Call cmdRef_Click




End Sub

Private Sub EnDate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then cboCustomer.SetFocus
End Sub
Sub fillgrid()

Dim f As New ADODB.Recordset
  
i = 1

vs.clear

txtTAmt.Text = 0

vs.Cols = 9
  
If f.State = 1 Then f.Close

If cboCustomer.Text = "" Then
   f.Open "select * from FoldingDetails order by EntryNo", con, adOpenDynamic, adLockOptimistic
Else
   f.Open "select * from FoldingDetails where PartyName ='" & cboCustomer.Text & "' order by EntryNo", con, adOpenDynamic, adLockOptimistic
End If

If f.EOF = False Then
vs.Rows = f.RecordCount + 1
End If

While f.EOF = False


vs.TextMatrix(i, 0) = f!EntryNo
vs.TextMatrix(i, 1) = f!EntDate
vs.TextMatrix(i, 2) = f!PartyName
vs.TextMatrix(i, 3) = f!Qty
vs.TextMatrix(i, 4) = f!NoOfType
vs.TextMatrix(i, 5) = f!TypeOfFold
vs.TextMatrix(i, 6) = f!Rate
vs.TextMatrix(i, 7) = f!Amount

vs.TextMatrix(i, 8) = f!descr

txtTAmt.Text = Format(Val(txtTAmt.Text) + f!Amount, "0.00")


i = i + 1

f.MoveNext
Wend




vs.FormatString = "EntryNo|EntDate|PartyName|Qty|NoOfType|TypeOfFold|Rate|Amount|Description"

vs.ColWidth(0) = 900
vs.ColWidth(1) = 1000
vs.ColWidth(2) = 2500
vs.ColWidth(3) = 1000
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 1000
vs.ColWidth(6) = 1000
vs.ColWidth(7) = 1200
vs.ColWidth(8) = 1500








End Sub
Private Sub Form_Load()
pos Me

EnDate.Value = Date
cboNumber.ListIndex = 0


fillcombo cboCustomer, "Name", "Customer", con

fillcombo cboTypeMode, "Name", "UnitMaster", con

'Call cmdRef_Click


txtEntryNo.Text = MaxSNo("FoldingDetails", "EntryNo")

cboCustomer.Text = ""
cboNumber.Text = ""
cboTypeMode.Text = ""

txtQty = ""
txtRate = ""


fillgrid


End Sub
Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboNumber.SetFocus
End Sub
Private Sub vs_Click()


Dim f As New ADODB.Recordset
  
If vs.TextMatrix(vs.RowSel, 0) = "" Then
   Exit Sub
End If
  
If f.State = 1 Then f.Close
f.Open "select * from FoldingDetails where EntryNo =" & vs.TextMatrix(vs.RowSel, 0) & "", con, adOpenDynamic, adLockOptimistic
If f.EOF = False Then
cmdDel.Enabled = True

txtEntryNo.Text = f!EntryNo
EnDate.Value = f!EntDate
cboCustomer.Text = f!PartyName
txtQty.Text = f!Qty
cboNumber.Text = f!NoOfType
cboTypeMode.Text = f!TypeOfFold
txtRate.Text = f!Rate

txtdesc.Text = f!descr & ""


End If


End Sub
