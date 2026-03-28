VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBookRecIssue 
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   12456
   Icon            =   "frmBookRecIssue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   12456
   WindowState     =   2  'Maximized
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
      Height          =   315
      Left            =   7080
      TabIndex        =   0
      Top             =   840
      Width           =   4305
   End
   Begin VB.ComboBox cboCustomer 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   2
      Top             =   1260
      Width           =   4305
   End
   Begin VB.TextBox txtQty 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1950
      MaxLength       =   8
      TabIndex        =   5
      Top             =   2520
      Width           =   1155
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
      Height          =   315
      Left            =   1950
      TabIndex        =   3
      Top             =   1680
      Width           =   5145
   End
   Begin VB.TextBox txtChallan 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1950
      TabIndex        =   4
      Top             =   2040
      Width           =   1635
   End
   Begin VB.TextBox txtEntryNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1950
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   825
      Left            =   7020
      TabIndex        =   8
      Top             =   2640
      Width           =   5265
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
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   555
         Left            =   2550
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   555
         Left            =   3855
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   165
         Width           =   1290
      End
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   555
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.TextBox txtdesc 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1950
      TabIndex        =   6
      Top             =   3000
      Width           =   5025
   End
   Begin MSComCtl2.DTPicker EnDate 
      Height          =   315
      Left            =   1950
      TabIndex        =   1
      Top             =   1260
      Width           =   1635
      _ExtentX        =   2879
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   139526145
      CurrentDate     =   40805
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5370
      Left            =   60
      TabIndex        =   13
      Top             =   3600
      Width           =   12165
      _cx             =   21458
      _cy             =   9472
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
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBookRecIssue.frx":0442
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
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
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
      Left            =   5475
      TabIndex        =   23
      Top             =   840
      Width           =   1620
   End
   Begin VB.Label lblGp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7095
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label lblHead 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   60
      Width           =   12195
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Rec. Date :"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1260
      Width           =   1740
   End
   Begin VB.Label lblBinder 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Press Name :"
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
      Left            =   5475
      TabIndex        =   19
      Top             =   1260
      Width           =   1620
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Qty  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Book Name  :"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   1740
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Challan No :"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2100
      Width           =   1740
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "St. No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1740
   End
End
Attribute VB_Name = "frmBookRecIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBook_Click()
 
 If rs.State = 1 Then rs.Close
 rs.Open "select ItemGp from ItemMaster where itemname='" & cboBook.text & "'", con
 If rs.EOF = False Then
    lblGp.Caption = rs(0)
    fillgrid
 End If
 
End Sub

Private Sub cboBook_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtChallan.SetFocus
End Sub

Private Sub cboCustomer_Click()
fillgrid
End Sub

Private Sub cboCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboBook.SetFocus
End Sub


Private Sub cboNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboTypeMode.SetFocus
End Sub

Private Sub cboTypeMode_Click()

If rs.State = 1 Then rs.Close
rs.Open "select rate from unitmaster where name='" & cboTypeMode.text & "'", con
If rs.EOF = False Then
  txtRate.text = Format(rs(0), ".00")
End If

End Sub
Private Sub cboTypeMode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then txtdesc.SetFocus

End Sub
Private Sub cboGp_Click()

cboBook.clear

If rs.State = 1 Then rs.Close
rs.Open "select ItemName from  ItemMaster where ItemGp='" & cbogp.text & "' order by ItemName", con
While rs.EOF = False
cboBook.AddItem rs(0)
rs.MoveNext
Wend


End Sub
Private Sub cbogp_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then EnDate.SetFocus
End Sub

Private Sub cbogp_LostFocus()
Call cboGp_Click
End Sub

Private Sub cmdDel_Click()

If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from BookReceive where EntryNo =" & txtEntryNo.text & ""
   Call cmdRef_Click
End If

End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub
Private Sub cmdRef_Click()

txtEntryNo.text = MaxSNo("BookReceive", "EntryNo")

'cboCustomer.Text = ""

If ss = "R" Then
   cboBook.text = ""
End If

txtChallan = ""

txtQty = ""
txtRate = ""
txtdesc.text = ""
lblGp.Caption = ""


EnDate.SetFocus

fillgrid


End Sub
Private Sub cmdSave_Click()






If ss = "R" Then

    If cboCustomer.text = "" Then
        MsgBox "Select Binder ...", vbCritical
        cboCustomer.SetFocus
        Exit Sub
    End If

'Else
'
'If rs.State = 1 Then rs.Close
'rs.Open "select * from BookReceive  where ChallanNo ='" & txtChallan & "'", con
'If rs.EOF = False Then
'   MsgBox "This challan No Already Exist ...", vbCritical
'   Exit Sub
'End If
'

End If



If cboBook.text = "" Then
    MsgBox "Select Book ...", vbCritical
    cboBook.SetFocus
    Exit Sub
End If


'If txtChallan.Text = "" Then
'    MsgBox "Enter Challan No ...", vbCritical
'    txtChallan.SetFocus
'    Exit Sub
'End If





If rs.State = 1 Then rs.Close
rs.Open "select * from BookReceive where EntryNo=" & txtEntryNo.text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then
   rs.AddNew
End If


rs!EntryNo = txtEntryNo.text
rs!EntDate = EnDate.Value

If ss = "R" Then
rs!PartyName = cboCustomer.text
Else
rs!PartyName = "Z"
End If

rs!qty = Val(txtQty)
rs!BookName = cboBook.text
rs!ChallanNo = txtChallan.text
rs!descr = txtdesc.text

If ss = "R" Then
  rs!Issue_Rec = "R"
Else
  rs!Issue_Rec = "D"
End If

rs.Update

Call cmdRef_Click




End Sub

Private Sub EnDate_KeyDown(KeyCode As Integer, Shift As Integer)
If cboCustomer.Visible = True Then
 If KeyCode = 13 Then cboCustomer.SetFocus
End If


If ss = "D" Then
 If KeyCode = 13 Then cboBook.SetFocus
End If


End Sub
Sub fillgrid()

Dim f As New ADODB.Recordset
  
i = 1

vs.clear


  
If f.State = 1 Then f.Close

If ss = "R" Then

    If cboBook.text = "" Then
       f.Open "select * from BookReceive where Issue_Rec='R' order by EntryNo", con, adOpenDynamic, adLockOptimistic
    Else
       f.Open "select * from BookReceive where BookName ='" & cboBook.text & "' and Issue_Rec='R' order by EntryNo", con, adOpenDynamic, adLockOptimistic
    End If


'================================================



If f.EOF = False Then
  vs.Rows = f.RecordCount + 1
End If

While f.EOF = False

vs.TextMatrix(i, 0) = f!EntryNo
vs.TextMatrix(i, 1) = f!EntDate
vs.TextMatrix(i, 2) = f!PartyName
vs.TextMatrix(i, 3) = f!BookName
vs.TextMatrix(i, 4) = f!qty
vs.TextMatrix(i, 5) = f!ChallanNo
vs.TextMatrix(i, 6) = f!descr


i = i + 1

f.MoveNext
Wend


vs.Cols = 7

vs.FormatString = "StNo|Date|PressName|<BookName|Qty|ChallanNo|Description"

vs.ColWidth(0) = 800
vs.ColWidth(1) = 1200
vs.ColWidth(2) = 2400
vs.ColWidth(3) = 3500
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 1000
vs.ColWidth(6) = 2000



'================================================

Else

    If cboBook = "" Then
       f.Open "select * from BookReceive where Issue_Rec='D' order by EntryNo", con, adOpenDynamic, adLockOptimistic
    Else
       f.Open "select * from BookReceive where BookName ='" & cboBook.text & "' and Issue_Rec='D' order by EntryNo", con, adOpenDynamic, adLockOptimistic
    End If


'================================================
'vs.Rows = 1

If f.EOF = False Then
vs.Rows = f.RecordCount + 1
End If


While f.EOF = False



vs.TextMatrix(i, 0) = f!EntryNo
vs.TextMatrix(i, 1) = f!EntDate
vs.TextMatrix(i, 2) = f!BookName
vs.TextMatrix(i, 3) = f!qty
vs.TextMatrix(i, 4) = f!ChallanNo
vs.TextMatrix(i, 5) = f!descr


i = i + 1

f.MoveNext
Wend


vs.Cols = 6

vs.FormatString = "StNo|RecDate|BookName|Qty|ChallanNo|Description"

vs.ColWidth(0) = 800
vs.ColWidth(1) = 1200
vs.ColWidth(2) = 4200
vs.ColWidth(3) = 1500
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 1000
'vs.ColWidth(6) = 3000




'================================================


End If










End Sub
Private Sub Form_Load()
pos Me

EnDate.Value = Date


fillcombo cbogp, "Name", "ConsumeGp", con


fillcombo cboCustomer, "Name", "Supplier", con

fillcombo cboBook, "ItemName", "ItemMaster", con

'Call cmdRef_Click


txtEntryNo.text = MaxSNo("BookReceive", "EntryNo")

cboCustomer.text = ""

txtQty = ""
txtRate = ""

cboBook.text = ""
fillgrid


If ss = "R" Then
 lblHead = "BOOK RECEIVED FROM PRESS : "
 lblBinder.Visible = True
 cboCustomer.Visible = True
 Label1.Caption = "Rec. Date :"
Else
 lblHead = "BOOK DELIVERED : "
 lblBinder.Visible = False
 cboCustomer.Visible = False
 Label1.Caption = "DEL. Date :"
End If


lblGp.Caption = ""

End Sub

Private Sub txtChallan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtQty.SetFocus
End Sub

Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtdesc.SetFocus
End Sub
Private Sub vs_Click()

On Error Resume Next

Dim f As New ADODB.Recordset

If vs.Rows <= 1 Then Exit Sub

If vs.TextMatrix(vs.RowSel, 0) = "" Then
   Exit Sub
End If
  
If f.State = 1 Then f.Close
f.Open "select * from BookReceive where EntryNo =" & vs.TextMatrix(vs.RowSel, 0) & "", con, adOpenDynamic, adLockOptimistic
If f.EOF = False Then
cmdDel.Enabled = True

txtEntryNo.text = f!EntryNo
EnDate.Value = f!EntDate
cboCustomer.text = f!PartyName
txtQty.text = f!qty

cboBook.text = f!BookName

txtChallan.text = f!ChallanNo & ""

txtdesc.text = f!descr & ""


End If


 If rs.State = 1 Then rs.Close
 rs.Open "select ItemGp from ItemMaster where itemname='" & cboBook.text & "'", con
 If rs.EOF = False Then
    'lblGp.Caption = rs(0)
    cbogp.text = rs(0)
 End If



End Sub

