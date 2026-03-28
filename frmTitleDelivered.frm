VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTitleDelivered 
   Caption         =   "Title Delivered"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14610
   Icon            =   "frmTitleDelivered.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   14610
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txtdesc 
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
      Left            =   1905
      TabIndex        =   7
      Top             =   2880
      Width           =   4245
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   825
      Left            =   7020
      TabIndex        =   12
      Top             =   2520
      Width           =   5220
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   555
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   555
         Left            =   3855
         Style           =   1  'Graphical
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   555
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   8
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
      Left            =   1905
      TabIndex        =   11
      Top             =   720
      Width           =   1710
   End
   Begin VB.TextBox txtChallan 
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
      Left            =   1905
      TabIndex        =   5
      Top             =   1920
      Width           =   1860
   End
   Begin VB.ComboBox cboBook 
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
      Left            =   1905
      TabIndex        =   4
      Top             =   1560
      Width           =   4260
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
      Height          =   315
      Left            =   1905
      MaxLength       =   8
      TabIndex        =   6
      Top             =   2400
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
      Left            =   6180
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox cbogp 
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
      Left            =   6180
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   9780
      TabIndex        =   9
      Top             =   660
      Width           =   2535
      Begin VB.ComboBox cboInner_Cover 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmTitleDelivered.frx":000C
         Left            =   60
         List            =   "frmTitleDelivered.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   2355
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Inner/Cover/Paper :"
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
         Left            =   60
         TabIndex        =   10
         Top             =   120
         Width           =   2355
      End
   End
   Begin MSComCtl2.DTPicker EnDate 
      Height          =   315
      Left            =   1905
      TabIndex        =   1
      Top             =   1140
      Width           =   1770
      _ExtentX        =   3122
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
      Format          =   64749569
      CurrentDate     =   40805
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5490
      Left            =   60
      TabIndex        =   16
      Top             =   3480
      Width           =   12165
      _cx             =   21458
      _cy             =   9684
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
      Rows            =   500
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTitleDelivered.frx":0042
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
      Caption         =   "Delivred To :"
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
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "St. No :"
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
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Challan No :"
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
      Left            =   120
      TabIndex        =   24
      Top             =   1980
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Book Name  :"
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
      Left            =   120
      TabIndex        =   23
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Qty  :"
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
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblBinder 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Press Name :"
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
      Left            =   4440
      TabIndex        =   21
      Top             =   1140
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Date :"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1140
      Width           =   1695
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
      Height          =   555
      Left            =   120
      TabIndex        =   19
      Top             =   60
      Width           =   12195
   End
   Begin VB.Label lblGp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6195
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Group :"
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
      Left            =   4440
      TabIndex        =   17
      Top             =   720
      Width           =   1710
   End
End
Attribute VB_Name = "frmTitleDelivered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBook_Click()
 
 If rs.State = 1 Then rs.Close
 rs.Open "select ItemGp from ItemMaster where itemname='" & cboBook.Text & "'", con
 If rs.EOF = False Then
    lblGp.Caption = rs(0)
    fillgrid
 End If
 
End Sub

Private Sub cboBook_KeyDown(KeyCode As Integer, Shift As Integer)

If cboBook.Text <> "" Then
   If KeyCode = 13 Then cboInner_Cover.SetFocus
End If

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
rs.Open "select rate from unitmaster where name='" & cboTypeMode.Text & "'", con
If rs.EOF = False Then
  txtRate.Text = Format(rs(0), ".00")
End If

End Sub
Private Sub cboTypeMode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then txtdesc.SetFocus

End Sub
Private Sub cboGp_Click()

cboBook.clear

If rs.State = 1 Then rs.Close
rs.Open "select ItemName from  ItemMaster where ItemGp='" & cbogp.Text & "' order by ItemName", con
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

Private Sub cboInner_Cover_KeyDown(KeyCode As Integer, Shift As Integer)
If cboInner_Cover.Text <> "" Then
   If KeyCode = 13 Then txtChallan.SetFocus
End If
End Sub

Private Sub cmdDel_Click()

If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from TitleDel where EntryNo =" & txtEntryNo.Text & ""
   Call cmdRef_Click
End If

End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub
Private Sub cmdRef_Click()

txtEntryNo.Text = MaxSNo("TitleDel", "EntryNo")

cboCustomer.Text = ""
'cboBook.Text = ""

txtChallan = ""

txtQty = ""
txtRate = ""
txtdesc.Text = ""
lblGp.Caption = ""
cboCustomer.Text = "-"


EnDate.SetFocus

fillgrid


End Sub
Private Sub cmdSave_Click()






If ss = "R" Then
''
''    If cboCustomer.Text = "" Then
''        MsgBox "Select Binder ...", vbCritical
''        cboCustomer.SetFocus
''        Exit Sub
''    End If

'Else
'
'If rs.State = 1 Then rs.Close
'rs.Open "select * from TitleDel  where ChallanNo ='" & txtChallan & "'", con
'If rs.EOF = False Then
'   MsgBox "This challan No Already Exist ...", vbCritical
'   Exit Sub
'End If
'

End If



If cboBook.Text = "" Then
    MsgBox "Select Book ...", vbCritical
    cboBook.SetFocus
    Exit Sub
End If


If txtdesc.Text = "" Then
    MsgBox "Select Delivered to...", vbCritical
    txtdesc.SetFocus
    Exit Sub
End If


If cboInner_Cover.Text = "" Then
    MsgBox "Select Inner/Cover/Paper...", vbCritical
    cboInner_Cover.SetFocus
    Exit Sub
End If





If rs.State = 1 Then rs.Close
rs.Open "select * from TitleDel where EntryNo=" & txtEntryNo.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then
   rs.AddNew
End If


rs!EntryNo = txtEntryNo.Text
rs!EntDate = EnDate.Value

'If ss = "R" Then
'rs!PartyName = cboCustomer.Text
'Else
'rs!PartyName = "Z"
'End If

rs!qty = Val(txtQty)
rs!BookName = cboBook.Text
rs!ChallanNo = txtChallan.Text
rs!descr = txtdesc.Text

If ss = "R" Then
  rs!Issue_Rec = "R"
Else
  rs!Issue_Rec = "D"
End If

rs!inner_paper_cover = cboInner_Cover.Text

rs.Update

Call cmdRef_Click




End Sub

Private Sub EnDate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then cboBook.SetFocus
End Sub
Sub fillgrid()

Dim f As New ADODB.Recordset
  
i = 1

vs.clear


  
If f.State = 1 Then f.Close

'If ss = "R" Then

    If cboBook.Text = "" Then
       'f.Open "select * from TitleDel where Issue_Rec='R' order by EntryNo", con, adOpenDynamic, adLockOptimistic
       f.Open "select * from TitleDel  order by EntryNo", con, adOpenDynamic, adLockOptimistic
    Else
       'f.Open "select * from TitleDel where Issue_Rec='R' and bookname='" & cboBook.Text & "' order by EntryNo", con, adOpenDynamic, adLockOptimistic
       f.Open "select * from TitleDel where  bookname='" & cboBook.Text & "' order by EntryNo", con, adOpenDynamic, adLockOptimistic
    End If


'================================================

If f.EOF = False Then
   vs.Rows = f.RecordCount + 1
End If


While f.EOF = False

vs.TextMatrix(i, 0) = f!EntryNo
vs.TextMatrix(i, 1) = f!EntDate
'vs.TextMatrix(i, 2) = f!PartyName
vs.TextMatrix(i, 2) = f!BookName
vs.TextMatrix(i, 3) = f!qty
vs.TextMatrix(i, 4) = f!ChallanNo & ""
vs.TextMatrix(i, 5) = f!descr


i = i + 1

f.MoveNext
Wend


vs.Cols = 6

vs.FormatString = "StNo|RecDate|<BookName|Qty|ChallanNo|Description"

vs.ColWidth(0) = 1000
vs.ColWidth(1) = 1500
vs.ColWidth(2) = 3000
vs.ColWidth(3) = 1500
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 1000
'vs.ColWidth(6) = 3000



'''================================================
''
''Else
''
''    If cboCustomer.Text = "" Then
''       f.Open "select * from TitleDel where Issue_Rec='D' order by EntryNo", con, adOpenDynamic, adLockOptimistic
''    Else
''       f.Open "select * from TitleDel where bookname ='" & cboBook & "' and Issue_Rec='D' order by EntryNo", con, adOpenDynamic, adLockOptimistic
''    End If
''
''
'''================================================
''
''While f.EOF = False
''
''vs.TextMatrix(i, 0) = f!EntryNo
''vs.TextMatrix(i, 1) = f!EntDate
''vs.TextMatrix(i, 2) = f!BookName
''vs.TextMatrix(i, 3) = f!qty
''vs.TextMatrix(i, 4) = f!ChallanNo
''vs.TextMatrix(i, 5) = f!descr
''
''
''i = i + 1
''
''f.MoveNext
''Wend
''
''
''vs.Cols = 6
''
''vs.FormatString = "StNo|RecDate|BookName|Qty|ChallanNo|Description"
''
''vs.ColWidth(0) = 1000
''vs.ColWidth(1) = 1500
''vs.ColWidth(2) = 3000
''vs.ColWidth(3) = 1500
''vs.ColWidth(4) = 1000
''vs.ColWidth(5) = 1000
'''vs.ColWidth(6) = 3000
''
''
''
''
'''================================================
''
''
''End If










End Sub
Private Sub Form_Load()
pos Me

EnDate.Value = Date


fillcombo cbogp, "Name", "ConsumeGp", con


fillcombo cboCustomer, "Name", "Supplier", con

fillcombo cboBook, "ItemName", "ItemMaster", con

fillcombo txtdesc, "Name", "Customer", con

'Call cmdRef_Click


txtEntryNo.Text = MaxSNo("TitleDel", "EntryNo")

cboCustomer.Text = ""

txtQty = ""
txtRate = ""

cboBook.Text = ""
fillgrid


'If ss = "R" Then
lblHead = "TitleDel/INNER/PAPER  : Delivered "
''lblBinder.Visible = True
''cboCustomer.Visible = True
'Else
'lblHead = "TitleDel CONSUMED : "
''
''lblBinder.Visible = False
''cboCustomer.Visible = False
''
'End If
''


'If rs.State = 1 Then rs.Close
'rs.Open "select name from "


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
  
If vs.TextMatrix(vs.RowSel, 0) = "" Then
   Exit Sub
End If
  
If f.State = 1 Then f.Close
f.Open "select * from TitleDel where EntryNo =" & vs.TextMatrix(vs.RowSel, 0) & "", con, adOpenDynamic, adLockOptimistic
If f.EOF = False Then
cmdDel.Enabled = True

txtEntryNo.Text = f!EntryNo
EnDate.Value = f!EntDate
'cboCustomer.Text = f!PartyName
txtQty.Text = f!qty

cboBook.Text = f!BookName

txtChallan.Text = f!ChallanNo & ""

txtdesc.Text = f!descr & ""


cboInner_Cover.Text = f!inner_paper_cover & ""



End If


 If rs.State = 1 Then rs.Close
 rs.Open "select ItemGp from ItemMaster where itemname='" & cboBook.Text & "'", con
 If rs.EOF = False Then
    'lblGp.Caption = rs(0)
    cbogp.Text = rs(0)
 End If



End Sub




