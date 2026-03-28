VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTitleST 
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   8
      Top             =   6960
      Width           =   1275
   End
   Begin VB.TextBox txtrem 
      Height          =   345
      Left            =   1860
      TabIndex        =   5
      Top             =   5340
      Width           =   8175
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
      Left            =   3000
      TabIndex        =   1
      Top             =   1500
      Width           =   4710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5160
      TabIndex        =   11
      Top             =   8460
      Width           =   1635
   End
   Begin Crystal.CrystalReport cr 
      Left            =   0
      Top             =   6900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6900
      TabIndex        =   12
      Top             =   8460
      Width           =   1635
   End
   Begin VB.TextBox txtConsumed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   9
      Top             =   7380
      Width           =   1275
   End
   Begin VB.TextBox txtBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   25
      Top             =   8040
      Width           =   1275
   End
   Begin VB.TextBox txtITCBook 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   7
      Top             =   6600
      Width           =   1275
   End
   Begin VB.TextBox txtNewBook 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   6
      Top             =   6240
      Width           =   1275
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1620
      TabIndex        =   22
      Top             =   8460
      Width           =   1755
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8580
      TabIndex        =   21
      Top             =   8460
      Width           =   1635
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3480
      TabIndex        =   10
      Top             =   8460
      Width           =   1635
   End
   Begin VB.TextBox txtRecTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   19
      Top             =   5760
      Width           =   1275
   End
   Begin VB.TextBox txtSTNO 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Top             =   1020
      Width           =   1515
   End
   Begin VB.ComboBox cboBook 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   3
      Top             =   1980
      Width           =   7035
   End
   Begin MSMask.MaskEdBox txtdate 
      Height          =   315
      Left            =   8520
      TabIndex        =   2
      Top             =   1530
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtBinder 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6660
      TabIndex        =   15
      Text            =   "SUNIL BOOK BINDER"
      Top             =   1020
      Width           =   3735
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   2970
      Left            =   780
      TabIndex        =   4
      Top             =   2340
      Width           =   9285
      _cx             =   16378
      _cy             =   5239
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTitleST.frx":0000
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
   Begin MSMask.MaskEdBox txtCLDate 
      Height          =   315
      Left            =   10080
      TabIndex        =   31
      Top             =   1530
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Title  Delivered:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   34
      Top             =   7020
      Width           =   4575
   End
   Begin VB.Label Label14 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   780
      TabIndex        =   33
      Top             =   5340
      Width           =   1095
   End
   Begin VB.Label Label13 
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
      Left            =   900
      TabIndex        =   32
      Top             =   1500
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "F2 For Search"
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
      Left            =   3000
      TabIndex        =   30
      Top             =   660
      Width           =   3495
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0FF&
      Caption         =   " Total Consumed :"
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
      Left            =   3960
      TabIndex        =   29
      Top             =   7380
      Width           =   4575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      Height          =   135
      Left            =   840
      TabIndex        =   28
      Top             =   7800
      Width           =   9255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      Height          =   75
      Left            =   780
      TabIndex        =   27
      Top             =   6120
      Width           =   9255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Balance  :"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   8040
      Width           =   4575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ITC  Book  (Cover Change):"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   6660
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   " New Book  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   23
      Top             =   6300
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   " Total Receive :"
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
      Left            =   3960
      TabIndex        =   20
      Top             =   5760
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Statement No. :"
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
      Left            =   900
      TabIndex        =   18
      Top             =   1020
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Name Of  Book :"
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
      Left            =   900
      TabIndex        =   17
      Top             =   1980
      Width           =   2115
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Date :"
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
      Left            =   7740
      TabIndex        =   16
      Top             =   1530
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Name Of  Binder :"
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
      Left            =   4620
      TabIndex        =   14
      Top             =   1020
      Width           =   2175
   End
   Begin VB.Label lblHead 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "STATEMENT OF TITLE"
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
      Left            =   0
      TabIndex        =   13
      Top             =   60
      Width           =   19395
   End
End
Attribute VB_Name = "frmTitleST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub searchData()
   
  
txtNewBook = 0
txtITCBook = 0
  
refVs
  
Dim tot_Rec As Double
  
tot_Rec = 0
  
  
If rs.State = 1 Then rs.Close
rs.Open "select BalQty,Date1,dates,des1 from titleStatent where BookName='" & cboBook.Text & "' order by auto", con
If rs.EOF = False Then
    

rs.MoveLast
vs.TextMatrix(1, 0) = rs!Dates
vs.TextMatrix(1, 1) = "OPENING BAlANCE - " & rs!des1
vs.TextMatrix(1, 2) = rs(0)
txtCLDate.Text = rs!Dates


'================================== if statement is exist of thies Book-------------------------------------------
'================================== fatch Title Receive-----------------------------------------------------------


j = 2

If Not IsDate(txtdate.Text) Then Exit Sub

If rs.State = 1 Then rs.Close
rs.Open "select entDate,Descr,Qty from title where st='n' and bookname='" & cboBook & "' and inner_paper_cover='Cover' and entDate<=datevalue('" & txtdate.Text & "') order by EntDate", con
While rs.EOF = False

    vs.TextMatrix(j, 0) = rs(0)
    vs.TextMatrix(j, 1) = rs(1)
    vs.TextMatrix(j, 2) = rs(2)
    
    tot_Rec = tot_Rec + rs(2)
    
    j = j + 1
    rs.MoveNext
    
Wend

txtRecTotal = tot_Rec


'================================== New Book/ITC Book-----------------------------------------------------------


If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from BookReceive where st='n' and bookname='" & cboBook.Text & "' and Issue_Rec='R'", con
'rs.Open "select sum(Qty) from BookReceive where st='n' and bookname='" & cboBook.Text & "' and entDate=datevalue('" & txtdate.Text & "') and Issue_Rec='R'", con
'rs.Open "select sum(Qty) from BookReceive where st='n' and bookname='" & cboBook.Text & "' and entDate>=datevalue('" & txtdate.Text & "') and Issue_Rec='R'", con
If Not IsNull(rs(0)) Then
   txtNewBook = rs(0)
Else
   txtNewBook = 0
End If


If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from itc where st='n' and bookname='" & cboBook.Text & "' and Issue_Rec='R' and inner_paper_cover='Cover'", con
'rs.Open "select sum(Qty) from itc where st='n' and bookname='" & cboBook.Text & "' and entDate>=datevalue('" & txtdate.Text & "') and Issue_Rec='R' and inner_paper_cover='Cover'", con
'rs.Open "select sum(Qty) from itc where st='n' and bookname='" & cboBook.Text & "' and entDate=datevalue('" & txtdate.Text & "') and Issue_Rec='R' and inner_paper_cover='Cover'", con
If Not IsNull(rs(0)) Then
   txtITCBook = rs(0)
End If



If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from TitleDel where st='n' and bookname='" & cboBook.Text & "' and inner_paper_cover='Cover'", con
If Not IsNull(rs(0)) Then
   txtTitle = rs(0)
Else
   txtTitle = 0
End If





'================================== end-------------------------------------------------------------------------


  
Else


'================================== for First-------------------------------------------
'================================== fatch Title Receive---------------------------------
  
  
If rs.State = 1 Then rs.Close
rs.Open "select OpeningStock from ItemMaster where itemname='" & cboBook.Text & "'", con
If rs.EOF = False Then


j = 1

If rs.State = 1 Then rs.Close
rs.Open "select entDate,Descr,Qty from title where bookname='" & cboBook & "' and  inner_paper_cover='Cover' order by EntDate", con
While rs.EOF = False

vs.TextMatrix(j, 0) = rs(0)
vs.TextMatrix(j, 1) = rs(1)
vs.TextMatrix(j, 2) = rs(2)
j = j + 1

rs.MoveNext
Wend


End If

'================================== New Book/ITC Book-----------------------------------------------------------

If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from BookReceive where bookname='" & cboBook.Text & "' and Issue_Rec='R'", con
If Not IsNull(rs(0)) Then
   txtNewBook = rs(0)
Else
   txtNewBook = 0
End If



'--------------------------------------------------------------------------


If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from itc where bookname='" & cboBook.Text & "' and Issue_Rec='R' and inner_paper_cover='Cover'", con
If Not IsNull(rs(0)) Then
   txtITCBook = rs(0)
End If


If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from itc where bookname='" & cboBook.Text & "' and Issue_Rec='R' and inner_paper_cover='I+P+C'", con
If Not IsNull(rs(0)) Then
   txtITCBook = Val(txtITCBook) + rs(0)
End If


If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from itc where bookname='" & cboBook.Text & "' and Issue_Rec='R' and inner_paper_cover='I+C'", con
If Not IsNull(rs(0)) Then
   txtITCBook = Val(txtITCBook) + rs(0)
End If

If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from itc where bookname='" & cboBook.Text & "' and Issue_Rec='R' and inner_paper_cover='P+C'", con
If Not IsNull(rs(0)) Then
   txtITCBook = Val(txtITCBook) + rs(0)
End If


'================================== end-------------------------------------------------------------------------

If rs.State = 1 Then rs.Close
rs.Open "select sum(Qty) from TitleDel where bookname='" & cboBook.Text & "' and inner_paper_cover='Cover'", con
If Not IsNull(rs(0)) Then
   txtTitle = rs(0)
End If




End If





'===============================================================
'---------------------------------------------------------------

txtRecTotal = 0

For j = 1 To vs.Rows - 1
  If vs.TextMatrix(j, 0) <> "" Then
     txtRecTotal = Val(txtRecTotal) + Val(vs.TextMatrix(j, 2))
  End If
Next



txtConsumed = (Val(txtNewBook) + Val(txtITCBook) + Val(txtTitle))


txtBal = (Val(txtRecTotal) - Val(txtConsumed))


End Sub
Sub cal()
'===============================================================
'---------------------------------------------------------------

txtRecTotal = 0

For j = 1 To vs.Rows - 1
  If vs.TextMatrix(j, 2) <> "" Then
     txtRecTotal = Val(txtRecTotal) + Val(vs.TextMatrix(j, 2))
  End If
Next



txtConsumed = (Val(txtNewBook) + Val(txtITCBook) + Val(txtTitle))


txtBal = (Val(txtRecTotal) - Val(txtConsumed))

End Sub
Private Sub cboBook_Click()
searchData
End Sub

Private Sub cboBook_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    vs.SetFocus
 End If
End Sub

Private Sub cboBook_LostFocus()
searchData
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
If KeyCode = 13 Then
   txtdate.SetFocus
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Sub refVs()

vs.clear
vs.FormatString = "Date|Description|Receive"

vs.ColWidth(0) = 1500
vs.ColWidth(1) = 6200
vs.ColWidth(2) = 1270
'vs.ColWidth(3) = 1500


End Sub

Private Sub CmdPrint_Click()

s11 = ""

If MsgBox("Want To Print", vbQuestion + vbYesNo) = vbYes Then
   
   ss1 = ""
   cr.Reset
   cr.ReportFileName = App.Path & "\STOFTITLE.rpt"
   cr.ReplaceSelectionFormula "{titleStatent.STNo}=" & txtSTNO.Text & ""
   cr.Formulas(0) = "firm_='" & COMPNAME & "'"
   cr.Formulas(1) = "binder='" & txtBinder & "'"
   
   
   If rs.State = 1 Then rs.Close
   rs.Open "select * from TitleDel where STNo='" & txtSTNO.Text & "'", con
   While rs.EOF = False
    
    If ss1 = "" Then
       ss1 = "ChallanNo:" & rs!ChallanNo & "," & "Qty:" & rs!qty & ",Delivered To:" & rs!descr
    Else
       ss1 = ss1 & " ; " & "ChallanNo:" & rs!ChallanNo & "," & "Qty:" & rs!qty & ",Delivered To:" & rs!descr
    End If
   
   rs.MoveNext
   Wend
   cr.Formulas(2) = "desc_title='" & ss1 & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1


End If

End Sub

Private Sub cmdRef_Click()
addSt
cmdSave.Enabled = True
txtSTNO = MaxSNo("titleStatent", "stno")

End Sub

Private Sub cmdSave_Click()



If rs.State = 1 Then rs.Close
rs.Open "select * from titleStatent where STNo=" & txtSTNO & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
 con.Execute "delete from titleStatent where STNo=" & txtSTNO & ""
End If


con.Execute "update ITC set st='y',stno=" & txtSTNO & "   where BookName='" & cboBook.Text & "' and st='n'"
con.Execute "update BookReceive set st='y',stno=" & txtSTNO & "   where BookName='" & cboBook.Text & "' and st='n'"
con.Execute "update title set st='y',stno=" & txtSTNO & "   where BookName='" & cboBook.Text & "' and st='n'"
con.Execute "update titleDel set st='y',stno=" & txtSTNO & "   where BookName='" & cboBook.Text & "' and st='n'"

For i = 1 To vs.Rows - 1

If vs.TextMatrix(i, 1) <> "" Then

rs.AddNew
rs!STNo = txtSTNO.Text
rs!Dates = txtdate.Text
rs!BookName = cboBook.Text
rs!date1 = vs.TextMatrix(i, 0)
rs!descr = vs.TextMatrix(i, 1)
rs!RecQty = vs.TextMatrix(i, 2)
rs!NewBookQty = Val(txtNewBook)
rs!ITCBookQty = Val(txtITCBook)
rs!TitleBookQty = Val(txtTitle.Text)
rs!BalQty = Val(txtBal)
rs!des1 = Trim(txtrem)
rs!binder = Trim(txtBinder)
rs.Update

End If

Next

'cmdSave.Enabled = False
MsgBox "Saved ....", vbInformation


End Sub

Sub searchData1()

Dim totRec, totConsume As Double

Dim rs1 As New ADODB.Recordset

totRec = 0
totConsume = 0

If rs.State = 1 Then rs.Close
rs.Open "select * from titleStatent where STNo=" & txtSTNO & "", con, adOpenDynamic, adLockOptimistic

For i = 1 To rs.RecordCount

If rs.EOF = False Then

'cmdSave.Enabled = False

txtdate.Text = rs!Dates
cboBook.Text = rs!BookName

txtrem = rs!des1 & ""
txtBinder = rs!binder & ""

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from itemmaster where ItemName='" & rs!BookName & "'", con, adOpenDynamic, adLockOptimistic
If rs1.EOF = False Then
  cbogp.Text = rs1!itemgp & ""
End If

vs.TextMatrix(i, 0) = rs!date1
vs.TextMatrix(i, 1) = rs!descr
vs.TextMatrix(i, 2) = rs!RecQty

totRec = totRec + rs!RecQty

txtNewBook = rs!NewBookQty
txtITCBook = rs!ITCBookQty
txtTitle.Text = rs!TitleBookQty & ""

txtBal = rs!BalQty



rs.MoveNext
End If


Next

totConsume = (Val(txtNewBook) + Val(txtITCBook) + Val(txtTitle.Text))

txtConsumed = totConsume
txtRecTotal = totRec

End Sub


Private Sub Command1_Click()

If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
 con.Execute "delete from titleStatent where STNo=" & txtSTNO & ""
 
 con.Execute "update ITC set st='n',stno=" & 0 & "   where BookName='" & cboBook.Text & "' and st='y' and STNo='" & txtSTNO & "'"
 con.Execute "update BookReceive set st='n',stno=" & 0 & "   where BookName='" & cboBook.Text & "' and st='y' and STNo='" & txtSTNO & "'"
 con.Execute "update title set st='n',stno=" & 0 & "   where BookName='" & cboBook.Text & "' and st='y' and STNo='" & txtSTNO & "'"
 
 Call cmdRef_Click
End If

End Sub

Private Sub Form_Load()
pos Me

fillcombo cbogp, "Name", "ConsumeGp", con

fillcombo cboBook, "ItemName", "ItemMaster", con


vs.FormatString = "Date|Description|Receive"

vs.ColWidth(0) = 1500
vs.ColWidth(1) = 6200
vs.ColWidth(2) = 1270
'vs.ColWidth(3) = 1500
'vs.ColWidth(4) = 1200

txtSTNO = MaxSNo("titleStatent", "stno")
txtdate.Text = Format(Date, "dd/MM/yyyy")


addSt

If rs.State = 1 Then rs.Close
rs.Open "select cname from setup1", con
If rs.EOF = False Then
   txtBinder.Text = rs!CNAME & ""
End If


End Sub

Sub addSt()


vs.clear

refVs

txtdate.Text = Date
cboBook.Text = ""
txtNewBook = ""
txtITCBook = ""
txtBal = ""

txtConsumed = ""
txtRecTotal = ""

txtrem = ""
txtTitle = ""





End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  cboBook.SetFocus
End If

End Sub

Private Sub txtITCBook_LostFocus()
cal
End Sub

Private Sub txtNewBook_LostFocus()
cal
End Sub

Private Sub txtSTNO_GotFocus()


If PopUpValue1 <> "" Then

addSt

txtSTNO = PopUpValue1

'cmdRef_Click

searchData1

PopUpValue1 = ""
PopUpValue2 = ""


End If

End Sub
Private Sub txtSTNO_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
  popuplist2 "Select distinct STNo,Dates,BookName from titleStatent order by STNo", con
End If

If KeyCode = 13 Then
cbogp.SetFocus
End If

End Sub

Private Sub txtSTNO_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   searchData1
End If

End Sub

Private Sub txtTitle_LostFocus()
cal
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtRecTotal.Text = 0
      'For i = 1 To vs.Rows - 1
      'If vs.TextMatrix(i, 2) <> "" Then
      '   txtRecTotal.Text = Val(txtRecTotal.Text) + Val(vs.TextMatrix(i, 2))
      'End If
      
      'Next
      'End If
      
      
      
      SendKeys "{down}"
    End If
    
    If KeyCode = 46 Then
       vs.RemoveItem (vs.RowSel)
       cal
    End If
    
    
End Sub
