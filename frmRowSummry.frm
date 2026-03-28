VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrawSummry 
   Caption         =   "Raw Stock Suumry"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRowSummry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   10245
      Top             =   -15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox CboPName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   2
      Top             =   450
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FDD6B3&
      Caption         =   "&Exit"
      Height          =   495
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19857411
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19857411
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FDD6B3&
      Caption         =   "&Print"
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FDD6B3&
      Caption         =   "&Search"
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6975
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   11775
      _cx             =   20770
      _cy             =   12303
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16705231
      ForeColor       =   0
      BackColorFixed  =   16706264
      ForeColorFixed  =   4227327
      BackColorSel    =   16448755
      ForeColorSel    =   16711680
      BackColorBkg    =   16770764
      BackColorAlternate=   16705231
      GridColor       =   16744448
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   -1  'True
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
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   5205
      TabIndex        =   10
      Top             =   165
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmrawSummry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboItem_Click()
  On Error Resume Next
  fillgrid
End Sub
Private Function LoadName()
CboPName.Clear
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(pname) from newRAWmaster where groupname='" & Trim(CboGroup.Text) & "'", adoConn1, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
    Do While Not rs.EOF
        CboPName.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Function


Private Sub CboGroup_Click()
Call LoadName
End Sub

Private Sub CboGroup_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Call LoadName
   CboPName.SetFocus
 End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdPrint_Click()
CR.Reset
CR.ReportFileName = App.Path & "/RawStockSummry.RPT"
CR.ReplaceSelectionFormula "{RawStockSummry.Dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {RawStockSummry.Dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "') "
CR.Formulas(0) = "FROMDATE='" & FromDate.Value & "'"
CR.Formulas(1) = "TODATE='" & ToDate.Value & "'"
CR.WindowShowCloseBtn = True
CR.WindowShowPrintBtn = True
CR.WindowControlBox = True
CR.WindowShowPrintSetupBtn = True
CR.WindowShowProgressCtls = True
CR.WindowState = crptMaximized

CR.Action = 1
End Sub
Sub UpdateCalanderStock()
     
     Dim save As New ADODB.Recordset
     Dim rs As New ADODB.Recordset
     Dim Search As New ADODB.Recordset
     Dim search1 As New ADODB.Recordset
     Dim rss As New ADODB.Recordset
     Dim oprs As New ADODB.Recordset
     Dim Receive
     Dim Issue
     Dim purchaseret
     Dim opening
     Dim opDate As Date
     Dim pr_rec, pr_issue
     
     
     Dim quality As String
     Screen.MousePointer = vbHourglass
      
         
         
         
Dim rs_S As New ADODB.Recordset
If CboGroup.Text = "" And CboPName.Text = "" Then
    rs_S.Open "select distinct(GroupName) from NewRawMaster", adoConn1
   ElseIf CboGroup.Text <> "" And CboPName.Text = "" Then
    rs_S.Open "select distinct(GroupName) from NewRawMaster where GroupName='" & CboGroup.Text & "'", adoConn1
   ElseIf CboGroup.Text <> "" And CboPName.Text <> "" Then
    rs_S.Open "select distinct(GroupName) from NewRawMaster where GroupName='" & CboGroup.Text & "' and Pname='" & CboPName.Text & "'", adoConn1
End If


If rs_S.EOF = False Then
         
         
         
While rs_S.EOF = False
'========== Code for Than================

quality = rs_S.Fields(0).Value

If Search.State = 1 Then Search.Close
Search.Open "select distinct(Dates) from [dates] where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "')  order by Dates", adoConn1

If Search.EOF = False Then
While Search.EOF = False
    
           If rs.State = 1 Then rs.Close
           
           If CboGroup.Text = "" And CboPName.Text = "" Then
               rs.Open "select distinct(PName) from NewRawMaster where GroupName='" & quality & "'", adoConn1
              ElseIf CboGroup.Text <> "" And CboPName.Text = "" Then
               rs.Open "select distinct(PName) from NewRawMaster where GroupName='" & quality & "'", adoConn1
              ElseIf CboGroup.Text <> "" And CboPName.Text <> "" Then
               rs.Open "select distinct(PName) from NewRawMaster where GroupName='" & quality & "' and pname='" & CboPName.Text & "'", adoConn1
           End If
           
           If rs.EOF = False Then
                While rs.EOF = False
                      '------------Calculate Opening--------
                      Op = 0
                      If oprs.State = 1 Then oprs.Close
                      oprs.Open "select OPQty,Opdate from NewRawMaster where GroupName='" & quality & "' and PName='" & rs.Fields(0).Value & "'", adoConn1
                      If oprs.EOF = False Then
                            If rss.State = 1 Then rss.Close
                            rss.Open "select * from RawStockSummry where GroupName='" & quality & "' and PName='" & rs.Fields(0).Value & "'", adoConn1
                            If rss.EOF = True Then
                               Op = oprs.Fields(0).Value
                            End If
                        Else
                            Op = 0
                      End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from RawPurchase where Invdate<datevalue('" & FromDate.Value & "') and [Group]='" & quality & "' and [Name]='" & rs.Fields(0).Value & "'", adoConn1
                     If Not IsNull(search1.Fields(0)) Then
                        pr_rec = search1.Fields(0).Value
                        Else
                        pr_rec = 0
                     End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from RawIssuetable where idate<datevalue('" & FromDate.Value & "') and RGroup='" & quality & "' and RName='" & rs.Fields(0).Value & "' and type='R'", adoConn1
                     If Not IsNull(search1.Fields(0)) Then
                        pr_issue = search1.Fields(0).Value
                        Else
                        pr_issue = 0
                     End If
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from RawPurchaseReturn where invdate<datevalue('" & FromDate.Value & "') and [Group]='" & quality & "' and [Name]='" & rs.Fields(0).Value & "'", adoConn1
                     If Not IsNull(search1.Fields(0)) Then
                        purchaseret = search1.Fields(0).Value
                        Else
                        purchaseret = 0
                     End If
                     
                     
                     
                     Op = (Op + (pr_rec - (pr_issue + purchaseret)))
                     
                     
                     
                     '-----------------end Code-------------------

                     
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from RawPurchase where Invdate=datevalue('" & Search(0) & "') and [Group]='" & quality & "' and Name='" & rs.Fields(0).Value & "'", adoConn1
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                        Else
                        Receive = 0
                     End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from RawIssuetable where idate=datevalue('" & Search(0) & "') and RGroup='" & quality & "' and RName='" & rs.Fields(0).Value & "' and type='R'", adoConn1
                     If Not IsNull(search1.Fields(0)) Then
                        Issue = search1.Fields(0).Value
                        Else
                        Issue = 0
                     End If
               
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from RawPurchaseReturn where invdate=datevalue('" & Search(0) & "') and [Group]='" & quality & "' and [Name]='" & rs.Fields(0).Value & "'", adoConn1
                     If Not IsNull(search1.Fields(0)) Then
                        purchaseret = search1.Fields(0).Value
                        Else
                        purchaseret = 0
                     End If
                     
                     
                     
        
        '--------------------- Data Fiter and Now  Save Coding========================
         opening = 0
         If Issue = 0 And Receive = 0 And Op = 0 And purchaseret = 0 Then
            GoTo aaa:
         End If
          
          Set save = New ADODB.Recordset
          If save.State = 1 Then save.Close
          save.Open "select * from RawStockSummry", adoConn1, adOpenDynamic, adLockOptimistic
          save.AddNew
          save!GroupName = quality
          save!PName = rs.Fields(0).Value
          ss = rs.Fields(0).Value
          save!dates = Search.Fields(0).Value
          save!OpenStock = Op
          save!Issue = Issue
          save!ReceiveStock = Receive
          save!PurhaseReturn = purchaseret
          save!ClosingStock = ((Op + Receive) - ((Issue) + (purchaseret)))
          save.Update
          
          
          
          If rss.State = 1 Then rss.Close
          rss.Open "select ClosingStock,Aouto  from RawStockSummry where PName='" & ss & "' and GroupName='" & quality & "' order by Dates,aouto", adoConn1
          If rss.EOF = False Then
                
              If rss.RecordCount = 1 Then
                opening = rss.Fields(0).Value
              ElseIf rss.RecordCount = 2 Then
                rss.MoveLast
                rss.MovePrevious
                opening = rss.Fields(0).Value
                rss.MoveNext
                nn = rss.Fields(1).Value
              ElseIf rss.RecordCount > 2 Then
                rss.MoveLast
                rss.MovePrevious
                opening = rss.Fields(0).Value
                rss.MoveNext
                nn = rss.Fields(1).Value
              
              End If
              
              
             If nn <> "" Then
              Set save = New ADODB.Recordset
              If save.State = 1 Then save.Close
              save.Open "select * from RawStockSummry where Aouto=" & nn & "", adoConn1, adOpenDynamic, adLockOptimistic
              If save.EOF = False Then
                 save!OpenStock = opening
                 save!ClosingStock = ((opening + Receive) - ((Issue) + (purchaseret)))
                 save.Update
                 opening = 0
                 nn = ""
              End If
             End If
             
             Else
              
              
          End If
aaa:
                                       
                                       
                                   rs.MoveNext
                                  Wend
                                  End If
                                   
                                   
                                   
                                   Search.MoveNext
                                   Wend
    
    
    Else
    
    
    
    
              '------------------- Opening Show No Transiction

                                
                                
                                If rs.State = 1 Then rs.Close
                                rs.Open "select distinct(GroupName) from RawStockSummry where GroupName='" & quality & "'", adoConn1
                                If rs.EOF = False Then
                                While rs.EOF = False
                                
                                                '------------Calculate Opening--------
                                Op = 0
                                If oprs.State = 1 Then oprs.Close
                                    oprs.Open "select OPQty,Opdate from NewRawMaster where GroupName='" & quality & "' and PName='" & rs.Fields(0).Value & "'", adoConn1
                                    If oprs.EOF = False Then
                                        If rss.State = 1 Then rss.Close
                                            rss.Open "select * from RawStockSummry where GroupName='" & quality & "' and PName='" & rs.Fields(0).Value & "'", adoConn1
                                            If rss.EOF = True Then
                                                Op = oprs.Fields(0).Value
                                            End If
                                            Else
                                        Op = 0
                                    End If
                                

                                    If search1.State = 1 Then search1.Close
                                    search1.Open "select sum(Qty) from RawPurchase where Invdate<datevalue('" & FromDate.Value & "') and [Group]='" & quality & "' and Name='" & rs.Fields(0).Value & "'", adoConn1
                                    If Not IsNull(search1.Fields(0)) Then
                                     pr_rec = search1.Fields(0).Value
                                     Else
                                     pr_rec = 0
                                    End If
                                    
                                    
                                    If search1.State = 1 Then search1.Close
                                    search1.Open "select sum(Qty) from RawIssuetable where idate<datevalue('" & FromDate.Value & "') and RGroup='" & quality & "' and RName='" & rs.Fields(0).Value & "' and type='R'", adoConn1
                                    If Not IsNull(search1.Fields(0)) Then
                                     pr_issue = search1.Fields(0).Value
                                     Else
                                     pr_issue = 0
                                    End If
                                                           
                                                           
                                                           
                                    If search1.State = 1 Then search1.Close
                                    search1.Open "select sum(Qty) from RawPurchaseReturn where invdate<datevalue('" & FromDate.Value & "') and [Group]='" & quality & "' and [Name]='" & rs.Fields(0).Value & "'", adoConn1
                                    If Not IsNull(search1.Fields(0)) Then
                                    purchaseret = search1.Fields(0).Value
                                    Else
                                    purchaseret = 0
                                    End If
                                                           
                                                           
                                                           
                                                                 
                                         Op = ((Op + pr_rec) - ((pr_issue) + (purchaseret)))

                                         If Op = 0 Then GoTo ab1:
                                         
                                         Set save = New ADODB.Recordset
                                         If save.State = 1 Then save.Close
                                         save.Open "select * from RawStockSummry", con, adOpenDynamic, adLockOptimistic
                                         save.AddNew
                                         save!GroupName = quality
                                         save!PName = rs.Fields(0).Value
                                         ss = rs.Fields(0).Value
                                         save!dates = FromDate.Value
                                         save!OpenStock = Op
                                         save!ReceiveStock = Receive
                                         save!Issue = Issue
                                         save!PurhaseReturn = purchaseret
                                         save!ClosingStock = ((save!OpenStock + Receive) - ((Issue) + (purchaseret)))
                                         save.Update
ab1:

                                       rs.MoveNext
                                   Wend
                                   End If
  
 '==============

    
      
    
End If
            
            
            
            rs_S.MoveNext
            Wend
       End If
  
   
    Screen.MousePointer = vbDefault

End Sub
Sub AddDate()
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim k As Integer
    Dim d As Date
    Set cmd.ActiveConnection = adoConn1
    
    cmd.CommandText = "delete from dates"
    cmd.Execute
    k = DateDiff("d", FromDate.Value, ToDate.Value)
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dates", adoConn1, adOpenDynamic, adLockOptimistic
     d = FromDate.Value
     For i = 0 To k
           rs.AddNew
           rs!dates = d
           rs.Update
           d = d + 1
       Next
    
End Sub
Private Sub cmdSearch_Click()
    
    Screen.MousePointer = vbHourglass
    
    AddDate
    Dim opening As Long
    Dim Search As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = adoConn1
    cmd.CommandText = "delete from RawStockSummry"
    cmd.Execute
    
    UpdateCalanderStock
    
    Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    fill.Open "select * from RawStockSummry where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "') order by Dates,Aouto", adoConn1
    If fill.EOF = False Then
       Set vs.DataSource = fill
       vs.FormatString = "GroupName|PName|^date|>OpenStock|>Purchase|>PurchaseReturn|>Issue|>ClosingStock"
     Else
       Set vs.DataSource = fill
       vs.FormatString = "GroupName|PName|^date|>OpenStock|>Purchase|>PurchaseReturn|>Issue|>ClosingStock"
     End If
     
     VsWidth
     Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     Unload Me
  End If
End Sub
Sub VsWidth()
       vs.ColWidth(0) = 2200
       vs.ColWidth(1) = 1500
       vs.ColWidth(2) = 1400
       vs.ColWidth(3) = 1400
       vs.ColWidth(4) = 1300
       vs.ColWidth(5) = 1300
       vs.ColWidth(6) = 1300
       vs.ColWidth(7) = 1300
End Sub
Sub fillgrid()
    Dim fill As New ADODB.Recordset
    fill.Open "select * from RawStockSummry order by dates,Aouto", con
    If fill.EOF = False Then
       Set vs.DataSource = fill
       vs.FormatString = "GroupName|PName|^date|>OpenStock|>Purchase|>PurchaseReturn|>Issue|>ClosingStock"
       VsWidth
     End If
End Sub

Private Sub Form_Load()
ToDate.Value = Date
fillgrid
Call LoadGroup
'On Error Resume Next
FromDate.Value = Date
End Sub

Private Function LoadGroup()
CboGroup.Clear
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(GroupName) from NewRawMaster", adoConn1, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
    Do While Not rs.EOF
        CboGroup.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Function
