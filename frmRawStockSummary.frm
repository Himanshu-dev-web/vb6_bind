VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRawStockSummary 
   Caption         =   "Stock Summary"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   Icon            =   "frmRawStockSummary.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print(Heat No Wise)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10050
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2385
      Width           =   1725
   End
   Begin VB.ComboBox cboGp 
      Height          =   315
      ItemData        =   "frmRawStockSummary.frx":0442
      Left            =   4665
      List            =   "frmRawStockSummary.frx":0461
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   390
      Width           =   2775
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Show(Date Wise)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10050
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   855
      Width           =   1725
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print(Date Wise)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10050
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1335
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4455
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19857411
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19857411
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10050
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2925
      Width           =   1740
   End
   Begin VB.ComboBox CboPName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7545
      TabIndex        =   0
      Top             =   375
      Width           =   2460
   End
   Begin Crystal.CrystalReport CR 
      Left            =   11115
      Top             =   255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6945
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Width           =   9960
      _cx             =   17568
      _cy             =   12250
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
      FormatString    =   $"frmRawStockSummary.frx":04E3
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4695
      TabIndex        =   11
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   7545
      TabIndex        =   10
      Top             =   75
      Width           =   870
   End
End
Attribute VB_Name = "frmRawStockSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ss5 As New ADODB.Recordset
Dim bb5 As Boolean
Dim boo As Boolean
Private Sub cboItemName_Click()
  On Error Resume Next
  fillgrid
End Sub
Private Function LoadName()
CboPName.Clear
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(itemname) from ItemMaster where ItemGp='" & cbogp.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
    Do While Not rs.EOF
        CboPName.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Function
Private Sub cboGp_Click()
  On Error Resume Next
  Call LoadName
End Sub

Private Sub CboPName_Click()
'    If rs.State = 1 Then rs.Close
'    rs.Open "select unit from ItemMaster where ItemName='" & CboPName.Text & "'", con
'    If rs.EOF = False Then
'       unit.Caption = rs(0)
'    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CmdPrint_Click()
    
   'Call cmdSearch_Click
   
   CR.Reset
   CR.ReportFileName = App.Path & "\rawstockregister.rpt"
   'cr.ReplaceSelectionFormula "{RowStockSummaryRegister.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {RowStockSummaryRegister.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
   CR.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "Todate='" & ToDate.Value & "'"
   CR.Formulas(2) = "gp='" & cbogp.Text & "'"
   CR.WindowShowCloseBtn = True
   CR.WindowShowPrintBtn = True
   CR.WindowControlBox = True
   CR.WindowShowPrintSetupBtn = True
   CR.WindowShowProgressCtls = True
   CR.WindowState = crptMaximized
   CR.Action = 1


End Sub
Sub UpdateStock()
     
     Dim save As New ADODB.Recordset
     Dim rs As New ADODB.Recordset
     Dim Search As New ADODB.Recordset
     Dim search1 As New ADODB.Recordset
     Dim rss As New ADODB.Recordset
     Dim oprs As New ADODB.Recordset
     Dim rs_heat As New ADODB.Recordset
     
     Dim heatno
     Dim Receive
     Dim Issue
     Dim purchaseret
     Dim opening
     Dim opDate As Date
     Dim pr_rec, pr_issue
     
     
     Dim quality As String
     Screen.MousePointer = vbHourglass
     heatno = 0
         
         
         
Dim rs_S As New ADODB.Recordset

    
quality = cbogp.Text
If Search.State = 1 Then Search.Close
Search.Open "select distinct(Dates) from [dates] where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "')  order by Dates", con

If Search.EOF = False Then
        
   While Search.EOF = False
             
                 
          '=================================
             
             If rs.State = 1 Then rs.Close
             If cbogp.Text <> "" And CboPName.Text <> "" Then
                 rs.Open "select distinct(ItemName) from itemMaster where ItemGp='" & quality & "' and ItemName='" & CboPName.Text & "'", con
             ElseIf cbogp.Text <> "" And CboPName.Text = "" Then
                 rs.Open "select distinct(ItemName) from itemMaster where ItemGp='" & quality & "'", con
             End If
           
             If rs.EOF = False Then
             While rs.EOF = False
                      '------------Calculate Opening--------
                      op = 0
                      If oprs.State = 1 Then oprs.Close
                      oprs.Open "select OpeningStock,Openingdate from itemMaster where ItemGp='" & quality & "' and itemName='" & rs.Fields(0).Value & "'", con
                      If oprs.EOF = False Then
                            If rss.State = 1 Then rss.Close
                            rss.Open "select * from RawStockSummary where Name='" & rs.Fields(0).Value & "'", con
                            If rss.EOF = True Then
                               op = oprs.Fields(0).Value
                            End If
                        Else
                            op = 0
                      End If
                  
                   
                  'If quality <> "Raw Item" Or rs.Fields(0).Value = "MA" Then
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_rec = search1.Fields(0).Value
                        Else
                        pr_rec = 0
                     End If
                  'Else
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from Purchase where RecDate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                        Else
                        Receive = 0
                     End If
                  'End If
                     
                   pr_rec = pr_rec + Receive
                     
                     
                  If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Issue'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_issue = search1.Fields(0).Value
                        Else
                        pr_issue = 0
                  End If
                     
                    
                     
                     
                  op = (op + (pr_rec - pr_issue))
                     
                  ' boo = False
                     
                     '-----------------end Code-------------------
                  
                  'If Search(0) = "31/05/2006" Then
                  '   MsgBox "sss"
                  'End If
                     
                  'If quality <> "Raw Item" Or rs.Fields(0).Value = "MA" Then
                     
                     
    '===================================
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where HeatingDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                        Else
                        Receive = 0
                     End If
                 'Else
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from Purchase where RecDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pReceive = search1.Fields(0).Value
                        Else
                        pReceive = 0
                     End If
                 ' End If
                  
                     Receive = Receive + pReceive
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where HeatingDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Issue'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Issue = search1.Fields(0).Value
                        Else
                        Issue = 0
                     End If
               
         ''''''''''''''
         ''''''''''''''
         heatno = ""
         If rs_heat.State = 1 Then rs_heat.Close
         rs_heat.Open "select Heatingno from IssuItem where HeatingDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "' and (Status='Issue' or Status='Receive')", con
         While rs_heat.EOF = False
             If heatno = "" Then
                heatno = rs_heat.Fields(0).Value
                Else
                heatno = heatno & " , " & rs_heat.Fields(0).Value
             End If
             rs_heat.MoveNext
         Wend
         '''''''''''''''
         '''''''''''''''
                     
        
        
        
        'If search1.State = 1 Then search1.Close
        'search1.Open "select * from RawStockSummary where Name='" & rs.Fields(0).Value & "'", con
        'If search1.EOF = True Then
        '   boo = False
        'Else
        '   boo = True
        'End If
        
        '--------------------- Data Fiter and Now  Save Coding========================
         opening = 0
         
         If Issue = 0 And Receive = 0 And op = 0 Then
               GoTo aaa:
         End If
          
          
          
       If search1.State = 1 Then search1.Close
       search1.Open "select * from RawStockSummary where name='" & rs.Fields(0).Value & "'", con
       If search1.EOF = True Then
          bb5 = True
       Else
          bb5 = False
       End If
        
       
       boo = False
       
       If bb5 = True Then
       
       
            Set save = New ADODB.Recordset
            If save.State = 1 Then save.Close
            save.Open "select * from RawStockSummary", con, adOpenDynamic, adLockOptimistic
            save.AddNew
            save!Name = rs.Fields(0).Value
            save!gpname = quality
            ss = rs.Fields(0).Value
            save!dates = Search.Fields(0).Value
            save!OpenStock = op
            save!Issue = Issue
            save!ReceiveStock = Receive
            save!ClosingStock = ((op + Receive) - (Issue))
            save!heatno = heatno
            save.Update
            boo = True
       
       Else
       
       If Issue > 0 Or Receive > 0 Then
           Set save = New ADODB.Recordset
           If save.State = 1 Then save.Close
           save.Open "select * from RawStockSummary", con, adOpenDynamic, adLockOptimistic
           save.AddNew
           save!Name = rs.Fields(0).Value
           save!gpname = quality
           ss = rs.Fields(0).Value
           save!dates = Search.Fields(0).Value
           save!OpenStock = op
           save!Issue = Issue
           save!ReceiveStock = Receive
           save!ClosingStock = ((op + Receive) - (Issue))
           save!heatno = heatno
           save.Update
           boo = True
      End If


      End If
        
     '==================================
          
          
          If rss.State = 1 Then rss.Close
          rss.Open "select ClosingStock,Aouto  from RawStockSummary where Name='" & ss & "'  order by Dates,aouto", con
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
              save.Open "select * from RawStockSummary where Aouto=" & nn & "", con, adOpenDynamic, adLockOptimistic
              If save.EOF = False Then
                 
              
              
                 save!OpenStock = opening
                 save!ClosingStock = ((opening + Receive) - (Issue))
              
              If boo = True Then
                 save.Update
              End If
                 
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
                                rs.Open "select distinct(Name) from RawStockSummary where GpName='" & quality & "'", con
                                If rs.EOF = False Then
                                While rs.EOF = False
                                
                                                '------------Calculate Opening--------
                                op = 0
                                If oprs.State = 1 Then oprs.Close
                                    oprs.Open "select OpeningStock,Openingdate from ItemMaster where ItemGp='" & quality & "' and ItemName='" & rs.Fields(0).Value & "'", con
                                    If oprs.EOF = False Then
                                        If rss.State = 1 Then rss.Close
                                            rss.Open "select * from RawStockSummary where GpName='" & quality & "' and Name='" & rs.Fields(0).Value & "'", con
                                            If rss.EOF = True Then
                                                op = oprs.Fields(0).Value
                                            End If
                                            Else
                                        op = 0
                                    End If
                                
                      
                    ' If quality = "Raw Item" Then
                            
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from IssueItem where HeatingDate<datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                            If Not IsNull(search1.Fields(0)) Then
                            Receive = search1.Fields(0).Value
                            Else
                            Receive = 0
                            End If
                     ' Else
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from Purchase where RecDate<datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", con
                            If Not IsNull(search1.Fields(0)) Then
                               Receive_p = search1.Fields(0).Value
                            Else
                               Receive_p = 0
                            End If
                      ' End If
                      
                            Receive = Receive + Receive_p
                      
                          
                            
                            
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & Search(0) & "') and Item='" & rs.Fields(0).Value & "' and Status='Issue'", con
                            If Not IsNull(search1.Fields(0)) Then
                            Issue = search1.Fields(0).Value
                            Else
                            Issue = 0
                            End If
           
                                                           
                                                          
                                                           
                                                           
                                                                 
                                         op = ((op + pr_rec) - pr_issue)

                                         If op = 0 Then GoTo ab1:
                                         
                                ' If Issue > 0 Or Receive > 0 Then
        
                                         
                                         Set save = New ADODB.Recordset
                                         If save.State = 1 Then save.Close
                                         save.Open "select * from RawStockSummary", con, adOpenDynamic, adLockOptimistic
                                         save.AddNew
                                         save!GroupName = quality
                                         save!gpname = quality
                                         save!Name = rs.Fields(0).Value
                                         ss = rs.Fields(0).Value
                                         save!dates = FromDate.Value
                                         save!OpenStock = op
                                         save!ReceiveStock = Receive
                                         save!Issue = Issue
                                         save!ClosingStock = ((save!OpenStock + Receive) - (Issue))
                                         save.Update
                                ' End If
                                         
ab1:

                                       rs.MoveNext
                                   Wend
                                   End If
  
 '==============

    
      
    
End If
            

  
   
Screen.MousePointer = vbDefault

End Sub
Sub UpdateStock_Heatno()
     
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

    
quality = cbogp.Text
If Search.State = 1 Then Search.Close
Search.Open "select distinct(Dates) from [dates] where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "')  order by Dates", con

If Search.EOF = False Then
        
   While Search.EOF = False
             
                 
          '=================================
             
             If rs.State = 1 Then rs.Close
             If cbogp.Text <> "" And CboPName.Text <> "" Then
                 rs.Open "select distinct(ItemName) from itemMaster where ItemGp='" & quality & "' and ItemName='" & CboPName.Text & "'", con
             ElseIf cbogp.Text <> "" And CboPName.Text = "" Then
                 rs.Open "select distinct(ItemName) from itemMaster where ItemGp='" & quality & "'", con
             End If
           
             If rs.EOF = False Then
             While rs.EOF = False
                      '------------Calculate Opening--------
                      op = 0
                      If oprs.State = 1 Then oprs.Close
                      oprs.Open "select OpeningStock,Openingdate from itemMaster where ItemGp='" & quality & "' and itemName='" & rs.Fields(0).Value & "'", con
                      If oprs.EOF = False Then
                            If rss.State = 1 Then rss.Close
                            rss.Open "select * from RawStockSummary where Name='" & rs.Fields(0).Value & "'", con
                            If rss.EOF = True Then
                               op = oprs.Fields(0).Value
                            End If
                        Else
                            op = 0
                      End If
                  
                   
                  'If quality <> "Raw Item" Or rs.Fields(0).Value = "MA" Then
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_rec = search1.Fields(0).Value
                        Else
                        pr_rec = 0
                     End If
                  'Else
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from Purchase where RecDate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                        Else
                        Receive = 0
                     End If
                  'End If
                     
                   pr_rec = pr_rec + Receive
                     
                     
                  If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Issue'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_issue = search1.Fields(0).Value
                        Else
                        pr_issue = 0
                  End If
                     
                    
                     
                     
                  op = (op + (pr_rec - pr_issue))
                     
                     
                     
                     '-----------------end Code-------------------
                  
                  'If Search(0) = "31/05/2006" Then
                  '   MsgBox "sss"
                  'End If
                     
                  'If quality <> "Raw Item" Or rs.Fields(0).Value = "MA" Then
                     
                     
    '===================================
     If ss5.State = 1 Then ss5.Close
     ss5.Open "select HeatingNo from IssuItem where HeatingDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                     
     While ss5.EOF = False
     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where HeatingNo=" & ss5.Fields(0).Value & " and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                        Else
                        Receive = 0
                     End If
                 'Else
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from Purchase where RecDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pReceive = search1.Fields(0).Value
                        Else
                        pReceive = 0
                     End If
                 ' End If
                  
                     Receive = Receive + pReceive
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where HeatingDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Issue'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Issue = search1.Fields(0).Value
                        Else
                        Issue = 0
                     End If
               
                     
                     

                     
                     
        
        '--------------------- Data Fiter and Now  Save Coding========================
         opening = 0
         If Issue = 0 And Receive = 0 And op = 0 Then
            GoTo aaa:
         End If
          
       'If Issue > 0 Or Receive > 0 Then
          
          Set save = New ADODB.Recordset
          If save.State = 1 Then save.Close
          save.Open "select * from RawStockSummary", con, adOpenDynamic, adLockOptimistic
          save.AddNew
          save!Name = rs.Fields(0).Value
          save!gpname = quality
          ss = rs.Fields(0).Value
          save!dates = Search.Fields(0).Value
          save!OpenStock = op
          save!Issue = Issue
          save!ReceiveStock = Receive
          save!ClosingStock = ((op + Receive) - (Issue))
          save.Update
          
        'End If
          
         ss5.MoveNext
     Wend
          
     '==================================
          
     
          
          
          If rss.State = 1 Then rss.Close
          rss.Open "select ClosingStock,Aouto  from RawStockSummary where Name='" & ss & "'  order by Dates,aouto", con
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
              save.Open "select * from RawStockSummary where Aouto=" & nn & "", con, adOpenDynamic, adLockOptimistic
              If save.EOF = False Then
                 save!OpenStock = opening
                 save!ClosingStock = ((opening + Receive) - (Issue))
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
                                rs.Open "select distinct(Name) from RawStockSummary where GpName='" & quality & "'", con
                                If rs.EOF = False Then
                                While rs.EOF = False
                                
                                                '------------Calculate Opening--------
                                op = 0
                                If oprs.State = 1 Then oprs.Close
                                    oprs.Open "select OpeningStock,Openingdate from ItemMaster where ItemGp='" & quality & "' and ItemName='" & rs.Fields(0).Value & "'", con
                                    If oprs.EOF = False Then
                                        If rss.State = 1 Then rss.Close
                                            rss.Open "select * from RawStockSummary where GpName='" & quality & "' and Name='" & rs.Fields(0).Value & "'", con
                                            If rss.EOF = True Then
                                                op = oprs.Fields(0).Value
                                            End If
                                            Else
                                        op = 0
                                    End If
                                
                      
                    ' If quality = "Raw Item" Then
                            
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from IssueItem where HeatingDate<datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                            If Not IsNull(search1.Fields(0)) Then
                            Receive = search1.Fields(0).Value
                            Else
                            Receive = 0
                            End If
                     ' Else
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from Purchase where RecDate<datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", con
                            If Not IsNull(search1.Fields(0)) Then
                               Receive_p = search1.Fields(0).Value
                            Else
                               Receive_p = 0
                            End If
                      ' End If
                      
                            Receive = Receive + Receive_p
                      
                          
                            
                            
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & Search(0) & "') and Item='" & rs.Fields(0).Value & "' and Status='Issue'", con
                            If Not IsNull(search1.Fields(0)) Then
                            Issue = search1.Fields(0).Value
                            Else
                            Issue = 0
                            End If
           
                                                           
                                                          
                                                           
                                                           
                                                                 
                                         op = ((op + pr_rec) - pr_issue)

                                         If op = 0 Then GoTo ab1:
                                         
                                         Set save = New ADODB.Recordset
                                         If save.State = 1 Then save.Close
                                         save.Open "select * from RawStockSummary", con, adOpenDynamic, adLockOptimistic
                                         save.AddNew
                                         save!GroupName = quality
                                         save!gpname = quality
                                         save!Name = rs.Fields(0).Value
                                         ss = rs.Fields(0).Value
                                         save!dates = FromDate.Value
                                         save!OpenStock = op
                                         save!ReceiveStock = Receive
                                         save!Issue = Issue
                                          save!ClosingStock = ((save!OpenStock + Receive) - (Issue))
                                         save.Update
                                         
                                         
ab1:

                                       rs.MoveNext
                                   Wend
                                   End If
  
 '==============

    
      
    
End If
            

  
   
Screen.MousePointer = vbDefault

End Sub

Sub AddDate()
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim k As Integer
    Dim d As Date
    Set cmd.ActiveConnection = con
    
    cmd.CommandText = "delete from dates"
    cmd.Execute
    k = DateDiff("d", FromDate.Value, ToDate.Value)
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dates", con, adOpenDynamic, adLockOptimistic
     
     
   '   d = FromDate.Value
   '  For i = 0 To k
   '        rs.AddNew
   '        rs!dates = d
   '        rs.Update
   '        d = d + 1
   '   Next
     
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select distinct(HeatingDate) from issuemaster", con, adOpenDynamic, adLockOptimistic
    While rs1.EOF = False
           rs.AddNew
           rs!dates = rs1.Fields(0).Value
           rs.Update
           rs1.MoveNext
     Wend

    If rs1.State = 1 Then rs1.Close
    rs1.Open "select distinct(RecDate) from Purchase", con, adOpenDynamic, adLockOptimistic
    While rs1.EOF = False

         If rs.State = 1 Then rs.Close
         rs.Open "select * from dates where Dates=datevalue('" & rs1.Fields(0).Value & "')", con, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.AddNew
            rs!dates = rs1.Fields(0).Value
            rs.Update
          End If
          rs1.MoveNext
     Wend
    
End Sub
Private Sub cmdSearch_Click()
    
    
    Screen.MousePointer = vbHourglass
    
    AddDate
    Dim opening As Long
    Dim Search As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = con
    cmd.CommandText = "delete from RawStockSummary"
    cmd.Execute
    
    UpdateStock
    
    fillgrid
   
    Screen.MousePointer = vbDefault
End Sub
Sub fillgrid()
     Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    fill.Open "select Name,dates,OpenStock,ReceiveStock,Issue,ClosingStock from RawStockSummary where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "') order by Dates,Aouto", con
    If fill.EOF = False Then
       'Set vs.DataSource = fill
       vs.FormatString = "Name|^date|>OpenStock|>Received|>Issued|>ClosingStock"
       
       '-----------------------------------------
       
       vs.Rows = fill.RecordCount + 1
       For i = 1 To fill.RecordCount
         If fill.EOF = False Then
         
            vs.TextMatrix(i, 0) = fill.Fields(0).Value
            vs.TextMatrix(i, 1) = fill.Fields(1).Value
            vs.TextMatrix(i, 2) = Format(fill.Fields(2).Value, "#,##.00")
            vs.TextMatrix(i, 3) = Format(fill.Fields(3).Value, "#,##.00")
            vs.TextMatrix(i, 4) = Format(fill.Fields(4).Value, "#,##.00")
            vs.TextMatrix(i, 5) = Format(fill.Fields(5).Value, "#,##.00")
            
         
         End If
         fill.MoveNext
       Next
       
       '-----------------------------------------
     
     Else
       Set vs.DataSource = fill
       vs.FormatString = "Name|^date|>OpenStock|>Purchase|>Issue|>ClosingStock"
     End If
     
     VsWidth

End Sub
Sub fillgrid_11()
    
    Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    fill.Open "select Name,dates,OpenStock,ReceiveStock,Issue,ClosingStock from RawStockSummary where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "') order by Dates,Aouto", con
    If fill.EOF = False Then
       'Set vs.DataSource = fill
       vs.FormatString = "Name|^date|>OpenStock|>Received|>Issued|>ClosingStock"
       
       '-----------------------------------------
       
       vs.Rows = fill.RecordCount + 1
       For i = 1 To fill.RecordCount
         If fill.EOF = False Then
         
            vs.TextMatrix(i, 0) = fill.Fields(0).Value
            vs.TextMatrix(i, 1) = fill.Fields(1).Value
            vs.TextMatrix(i, 2) = Format(fill.Fields(2).Value, "#,##.00")
            vs.TextMatrix(i, 3) = Format(fill.Fields(3).Value, "#,##.00")
            vs.TextMatrix(i, 4) = Format(fill.Fields(4).Value, "#,##.00")
            vs.TextMatrix(i, 5) = Format(fill.Fields(5).Value, "#,##.00")
            
         
         End If
         fill.MoveNext
       Next
       
       '-----------------------------------------
     
     Else
       Set vs.DataSource = fill
       vs.FormatString = "Name|^date|>OpenStock|>Purchase|>Issue|>ClosingStock"
     End If
     
     VsWidth

End Sub

Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass
    
    AddDate
    Dim opening As Long
    Dim Search As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = con
    cmd.CommandText = "delete from RawStockSummary"
    cmd.Execute
    
    UpdateStock_Heatno
    
    fillgrid_11
   
    Screen.MousePointer = vbDefault

End Sub

Private Sub Command2_Click()
   CR.Reset
   CR.ReportFileName = App.Path & "\heating.rpt"
   
   If CboPName.Text = "" Then
   
    CR.ReplaceSelectionFormula "{issuitem.heatingdate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {issuitem.heatingdate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "') and {issuitem.itemgp}='" & cbogp.Text & "'"
   
   Else
    
    CR.ReplaceSelectionFormula "{issuitem.heatingdate}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {issuitem.heatingdate}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "') and {issuitem.itemgp}='" & cbogp.Text & "' and {issuitem.itemname}='" & CboPName.Text & "'"
   
   End If
   
   CR.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "Todate='" & ToDate.Value & "'"
   'CR.Formulas(2) = "gp='" & cboGp.Text & "'"
   CR.WindowShowCloseBtn = True
   CR.WindowShowPrintBtn = True
   CR.WindowControlBox = True
   CR.WindowShowPrintSetupBtn = True
   CR.WindowShowProgressCtls = True
   CR.WindowShowSearchBtn = True
   CR.WindowState = crptMaximized
   CR.Action = 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     Unload Me
  End If
End Sub
Sub VsWidth()
       
       vs.ColWidth(0) = 2500
       vs.ColWidth(1) = 1100
       vs.ColWidth(2) = 1600
       vs.ColWidth(3) = 1600
       vs.ColWidth(4) = 1700
       vs.ColWidth(5) = 1700
       
 
End Sub

Private Sub Form_Load()
ToDate.Value = Date
fillgrid
FromDate.Value = Date




cbogp.ListIndex = 0

End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then ToDate.SetFocus
End Sub
