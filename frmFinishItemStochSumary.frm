VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmFinishItemStochSumary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statement Of Title"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmFinishItemStochSumary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8145
   Begin VB.CheckBox Check1_withbillNoQty 
      Caption         =   "With Bill Details"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   1980
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrintAllBook 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print (Group Wise Ledger)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1920
      Width           =   3660
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
      Left            =   2040
      TabIndex        =   12
      Top             =   120
      Width           =   4980
   End
   Begin VB.ComboBox CboPName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   8
      Top             =   660
      Width           =   4980
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
      Height          =   495
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   3630
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   8070
      TabIndex        =   2
      Top             =   960
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
         Format          =   64487427
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
         Format          =   64487427
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print (Book Wise Ledger)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   3660
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10155
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   855
      Width           =   1440
   End
   Begin Crystal.CrystalReport CR 
      Left            =   11070
      Top             =   255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4365
      Left            =   15
      TabIndex        =   9
      Top             =   3405
      Width           =   10050
      _cx             =   17727
      _cy             =   7699
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
      FormatString    =   $"frmFinishItemStochSumary.frx":0442
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
      BackStyle       =   0  'Transparent
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
      Left            =   840
      TabIndex        =   13
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label unit 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7875
      TabIndex        =   11
      Top             =   390
      Width           =   675
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   660
      Width           =   1170
   End
End
Attribute VB_Name = "frmFinishItemStochSumary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboItemName_Click()
  On Error Resume Next
  fillgrid
End Sub
Private Function LoadName()
CboPName.clear
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(itemname) from ItemMaster", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
    Do While Not rs.EOF
        CboPName.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Function



Private Sub cboGp_Click()
CboPName.clear

If rs.State = 1 Then rs.Close
rs.Open "select ItemName from  ItemMaster where ItemGp='" & cbogp.Text & "'", con
While rs.EOF = False
CboPName.AddItem rs(0)
rs.MoveNext
Wend

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
   
   Dim all_book As String
   
   
   
   If ss1 = "new" Then
   
   cr.Reset
   cr.ReportFileName = App.Path & "\TaxSummery.rpt"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   'cr.ReplaceSelectionFormula "{BookReceive.BookName}='" & CboPName.Text & "' and {BookReceive.EntDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {BookReceive.entDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.ReplaceSelectionFormula "{BookReceive.BookName}='" & CboPName.Text & "'"
   cr.Formulas(2) = "firm_='" & COMPNAME & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1
   
   
  ElseIf ss1 = "itc" Then
  
   cr.Reset
   cr.ReportFileName = App.Path & "\TaxSummeryitc.rpt"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   'cr.ReplaceSelectionFormula "{BookReceive.BookName}='" & CboPName.Text & "' and {BookReceive.EntDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {BookReceive.entDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.ReplaceSelectionFormula "{BookReceive.BookName}='" & CboPName.Text & "'"
   cr.Formulas(2) = "firm_='" & COMPNAME & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1

  ElseIf ss1 = "title" Then
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
   
tempfortitle
  
  '----------------------------------------------------------------------


If MsgBox("Want To View", vbQuestion + vbYesNo) = vbYes Then

 Screen.MousePointer = vbHourglass
   
   cr.Reset
   cr.ReportFileName = App.Path & "\TaxSummeryTitle.rpt"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   'cr.ReplaceSelectionFormula "{BookReceive.BookName}='" & CboPName.Text & "' and {BookReceive.EntDate}>=datevalue('" & Format(fromDate.Value, "MM/dd/yyyy") & "') and {BookReceive.entDate}<=datevalue('" & Format(todate.Value, "MM/dd/yyyy") & "')"
   cr.ReplaceSelectionFormula "{BookReceive.BookName}='" & CboPName.Text & "'"
   cr.Formulas(2) = "firm_='" & COMPNAME & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 1
 
 Screen.MousePointer = vbDefault

 End If
   
   
   
   
   End If
   
   
   ' myr.ExportOptions.ExportDestinationType = CrystalDecisions.[Shared].ExportDestinationType.DiskFile
   ' myr.ExportOptions.ExportFormatType = CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat
   ' Fname = "c:\" & Session.SessionID.ToString & ".pdf"
   ' DiskOpts.DiskFileName = Fname
   ' myr.ExportOptions.DestinationOptions = DiskOpts
                 


End Sub
Sub tempfortitle()
  Dim rss As New ADODB.Recordset
  
  
  
    
    
    con.Execute "delete from tmpITC where len(BookName)>0"
    
    'Inner_paper_cover
    
    
     
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover)  select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Inner' as inner_paper_cover from ITC where inner_paper_cover='I+P+C'"
     
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover)  select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Paper' as inner_paper_cover from ITC where inner_paper_cover='I+P+C'"
     
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover)  select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Cover' as inner_paper_cover from ITC where inner_paper_cover='I+P+C'"
     
    
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover)  select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Paper' as inner_paper_cover from ITC where inner_paper_cover='P+C'"
     
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover)  select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Cover' as inner_paper_cover from ITC where inner_paper_cover='P+C'"
    
    
    
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover)  select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Paper' as inner_paper_cover from ITC where inner_paper_cover='I+P'"
     
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover)  select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Inner' as inner_paper_cover from ITC where inner_paper_cover='I+P'"
    
     
    
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover) select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Inner' as inner_paper_cover from ITC where inner_paper_cover='I+C'"
     
     con.Execute "insert into tmpITC(EntDate,PartyName,BookName,ChallanNo,Qty,Descr,Issue_Rec,inner_paper_cover)  select EntDate,PartyName,BookName,ChallanNo,Qty," & _
     "Descr,Issue_Rec,'Cover' as inner_paper_cover from ITC where inner_paper_cover='I+C'"

End Sub
Sub UpdateStock()
     
     Dim save As New ADODB.Recordset
     Dim rs As New ADODB.Recordset
     Dim Search As New ADODB.Recordset
     Dim search1 As New ADODB.Recordset
     Dim rss As New ADODB.Recordset
     Dim rs_heat As New ADODB.Recordset
     Dim oprs As New ADODB.Recordset
     Dim Receive
     Dim Issue
     Dim heatno
     Dim purchaseret
     Dim opening
     Dim opDate As Date
     Dim pr_rec, pr_issue
     
     
     Dim quality As String
     Screen.MousePointer = vbHourglass
      
         
         
         
    Dim rs_S As New ADODB.Recordset

    
    quality = "Finish Item"
    If Search.State = 1 Then Search.Close
    Search.Open "select distinct(Dates) from [dates] where Dates>=datevalue('" & fromDate.Value & "') and Dates<=datevalue('" & todate.Value & "')  order by Dates", con
    If Search.EOF = False Then
        
        While Search.EOF = False
           
            If rs.State = 1 Then rs.Close
            If CboPName.Text <> "" Then
               rs.Open "select distinct(ItemName) from itemMaster where ItemName='" & CboPName.Text & "'", con
            '   Else
            '   rs.Open "select distinct(ItemName) from itemMaster where emGp='Finish Item'", con
            End If
           
           If rs.EOF = False Then
                While rs.EOF = False
                      '------------Calculate Opening--------
                      op = 0
                      If oprs.State = 1 Then oprs.Close
                      oprs.Open "select OpeningStock,Openingdate from itemMaster where ItemGp='" & quality & "' and itemName='" & rs.Fields(0).Value & "'", con
                      If oprs.EOF = False Then
                            If rss.State = 1 Then rss.Close
                            rss.Open "select * from FinishStockSummary where Name='" & rs.Fields(0).Value & "'", con
                            If rss.EOF = True Then
                               op = oprs.Fields(0).Value
                            End If
                        Else
                            op = 0
                      End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from BookReceive Item where entDate<datevalue('" & fromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Issue_Rec='R'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_rec = search1.Fields(0).Value
                        Else
                        pr_rec = 0
                     End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from BookReceive Item where entDate<datevalue('" & fromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Issue_Rec='D'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_issue = search1.Fields(0).Value
                        Else
                        pr_issue = 0
                     End If
                     
                     
                     op = (op + (pr_rec - pr_issue))
                     
                     
                     
                     '-----------------end Code-------------------

                     
                     
                     
                     If search1.State = 1 Then search1.Close
                     'search1.Open "select sum(Qty) from IssuItem where HeatingDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", con
                     search1.Open "select sum(Qty) from BookReceive Item where entDate=datevalue('" & fromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Issue_Rec='R'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                        Else
                        Receive = 0
                     End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     'search1.Open "select sum(Qty) from Invoice where RecDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", con
                     search1.Open "select sum(Qty) from BookReceive Item where entDate=datevalue('" & fromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Issue_Rec='D'", con
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
          
          Set save = New ADODB.Recordset
          If save.State = 1 Then save.Close
          save.Open "select * from FinishStockSummary", con, adOpenDynamic, adLockOptimistic
          save.AddNew
          save!Name = rs.Fields(0).Value
          save!gpname = quality
          ss = rs.Fields(0).Value
          save!Dates = Search.Fields(0).Value
          save!OpenStock = op
          save!Issue = Issue
          save!ReceiveStock = Receive
          save!ClosingStock = ((op + Receive) - (Issue))
          save!heatno = heatno
          save.Update
          
          
          
          If rss.State = 1 Then rss.Close
          rss.Open "select ClosingStock,Aouto  from FinishStockSummary where Name='" & ss & "'  order by Dates,aouto", con
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
              save.Open "select * from FinishStockSummary where Aouto=" & nn & "", con, adOpenDynamic, adLockOptimistic
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
                                rs.Open "select distinct(Name) from FinishStockSummary where GpName='" & quality & "'", con
                                If rs.EOF = False Then
                                While rs.EOF = False
                                
                                                '------------Calculate Opening--------
                                op = 0
                                If oprs.State = 1 Then oprs.Close
                                    oprs.Open "select OpeningStock,Openingdate from ItemMaster where ItemGp='" & quality & "' and ItemName='" & rs.Fields(0).Value & "'", con
                                    If oprs.EOF = False Then
                                        If rss.State = 1 Then rss.Close
                                            rss.Open "select * from FinishStockSummary where GpName='" & quality & "' and Name='" & rs.Fields(0).Value & "'", con
                                            If rss.EOF = True Then
                                                op = oprs.Fields(0).Value
                                            End If
                                            Else
                                        op = 0
                                    End If
                                
                            
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from IssuItem Item where HeatingDate<datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", con
                            If Not IsNull(search1.Fields(0)) Then
                            Receive = search1.Fields(0).Value
                            Else
                            Receive = 0
                            End If
                            
                            
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from Invoice where RecDate<datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", con
                            If Not IsNull(search1.Fields(0)) Then
                            Issue = search1.Fields(0).Value
                            Else
                            Issue = 0
                            End If
           
                                                           
'
'                                    If search1.State = 1 Then search1.Close
'                                    search1.Open "select sum(Qty) from RawPurchaseReturn where invdate<datevalue('" & FromDate.Value & "') and [Group]='" & quality & "' and [Name]='" & rs.Fields(0).Value & "'", con
'                                    If Not IsNull(search1.Fields(0)) Then
'                                    purchaseret = search1.Fields(0).Value
'                                    Else
'                                    purchaseret = 0
'                                    End If
                                                           
                                                           
                                                           
                                                                 
                                         op = ((op + pr_rec) - pr_issue)

                                         If op = 0 Then GoTo ab1:
                                         
                                         Set save = New ADODB.Recordset
                                         If save.State = 1 Then save.Close
                                         save.Open "select * from FinishStockSummary", con, adOpenDynamic, adLockOptimistic
                                         save.AddNew
                                         save!GroupName = quality
                                         save!gpname = quality
                                         save!Name = rs.Fields(0).Value
                                         'save!unit = unit
                                         ss = rs.Fields(0).Value
                                         save!Dates = fromDate.Value
                                         save!OpenStock = op
                                         save!ReceiveStock = Receive
                                         save!Issue = Issue
                                         'save!PurhaseReturn = purchaseret
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
    Dim cmd As New ADODB.Command
    Dim K As Integer
    Dim D As Date
    Set cmd.ActiveConnection = con
    
    cmd.CommandText = "delete from dates"
    cmd.Execute
    K = DateDiff("d", fromDate.Value, todate.Value)
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dates", con, adOpenDynamic, adLockOptimistic
     D = fromDate.Value
     For i = 0 To K
           rs.AddNew
           rs!Dates = D
           rs.Update
           D = D + 1
       Next
    
End Sub

Private Sub cmdPrintAllBook_Click()

'=======================================

Dim rs_1 As New ADODB.Recordset
Dim strVal As String
con.Execute "update ItemMaster set TitleFarmNo=''"
If Check1_withbillNoQty.Value = 1 Then
   
   If rs1.State = 1 Then rs1.Close
   If ss1 = "itc" Then
      rs1.Open "select  BOOKCODE  from INVOICEB where type<>'Book' and type<>'Title' and  bookgp='" & cbogp.Text & "' group by BOOKCODE", con
   ElseIf ss1 = "title" Then
     rs1.Open "select  BOOKCODE  from INVOICEB where  type='Title' and  bookgp='" & cbogp.Text & "' group by BOOKCODE", con
   Else
     rs1.Open "select  BOOKCODE  from INVOICEB where  (type='Book' or type='Stich') and  bookgp='" & cbogp.Text & "' group by BOOKCODE", con
   End If
   
  While rs1.EOF = False
  
  strVal = ""
  
   If ss1 = "itc" Then
   str1 = "SELECT  INVOICEB.SUBLEDGER,INVOICEB.type,INVOICEB.INVOICENO, INVOICEB.bookgp, INVOICEB.BOOKCODE, sum(INVOICEB.qty) as Qty FROM INVOICEB where INVOICEB.type<>'Book' and INVOICEB.type<>'Title'  and INVOICEB.BOOKCODE='" & rs1!BOOKCODE & "'  group by  INVOICEB.SUBLEDGER,INVOICEB.type,INVOICEB.INVOICENO, INVOICEB.bookgp, INVOICEB.BOOKCODE"
   ElseIf ss1 = "title" Then
   str1 = "SELECT  INVOICEB.SUBLEDGER,INVOICEB.type,INVOICEB.INVOICENO, INVOICEB.bookgp, INVOICEB.BOOKCODE, sum(INVOICEB.qty) as Qty FROM INVOICEB where INVOICEB.type='Title'  and INVOICEB.BOOKCODE='" & rs1!BOOKCODE & "'  group by  INVOICEB.SUBLEDGER,INVOICEB.type,INVOICEB.INVOICENO, INVOICEB.bookgp, INVOICEB.BOOKCODE"
   Else
   str1 = "SELECT  INVOICEB.SUBLEDGER,INVOICEB.type,INVOICEB.INVOICENO, INVOICEB.bookgp, INVOICEB.BOOKCODE, sum(INVOICEB.qty) as Qty FROM INVOICEB where (INVOICEB.type='Book' or INVOICEB.type='Stich')  and INVOICEB.BOOKCODE='" & rs1!BOOKCODE & "'  group by  INVOICEB.SUBLEDGER,INVOICEB.type,INVOICEB.INVOICENO, INVOICEB.bookgp, INVOICEB.BOOKCODE"
   End If
   
   If rs.State = 1 Then rs.Close
   rs.Open str1, con
   While rs.EOF = False
      
       If strVal = "" Then
          strVal = "B-" & rs(2) & ",Qt-" & rs(5) & "-" & Mid(rs(0), 1, 1)
       Else
          strVal = strVal & ";B-" & rs(2) & ",Qt:" & rs(5) & "-" & Mid(rs(0), 1, 1)
       End If
      rs.MoveNext
   Wend
   
   con.Execute "update ItemMaster set TitleFarmNo='" & strVal & "' where itemname='" & rs1!BOOKCODE & "'"
   
   rs1.MoveNext
  Wend
   
End If

If MsgBox("want to print ?", vbQuestion + vbYesNo) = vbNo Then
  Exit Sub
End If

'=======================================


If ss1 = "itc" Then
  
   cr.Reset
   cr.ReportFileName = App.Path & "\TaxSummeryITC_gp.rpt"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   If cbogp.Text <> "" Then
      cr.ReplaceSelectionFormula "{BookReceive.itemgp}='" & cbogp.Text & "'"
   End If
   cr.Formulas(2) = "firm_='" & COMPNAME & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.WindowShowRefreshBtn = True
   cr.Action = 1
   
   
ElseIf ss1 = "title" Then
   
   
   tempfortitle
   
   cr.Reset
   cr.ReportFileName = App.Path & "\TaxSummeryTitle_gp.rpt"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   If cbogp.Text <> "" Then
      cr.ReplaceSelectionFormula "{BookReceive.itemgp}='" & cbogp.Text & "'"
   End If
   cr.Formulas(2) = "firm_='" & COMPNAME & "'"
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowShowRefreshBtn = True
   cr.WindowState = crptMaximized
   cr.Action = 1
   
   
Else

   cr.Reset
   cr.ReportFileName = App.Path & "\TaxSummeryNewBook_gp.rpt"
   cr.Formulas(0) = "Fromdate='" & fromDate.Value & "'"
   cr.Formulas(1) = "Todate='" & todate.Value & "'"
   cr.Formulas(2) = "firm_='" & COMPNAME & "'"
   If cbogp.Text <> "" Then
      cr.ReplaceSelectionFormula "{BookReceive.itemgp}='" & cbogp.Text & "'"
   End If
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.WindowShowRefreshBtn = True
   cr.Action = 1


End If




End Sub

Private Sub cmdSearch_Click()
    
   ' If CboPName.Text = "" Then
   '    MsgBox "Please Select ItemName !!", vbCritical
   '    Exit Sub
   ' End If
    
    
    Screen.MousePointer = vbHourglass
    
    AddDate
    Dim opening As Long
    Dim Search As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = con
    cmd.CommandText = "delete from FinishStockSummary"
    cmd.Execute
    
    UpdateStock
    
    fillgrid
    Screen.MousePointer = vbDefault
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
Sub fillgrid()

''    Dim rss As New ADODB.Recordset
''
''    con.Execute "delete from tmpITC where len(BookName)>0"
''
''    'Inner_paper_cover
''
''
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Inner' as inner_paper_cover from ITC where inner_paper_cover='I+P+C'"
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Paper' as inner_paper_cover from ITC where inner_paper_cover='I+P+C'"
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Cover' as inner_paper_cover from ITC where inner_paper_cover='I+P+C'"
''
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Paper' as inner_paper_cover from ITC where inner_paper_cover='P+C'"
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Cover' as inner_paper_cover from ITC where inner_paper_cover='P+C'"
''
''
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Paper' as inner_paper_cover from ITC where inner_paper_cover='I+P'"
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Inner' as inner_paper_cover from ITC where inner_paper_cover='I+P'"
''
''
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Inner' as inner_paper_cover from ITC where inner_paper_cover='I+C'"
''
''     con.Execute "insert into tmpITC  select EntryNo,EntDate,PartyName,BookName,ChallanNo,Qty," & _
''     "Descr,Issue_Rec,'Cover' as inner_paper_cover from ITC where inner_paper_cover='I+C'"
''
    
    
    
    
    
    Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    fill.Open "select Name,dates,Unit,OpenStock,ReceiveStock,Issue,ClosingStock from FinishStockSummary where Dates>=datevalue('" & fromDate.Value & "') and Dates<=datevalue('" & todate.Value & "') order by Dates,Aouto", con
    If fill.EOF = False Then
       'Set vs.DataSource = fill
       vs.FormatString = "Date|ChallanNo.|Description|>OpenStock|>Received Qty|>Delivered Qty|>Balance"
       
       '------------------------------------
    
       vs.Rows = fill.RecordCount + 1
       For i = 1 To fill.RecordCount
         If fill.EOF = False Then
         
            vs.TextMatrix(i, 0) = fill!Dates
            vs.TextMatrix(i, 1) = fill!unit
            vs.TextMatrix(i, 2) = fill!Name
            vs.TextMatrix(i, 3) = fill!OpenStock
            
            vs.TextMatrix(i, 4) = fill!ReceiveStock
            vs.TextMatrix(i, 5) = fill!ClosingStock
          
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

Private Sub Form_Load()


If rs.State = 1 Then rs.Close
rs.Open "select fax from setup1", con
If rs.EOF = False Then

If rs!FAX = "y" Then
   Check1_withbillNoQty.Visible = True
 Else
   Check1_withbillNoQty.Visible = False
End If

End If


todate.Value = Date
fillgrid

cmdPrintAllBook.Visible = True

If ss1 = "new" Then
Me.Caption = "New Book Statement"
ElseIf ss1 = "title" Then
Me.Caption = "Title Book Statement"
'cmdPrintAllBook.Visible = False
ElseIf ss1 = "itc" Then
Me.Caption = "ITC Book Statement"
End If


fillcombo cbogp, "Name", "ConsumeGp", con

'fillcombo CboPName, "ItemName", "ItemMaster", con

fromDate.Value = Date
LoadName
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then todate.SetFocus
End Sub
