VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmMonthlySummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Summary of Raw Material"
   ClientHeight    =   2115
   ClientLeft      =   195
   ClientTop       =   720
   ClientWidth     =   4605
   Icon            =   "frmMonthlySummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4605
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
      Left            =   4920
      TabIndex        =   9
      Top             =   870
      Visible         =   0   'False
      Width           =   465
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
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1485
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   4455
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   840
         TabIndex        =   4
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
         Format          =   24379395
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   2880
         TabIndex        =   5
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
         Format          =   24379395
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   285
         Width           =   615
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
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1050
      Width           =   2100
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
      Height          =   360
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1530
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ComboBox cboGp 
      Height          =   315
      ItemData        =   "frmMonthlySummary.frx":0442
      Left            =   4710
      List            =   "frmMonthlySummary.frx":045B
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   330
   End
   Begin Crystal.CrystalReport CR 
      Left            =   4755
      Top             =   750
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   330
      Left            =   4815
      TabIndex        =   10
      Top             =   240
      Width           =   255
      _cx             =   450
      _cy             =   582
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
      FormatString    =   $"frmMonthlySummary.frx":04CB
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
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
      Left            =   4815
      TabIndex        =   12
      Top             =   450
      Visible         =   0   'False
      Width           =   330
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
      Left            =   4725
      TabIndex        =   11
      Top             =   555
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "frmMonthlySummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboItemName_Click()
  On Error Resume Next
  fillgrid
End Sub
Private Function LoadName()
CboPName.Clear
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(itemname) from ItemMaster where ItemGp='" & cboGp.Text & "'", con, adOpenDynamic, adLockOptimistic
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
    
   Call cmdSearch_Click
   UpdateStock1
   
   CR.Reset
   CR.ReportFileName = App.Path & "/MainReport.rpt"
   CR.Formulas(1) = "dates='" & ToDate.Value & "'"
   
   'If rs.State = 1 Then rs.Close
   'rs.Open "select sum(Purchase) from asonaside where Side='a'", con
   'If Not IsNull(rs(0)) Then
   '  CR.Formulas(2) = "sumdp=" & rs(0) & ""
   'End If
   
   
   
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
Dim Receive
Dim Issue
Dim purchaseret
Dim opening
Dim opDate As Date
Dim pr_rec, pr_issue
Dim despatch
Dim side1


Dim quality As String
Screen.MousePointer = vbHourglass
Dim rs_S As New ADODB.Recordset

    

If Search.State = 1 Then Search.Close
Search.Open "select distinct(itemgp) from itemMaster where itemgp<>'Consumable'", con
If Search.EOF = False Then
   
   
quality = Search.Fields("Itemgp").Value
op = 0
If oprs.State = 1 Then oprs.Close
oprs.Open "select sum(OpeningStock) from itemMaster", con
If oprs.EOF = False Then
   op = oprs.Fields(0).Value
Else
   op = 0
End If
     


'If search1.State = 1 Then search1.Close
'search1.Open "select sum(Qty) from Invoice where HeatNo= 'Null'", con
'If Not IsNull(search1.Fields(0)) Then
'   op = op - search1.Fields(0).Value
'End If





If search1.State = 1 Then search1.Close
search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and Status='Receive' and itemgp<>'Losses'", con
If Not IsNull(search1.Fields(0)) Then
   pr_rec = search1.Fields(0).Value
Else
   pr_rec = 0
End If

If search1.State = 1 Then search1.Close
search1.Open "select sum(Qty) from Purchase where RecDate<datevalue('" & FromDate.Value & "')", con
If Not IsNull(search1.Fields(0)) Then
   Receive = search1.Fields(0).Value
Else
   Receive = 0
End If

pr_rec = pr_rec + Receive
     
     
     
If search1.State = 1 Then search1.Close
search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and Status='Issue'", con
   If Not IsNull(search1.Fields(0)) Then
      pr_issue = search1.Fields(0).Value
   Else
      pr_issue = 0
End If
   
  
  
  
  If search1.State = 1 Then search1.Close
  search1.Open "select sum(Qty) from Invoice where RecDate<datevalue('" & FromDate.Value & "')", con
  If Not IsNull(search1.Fields(0)) Then
     despatch = search1.Fields(0).Value
  Else
     despatch = 0
  End If
     
     
    
     
     
  op = ((op + (pr_rec - pr_issue)) - despatch)
     
     
  '-------- A Side------------------------------------
  '===================================================
  '---------------------------------------------------
  
    Set save = New ADODB.Recordset
    If save.State = 1 Then save.Close
    save.Open "select * from AsOnAside", con, adOpenDynamic, adLockOptimistic
    save.AddNew
    save!Name = "Opening"
    save!gpname = "Opening"
    save!dates = ToDate.Value
    save!Purchase = op
    save!side = "a"
    
    save!head = "Opening"
    save.Update


  End If
  
   
   


'---------------- Code For Purchase-------------------
'=====================================================

  
If rs.State = 1 Then rs.Close
rs.Open "select distinct(ItemName) from Purchase", con
If rs.EOF = False Then

While rs.EOF = False
    
    If rss.State = 1 Then rss.Close
    rss.Open "select distinct(itemgp) from itemMaster where ItemName='" & rs.Fields(0).Value & "'", con
    If rss.EOF = False Then
       quality = rss.Fields(0).Value
    End If
    
    If search1.State = 1 Then search1.Close
    search1.Open "select sum(Qty) from Purchase where (RecDate>=datevalue('" & FromDate.Value & "') and RecDate<=datevalue('" & ToDate.Value & "')) and ItemName='" & rs.Fields(0).Value & "'", con
    If Not IsNull(search1.Fields(0)) Then
        Receive_p = search1.Fields(0).Value
    Else
        Receive_p = 0
    End If
 
    
    If Receive_p <> 0 Then
    
        If rss.State = 1 Then rss.Close
        Set save = New ADODB.Recordset
        If save.State = 1 Then save.Close
        save.Open "select * from AsOnAside", con, adOpenDynamic, adLockOptimistic
        save.AddNew
        save!Name = rs.Fields(0).Value
        save!gpname = quality
        save!dates = ToDate.Value
        save!Purchase = Receive_p
        save!side = "a"
        
        save!head = "Purchase"
        save.Update
    
    End If
    rs.MoveNext
   
   Wend
  End If
   

'-------------- end Purchase-------------------------
        
                  
                  
                  
                  
'---------------- Code For Invoice-------------------
'=====================================================
                  
                  

'''If rs.State = 1 Then rs.Close
'''rs.Open "select distinct(ItemName) from Invoice where (RecDate>=datevalue('" & FromDate.Value & "') and RecDate<=datevalue('" & ToDate.Value & "'))", con
'''If rs.EOF = False Then
'''While rs.EOF = False
'''
'''    If rss.State = 1 Then rss.Close
'''    rss.Open "select distinct(itemgp) from itemMaster where ItemName='" & rs.Fields(0).Value & "'", con
'''    If rss.EOF = False Then
'''       quality = rss.Fields(0).Value
'''    End If
'''
'''
'''     If search1.State = 1 Then search1.Close
'''     search1.Open "select sum(Qty) from Invoice where (RecDate>=datevalue('" & FromDate.Value & "') and RecDate<=datevalue('" & ToDate.Value & "')) and ItemName='" & rs.Fields(0).Value & "'", con
'''     If Not IsNull(search1.Fields(0)) Then
'''        despatch = search1.Fields(0).Value
'''     Else
'''        despatch = 0
'''     End If
'''
'''     If despatch <> 0 Then
'''
'''     Set save = New ADODB.Recordset
'''     If save.State = 1 Then save.Close
'''     save.Open "select * from AsOnAside", con, adOpenDynamic, adLockOptimistic
'''     save.AddNew
'''     save!Name = rs.Fields(0).Value
'''     save!gpname = quality
'''     save!dates = ToDate.Value
'''     save!Purchase = despatch
'''     save!side = "b"
'''     save!head = "invoice"
'''     save.Update
'''
'''    End If
'''
'''
'''rs.MoveNext
'''
'''Wend
'''
'''End If
'''
''''------------End Invoice-----------------------------------------------------------
'''
'''
'''If rs.State = 1 Then rs.Close
'''rs.Open "select distinct(ItemName) from issuitem where (Heatingdate >=datevalue('" & FromDate.Value & "') and Heatingdate<=datevalue('" & ToDate.Value & "')) and itemgp<>'Losses'", con
'''If rs.EOF = False Then
'''
'''While rs.EOF = False
'''
'''
'''
'''
'''                If rss.State = 1 Then rss.Close
'''                rss.Open "select distinct(itemgp) from itemMaster where ItemName='" & rs.Fields(0).Value & "'", con
'''                If rss.EOF = False Then
'''                   quality = rss.Fields(0).Value
'''                End If
'''
'''
'''
'''                 If search1.State = 1 Then search1.Close
'''                 search1.Open "select sum(Qty) from IssuItem where (HeatingDate>=datevalue('" & FromDate.Value & "') and HeatingDate<=datevalue('" & ToDate.Value & "'))and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
'''                 If Not IsNull(search1.Fields(0)) Then
'''                   Receive = search1.Fields(0).Value
'''                 Else
'''                   Receive = 0
'''                 End If
'''
'''
'''
'''
'''                 If Receive <> 0 Then
'''
'''                    Set save = New ADODB.Recordset
'''                    If save.State = 1 Then save.Close
'''                    save.Open "select * from AsOnAside", con, adOpenDynamic, adLockOptimistic
'''                    save.AddNew
'''                    save!Name = rs.Fields(0).Value
'''                    save!gpname = quality
'''                    save!dates = ToDate.Value
'''                    save!Purchase = Receive
'''                    save!side = "b"
'''                    save!head = "Stock"
'''                    save.Update
'''
'''                End If
'''
'''
'''
'''              rs.MoveNext
'''              Wend
'''
'''              End If
'''
'''
'''
'''
'''
'''             '------------Losses---------------------------------
'''
'''If rs.State = 1 Then rs.Close
'''rs.Open "select distinct(ItemName) from issuitem where (Heatingdate >=datevalue('" & FromDate.Value & "') and Heatingdate<=datevalue('" & ToDate.Value & "')) and itemgp='Losses'", con
'''If rs.EOF = False Then
'''
'''While rs.EOF = False
'''
'''
'''
'''
'''                If rss.State = 1 Then rss.Close
'''                rss.Open "select distinct(itemgp) from itemMaster where ItemName='" & rs.Fields(0).Value & "'", con
'''                If rss.EOF = False Then
'''                   quality = rss.Fields(0).Value
'''                End If
'''
'''
'''
'''                 If search1.State = 1 Then search1.Close
'''                 search1.Open "select sum(Qty) from IssuItem where (HeatingDate>=datevalue('" & FromDate.Value & "') and HeatingDate<=datevalue('" & ToDate.Value & "'))and ItemName='" & rs.Fields(0).Value & "' and Status='Receive' and itemgp='Losses'", con
'''                 If Not IsNull(search1.Fields(0)) Then
'''                    Receive = search1.Fields(0).Value
'''                 Else
'''                    Receive = 0
'''                 End If
'''
'''
'''                 If Receive <> 0 Then
'''
'''                    Set save = New ADODB.Recordset
'''                    If save.State = 1 Then save.Close
'''                    save.Open "select * from AsOnAside", con, adOpenDynamic, adLockOptimistic
'''                    save.AddNew
'''                    save!Name = rs.Fields(0).Value
'''                    save!gpname = quality
'''                    save!dates = ToDate.Value
'''                    save!Purchase = Receive
'''                    save!side = "b"
'''                    save!head = "Losses"
'''                    save.Update
'''
'''                End If
'''
'''
'''              rs.MoveNext
'''              Wend
'''
'''              End If
'''
'''
'''
'''    '----------------------------------------------
              
            
    
    Screen.MousePointer = vbDefault

End Sub
Sub AddDate()
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim k As Integer
    Dim d As Date
    Set cmd.ActiveConnection = con
    
    cmd.CommandText = "delete from dates"
    cmd.Execute
    k = DateDiff("d", FromDate.Value, ToDate.Value)
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dates", con, adOpenDynamic, adLockOptimistic
     d = FromDate.Value
     For i = 0 To k
           rs.AddNew
           rs!dates = d
           rs.Update
           d = d + 1
       Next
    
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
    cmd.CommandText = "delete from AsOnAside"
    cmd.Execute
    
    UpdateStock
    
    fillgrid
   
    Screen.MousePointer = vbDefault
End Sub
Sub UpdateStock1()
     
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
     Dim despatch
     Dim side1
     
     
Dim quality As String


Screen.MousePointer = vbHourglass
      
         
         
         
Dim rs_S As New ADODB.Recordset

    

If Search.State = 1 Then Search.Close
Search.Open "select distinct(itemgp) from itemMaster", con

If Search.EOF = False Then

        

   While Search.EOF = False
             
            quality = Search.Fields("Itemgp").Value
             
             
             
             
             If rs.State = 1 Then rs.Close
             rs.Open "select distinct(ItemName) from itemMaster where ItemGp='" & quality & "'", con
           
             If rs.EOF = False Then
             While rs.EOF = False
                      '------------Calculate Opening--------
                      
                      
                      
                      
                      op = 0
                      If oprs.State = 1 Then oprs.Close
                      oprs.Open "select OpeningStock,Openingdate from itemMaster where ItemGp='" & quality & "' and itemName='" & rs.Fields(0).Value & "'", con
                      If oprs.EOF = False Then
                            'If rss.State = 1 Then rss.Close
                            'rss.Open "select * from AsOnAside where Name='" & rs.Fields(0).Value & "'", con
                            'If rss.EOF = True Then
                               op = oprs.Fields(0).Value
                            'End If
                        Else
                            op = 0
                       End If
                     
                     
                     
                     
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_rec = search1.Fields(0).Value
                        Else
                        pr_rec = 0
                     End If
                  
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from Purchase where RecDate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                     Else
                        Receive = 0
                     End If
                     
                   pr_rec = pr_rec + Receive
                     
                     
                  If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Issue'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_issue = search1.Fields(0).Value
                        Else
                        pr_issue = 0
                  End If
                     
                  If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where Heatingdate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "' and Status='Issue'", con
                     If Not IsNull(search1.Fields(0)) Then
                        pr_issue = search1.Fields(0).Value
                        Else
                        pr_issue = 0
                  End If
                     
                  If search1.State = 1 Then search1.Close
                  search1.Open "select sum(Qty) from Invoice where RecDate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "'", con
                  If Not IsNull(search1.Fields(0)) Then
                     despatch = search1.Fields(0).Value
                  Else
                     despatch = 0
                  End If
                  
                     
                    
                     
                     
                  op = ((op + (pr_rec - pr_issue)) - despatch)
                     
                  'If op > 1 Then
                  '   MsgBox "sdsd"
                  'End If
                     
                     '-----------------end Code-------------------

                     
                  'If quality <> "Raw Item" Or UCase(rs.Fields(0).Value) = "MA" Then
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where (HeatingDate>=datevalue('" & FromDate.Value & "') and HeatingDate<=datevalue('" & ToDate.Value & "'))and ItemName='" & rs.Fields(0).Value & "' and Status='Receive'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                        Else
                        Receive = 0
                     End If
                 'Else
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from Purchase where (RecDate>=datevalue('" & FromDate.Value & "') and RecDate<=datevalue('" & ToDate.Value & "')) and ItemName='" & rs.Fields(0).Value & "'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Receive_p = search1.Fields(0).Value
                        Else
                        Receive_p = 0
                     End If
                  'End If
                  
                     
                   Receive = Receive + Receive_p
                     
                     
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssuItem where (HeatingDate>=datevalue('" & FromDate.Value & "') and HeatingDate<=datevalue('" & ToDate.Value & "')) and ItemName='" & rs.Fields(0).Value & "' and Status='Issue'", con
                     If Not IsNull(search1.Fields(0)) Then
                        Issue = search1.Fields(0).Value
                        Else
                        Issue = 0
                     End If
               
                     
                  If search1.State = 1 Then search1.Close
                  search1.Open "select sum(Qty) from Invoice where (RecDate>=datevalue('" & FromDate.Value & "') and RecDate<=datevalue('" & ToDate.Value & "')) and ItemName='" & rs.Fields(0).Value & "'", con
                  If Not IsNull(search1.Fields(0)) Then
                     despatch = search1.Fields(0).Value
                  Else
                     despatch = 0
                  End If
   

                     
        
       
        
        '--------------------- Data Fiter and Now  Save Coding========================
         opening = 0
         If Issue = 0 And Receive = 0 And despatch = 0 And op = 0 Then
            GoTo aaa:
         End If
          
         If (op + ((Receive - Issue) - despatch)) <> 0 Then
          Set save = New ADODB.Recordset
          If save.State = 1 Then save.Close
          save.Open "select * from AsOnAside", con, adOpenDynamic, adLockOptimistic
          save.AddNew
          save!Name = rs.Fields(0).Value
          save!gpname = quality
          save!dates = ToDate.Value
          If quality = "Losses" Then
            If (((Receive - Issue) - despatch)) > 0 Then
             save!Purchase = (((Receive - Issue) - despatch))
             save!side = "b"
             save.Update
            End If
        
          Else
           If (op + ((Receive - Issue) - despatch)) > 0 Then
             save!Purchase = (op + ((Receive - Issue) - despatch))
             save!side = "b"
             save.Update
            End If
          End If
         End If
          
          If despatch > 0 Then
            Set save = New ADODB.Recordset
            If save.State = 1 Then save.Close
            save.Open "select * from AsOnAside", con, adOpenDynamic, adLockOptimistic
            save.AddNew
            save!Name = rs.Fields(0).Value
            save!gpname = "Dispatch"
            save!dates = ToDate.Value
            save!Purchase = despatch
            save!side = "b"
            save.Update
           
          End If
          
aaa:
          
      rs.MoveNext
     
     Wend
                                  
 End If
 
 Search.MoveNext
 
 Wend
                                   
 
 End If
  
   
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




cboGp.ListIndex = 0

End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then ToDate.SetFocus
End Sub

