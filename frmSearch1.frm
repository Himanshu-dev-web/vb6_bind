VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearch1 
   Caption         =   "Summary of Stock of Raw Item"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   Icon            =   "frmSearch1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotal2 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   10350
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   7320
      Width           =   1365
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   7320
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   345
      Left            =   8610
      TabIndex        =   5
      Top             =   210
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   7200
      TabIndex        =   4
      Top             =   210
      Width           =   1365
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   345
      Left            =   5820
      TabIndex        =   3
      Top             =   210
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker dates 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   180
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38740
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6555
      Left            =   270
      TabIndex        =   2
      Top             =   690
      Width           =   11520
      _cx             =   20320
      _cy             =   11562
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   3000
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSearch1.frx":0442
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
   Begin MSComCtl2.DTPicker toDate 
      Height          =   315
      Left            =   4140
      TabIndex        =   9
      Top             =   210
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38740
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3150
      TabIndex        =   10
      Top             =   210
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   315
      Left            =   4140
      TabIndex        =   7
      Top             =   7350
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   1305
   End
End
Attribute VB_Name = "frmSearch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdView_Click()
   Dim op As Double
   Dim rss As New ADODB.Recordset
   Dim cmd As New ADODB.Command
   Dim save As New ADODB.Recordset
   Dim item As String
   Dim wt, wt1 As Double
   
   '-------------------
   wt = 0
   wt1 = 0

   op = 0
   '-------------------
   
   
   Screen.MousePointer = vbHourglass
   Set cmd.ActiveConnection = con
   
   cmd.CommandText = "delete from StockSummaryAsOn"
   cmd.Execute
   
   
   
   
   
   If save.State = 1 Then save.Close
   save.Open "select * from StockSummaryAsOn", con, adOpenDynamic, adLockOptimistic
   
   
   
   
   '----------- For Openig
   Dim rs5 As New ADODB.Recordset
   
   If rs5.State = 1 Then rs5.Close
   rs5.Open "select distinct(ItemGp) from OPStock order by ItemGp", con
   
   While rs5.EOF = False
   
              
        save.AddNew
        save.Fields("Name").Value = rs5.Fields(0).Value
        save.Fields("wt").Value = 0
        save.Fields("status").Value = "g"
        
        save.Update
           
    
        
        
        If rss.State = 1 Then rss.Close
        rss.Open "select OpeningStock,ItemName from OPStock where ItemGp='" & rs5.Fields(0).Value & "'", con
        While rss.EOF = False
             
            If "cu wire (satna)" = rss.Fields(1).Value Then
               MsgBox "f" & rss.Fields(1).Value
            End If
            
            op = rss.Fields(0).Value
            save.AddNew
            save.Fields("Name").Value = rss.Fields(1).Value
            save.Fields("wt").Value = rss.Fields(0).Value
            save.Fields("status").Value = "a"
            save.Fields("gp1").Value = rs5.Fields(0).Value
            save.Update
            
            rss.MoveNext
             
        Wend
       
  
    
    rs5.MoveNext
    Wend
    
    '---------For Receipt-----------------
   
   
   If rs.State = 1 Then rs.Close
   rs.Open "select itemname,total,itemgp from RawItemPurchsae where RecDate<=datevalue('" & dates.Value & "')", con
   If rs.EOF = False Then
      
      While rs.EOF = False
          
          save.AddNew
          save.Fields("Name").Value = rs.Fields("ItemName").Value
          save.Fields("wt").Value = rs.Fields("total").Value
          save.Fields("gp1").Value = rs.Fields("itemgp").Value
          save.Fields("status").Value = "g"
          save.Update
          
          
          rs.MoveNext
      Wend
      
   End If
   
   
 ''******************************************************************************************************************
 ''******************************************************************************************************************
 
 
 '---------For Dispatch & Losses-----------------
   
          If rs.State = 1 Then rs.Close
          rs.Open "select sum(Qty) from invoice where RecDate<=datevalue('" & dates.Value & "')", con
          If Not IsNull(rs(0)) Then
           save.AddNew
           save.Fields("Name").Value = "Dispached"
           save.Fields("wt").Value = rs(0)
           save.Fields("status").Value = "b"
           save.Update
           
          End If
         
        
    '-------- For Losses
   
     If rss.State = 1 Then rss.Close
     rss.Open "select distinct(itemname) from LossesItem", con
     If rss.EOF = False Then
     While rss.EOF = False
   
          If rs.State = 1 Then rs.Close
          rs.Open "select sum(total) from LossesItem where heatingDate<=datevalue('" & dates.Value & "') and itemname='" & rss.Fields("ItemName").Value & "'", con
          If Not IsNull(rs(0)) Then
           
           save.AddNew
           save.Fields("Name").Value = rss.Fields("ItemName").Value
           save.Fields("wt").Value = rs(0)
           save.Fields("status").Value = "b"
           save.Update
           
           
          End If
          rss.MoveNext
       
    Wend
    End If
   
   
   '-------- Stoch In Hand----
       
    Dim dd As Long
    dd = 0
    
    If rs.State = 1 Then rs.Close
    rs.Open "select sum(qty) from StockInHand where HeatingDate<=datevalue('" & dates.Value & "') and itemgp='Losses'", con
    If rs.EOF = False Then
       dd = rs.Fields(0).Value
    End If

    If rs.State = 1 Then rs.Close
    rs.Open "select sum(qty) from StockInHand where HeatingDate<=datevalue('" & dates.Value & "')", con
     
    If Not IsNull(rs(0)) Then
     
       save.AddNew
       save.Fields("Name").Value = "Stock In Hand"
       save.Fields("wt").Value = (rs.Fields(0).Value - dd)
       save.Fields("status").Value = "b"
       save.Update
     
     
     
    End If
 
   
   
  
  
   
   vs.Rows = 1
   If rs.State = 1 Then rs.Close
   rs.Open "select * from StockSummaryAsOn", con, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      For i = 1 To rs.RecordCount
        vs.Rows = vs.Rows + 1
        If rs.Fields("status") = "a" Or rs.Fields("status") = "g" Then
           vs.TextMatrix(i, 0) = rs.Fields(0).Value
           vs.TextMatrix(i, 1) = rs.Fields(1).Value
           If rs.Fields("status") = "g" Then
              vs.Cell(flexcpForeColor, i, 0) = &HFF&
              vs.Cell(flexcpBackColor, i, 0) = &HD3EB83
              vs.Cell(flexcpForeColor, i, 1) = &HFF&
              vs.Cell(flexcpBackColor, i, 1) = &HD3EB83
              vs.Cell(flexcpFontBold, i, 0) = True
              vs.TextMatrix(i, 1) = ""
           End If
            
           wt = wt + rs(1)
        End If
      rs.MoveNext
      Next
   End If
     
'--- b side

   
   If rs.State = 1 Then rs.Close
   rs.Open "select * from StockSummaryAsOn where Status='b'", con, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      For i = 1 To rs.RecordCount
           vs.TextMatrix(i, 2) = rs.Fields(0).Value
           vs.TextMatrix(i, 3) = rs.Fields(1).Value
           wt1 = wt1 + rs(1)
      rs.MoveNext
      
      Next
   End If
   
  
   
   SetHead
   
   txtTotal1.Text = 0
   txtTotal2.Text = 0
    
   For i = 1 To vs.Rows - 1
     
     If vs.TextMatrix(i, 1) = "" Then
        'vs.TextMatrix(i, 1) = 0
     End If
     
     If vs.TextMatrix(i, 3) = "" Then
        vs.TextMatrix(i, 3) = 0
     End If
     
     
     If vs.TextMatrix(i, 0) <> "" Then
        If (vs.TextMatrix(i, 1)) <> "" Then
           txtTotal1.Text = Format(CDbl(txtTotal1.Text) + CDbl(vs.TextMatrix(i, 1)), "0.000")
        End If
        
        txtTotal2.Text = Format(CDbl(txtTotal2.Text) + CDbl(vs.TextMatrix(i, 3)), "0.000")
     End If
       
   Next
   Screen.MousePointer = vbDefault
   

End Sub
Sub SetHead()
    vs.FormatString = "Receipt|Weight|Issue|Weight"
End Sub

Private Sub Form_Load()
  dates.Value = Date
End Sub
