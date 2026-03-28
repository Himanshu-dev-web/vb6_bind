VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmItemMaster 
   BackColor       =   &H00C0C0C0&
   Caption         =   " Item Msater                                      "
   ClientHeight    =   8484
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   15840
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItemMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8484
   ScaleWidth      =   15840
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtHSNCODE 
      Height          =   285
      Left            =   5895
      TabIndex        =   31
      Top             =   1215
      Width           =   1320
   End
   Begin VB.TextBox txtTitleFarm 
      Height          =   315
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   4560
      TabIndex        =   23
      Top             =   1800
      Width           =   4755
      Begin VB.TextBox txtTitleRate 
         Height          =   315
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   29
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "Rate :"
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
         Index           =   3
         Left            =   900
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1035
      Left            =   135
      TabIndex        =   22
      Top             =   1620
      Width           =   4095
      Begin VB.TextBox txtBookRate 
         Height          =   315
         Left            =   2100
         MaxLength       =   15
         TabIndex        =   28
         Top             =   660
         Width           =   1395
      End
      Begin VB.TextBox txtFarmNo 
         Height          =   315
         Left            =   2100
         MaxLength       =   15
         TabIndex        =   25
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "Rate :"
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
         Index           =   1
         Left            =   840
         TabIndex        =   26
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Farm No :"
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
         Index           =   0
         Left            =   840
         TabIndex        =   24
         Top             =   300
         Width           =   1215
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5640
      Left            =   120
      TabIndex        =   19
      Top             =   2700
      Width           =   15585
      _cx             =   27490
      _cy             =   9948
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
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmItemMaster.frx":0442
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
   Begin VB.TextBox txtOrder 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13605
      MaxLength       =   15
      TabIndex        =   6
      Text            =   "0"
      Top             =   2190
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtOpening 
      Height          =   315
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "0"
      Top             =   840
      Width           =   1275
   End
   Begin VB.TextBox txtItemName 
      Height          =   315
      Left            =   1440
      MaxLength       =   150
      TabIndex        =   1
      Top             =   480
      Width           =   5775
   End
   Begin VB.ComboBox cboGp 
      Height          =   315
      ItemData        =   "frmItemMaster.frx":04DC
      Left            =   1440
      List            =   "frmItemMaster.frx":04DE
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5790
   End
   Begin VB.ComboBox cboUnit 
      Height          =   315
      ItemData        =   "frmItemMaster.frx":04E0
      Left            =   13380
      List            =   "frmItemMaster.frx":04E7
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2565
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13620
      MaxLength       =   15
      TabIndex        =   3
      Text            =   "0"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   885
      Left            =   7815
      TabIndex        =   16
      Top             =   60
      Width           =   5730
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   615
         Left            =   3675
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   165
         Width           =   1050
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Height          =   615
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   165
         Width           =   825
      End
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   615
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   615
         Left            =   4740
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   165
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   615
         Left            =   1005
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   165
         Width           =   915
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   615
         Left            =   2745
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   165
         Width           =   930
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "HSN Code"
      Height          =   285
      Left            =   5940
      TabIndex        =   32
      Top             =   945
      Width           =   1230
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
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
      Index           =   2
      Left            =   120
      TabIndex        =   30
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label id 
      Height          =   285
      Left            =   14940
      TabIndex        =   21
      Top             =   1470
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Order Level"
      Height          =   285
      Left            =   12960
      TabIndex        =   18
      Top             =   2205
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Stock "
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Rate"
      Height          =   285
      Left            =   13020
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Unit"
      Height          =   285
      Left            =   13020
      TabIndex        =   14
      Top             =   2565
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Group "
      Height          =   285
      Left            =   45
      TabIndex        =   13
      Top             =   90
      Width           =   1650
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Item Name "
      Height          =   285
      Left            =   45
      TabIndex        =   7
      Top             =   450
      Width           =   1650
   End
End
Attribute VB_Name = "frmItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemdes As String
Dim b As Boolean

Private Sub cboGp_Click()
fillgrid
End Sub

Private Sub cbogp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtItemName.SetFocus
End Sub

Private Sub cmdDel_Click()
    If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
       DeleteData
       Call cmdRef_Click
       fillgrid
    End If
End Sub
Sub UpdateTransiction()
    
con.Execute "update BookReceive set BookName='" & txtItemName.Text & "',bGp='" & cbogp.Text & "' where BookName='" & itemdes & "'"
con.Execute "update ITC set BookName='" & txtItemName.Text & "' where BookName='" & itemdes & "'"
con.Execute "update Title set BookName='" & txtItemName.Text & "' where BookName='" & itemdes & "'"
con.Execute "update Titledel set BookName='" & txtItemName.Text & "' where BookName='" & itemdes & "'"
con.Execute "update INVOICEB set BOOKCODE='" & txtItemName.Text & "',bookgp='" & cbogp.Text & "' where BOOKCODE='" & itemdes & "'"

con.Execute "update titleStatent set bookname='" & txtItemName.Text & "' where BOOKname='" & itemdes & "'"


   
End Sub
Sub DeleteData()
   
Set cmd.ActiveConnection = con
cmd.CommandText = "delete from ItemMaster where Aouto=" & id.Caption & ""
cmd.Execute
   
End Sub
Private Sub cmdMain_Click()
Unload Me
End Sub
Sub SetWidth()
    vs.Cols = 9
    'vs.FormatString = "|Item Name|<ItemGp|Opening|BookFarmNo|BookRate|TitleFarmNo|TitleRate"
    vs.FormatString = "|Item Name|<ItemGp|Opening|BookFarmNo|BookRate|Description|TitleRate|HSNCode"
    vs.ColWidth(0) = 0
    vs.ColWidth(1) = 4500
    vs.ColWidth(2) = 3100
    vs.ColWidth(3) = 1000
    vs.ColWidth(4) = 1000
    vs.ColWidth(5) = 1000
    vs.ColWidth(6) = 1100
    vs.ColWidth(7) = 1000
    vs.ColWidth(8) = 900
    
End Sub

Private Sub cmdModify_Click()
        
    If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
       DeleteData
       SaveData
       UpdateTransiction
       Call cmdRef_Click
       fillgrid
    End If
End Sub

Private Sub CmdPrint_Click()
  frmMain.cr.Reset
  frmMain.cr.ReportFileName = App.Path & "/BookList.rpt"
  If (cbogp.Text <> "") Then
  frmMain.cr.ReplaceSelectionFormula "{ItemMaster.ItemGp}='" & cbogp.Text & "'"
  'Else
  '  frmMain.CR.ReplaceSelectionFormula "{ItemMaster.OpeningStock}<>0"
  End If
  
  setReportOption
End Sub

Private Sub cmdRef_Click()
    'cboGp.ListIndex = -1
    'cboUnit.ListIndex = -1
    txtItemName.Text = ""
    txtRate.Text = "0"
    txtOpening.Text = 0
    txtOrder.Text = 0
    txtHSNCODE.Text = ""
    cmdDel.Enabled = False
    cmdModify.Enabled = False
    cmdSave.Enabled = True
       
    cbogp.SetFocus
    
End Sub
Private Sub cmdSave_Click()
    
    If rs.State = 1 Then rs.Close
    'rs.Open "select * from ItemMaster where ItemGp='" & cboGp.Text & "' and ItemName='" & txtItemName.Text & "'", con, adOpenDynamic, adLockOptimistic
    
    rs.Open "select * from ItemMaster where ItemName='" & txtItemName.Text & "'", con, adOpenDynamic, adLockOptimistic
    
    If rs.EOF = True Then
       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
          SaveData
          Call cmdRef_Click
       End If
    Else
       MsgBox "This Item is Already Exist !!", vbCritical
       Exit Sub
    End If
    
End Sub
Sub fillgrid()
    Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    If cbogp.Text = "" Then
       fill.Open "select * from ItemMaster order by ItemGp,itemname", con, adOpenDynamic, adLockOptimistic
    Else
       
    If cbogp.Text <> "All" Then
       fill.Open "select * from ItemMaster where ItemGp='" & cbogp.Text & "'  order by ItemGp,itemname", con, adOpenDynamic, adLockOptimistic
       Else
       fill.Open "select * from ItemMaster order by ItemGp,itemname", con, adOpenDynamic, adLockOptimistic
    End If
    
    End If
    
    vs.Cols = 9
    
    If fill.EOF = False Then
        vs.Rows = fill.RecordCount + 1
        For i = 1 To fill.RecordCount
           If fill.EOF = False Then
                vs.TextMatrix(i, 0) = fill.Fields("ItemGp").Value
                vs.TextMatrix(i, 1) = fill.Fields("ItemName").Value
                vs.TextMatrix(i, 2) = fill.Fields("ItemGp").Value
                vs.TextMatrix(i, 3) = fill.Fields("OpeningStock").Value
                
                vs.TextMatrix(i, 4) = fill.Fields("BookFarmNo").Value & ""
                vs.TextMatrix(i, 5) = fill.Fields("BookRate").Value & ""
                'vs.TextMatrix(i, 6) = fill.Fields("TitleFarmNo").Value & ""
                vs.TextMatrix(i, 6) = fill.Fields("remarks").Value & ""
                vs.TextMatrix(i, 7) = fill.Fields("TitleRate").Value & ""
                vs.TextMatrix(i, 8) = fill.Fields("hsncode").Value & ""
                
                fill.MoveNext
            End If
        Next
     
    End If
    SetWidth
    
End Sub
Sub SaveData()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from ItemMaster", con, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!itemgp = cbogp.Text
    rs!itemname = txtItemName.Text
    rs!unit = cboUnit.Text
    rs!Rate = txtRate.Text
    rs!OpeningStock = txtOpening.Text
    rs!OrderLevel = txtOrder.Text
    
    rs!BookFarmNo = Val(txtFarmNo.Text)
    rs!BookRate = Val(txtBookRate.Text)
    
    rs!remarks = Trim(txtTitleFarm.Text)
    rs!TitleRate = Val(txtTitleRate.Text)
    
    rs!hsncode = txtHSNCODE.Text
    
    rs.Update
    fillgrid
End Sub
Sub searchData()
    On Error Resume Next
    If rs.State = 1 Then rs.Close
    rs.Open "select * from ItemMaster where ItemGp='" & vs.TextMatrix(vs.RowSel, 0) & "' and ItemName='" & vs.TextMatrix(vs.RowSel, 1) & "'", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
       cbogp.Text = rs!itemgp
       txtItemName.Text = rs!itemname
       itemdes = rs!itemname
       cboUnit.Text = rs!unit
       txtRate.Text = Format(rs!Rate, "#,###.00")
       txtOpening.Text = rs!OpeningStock
       txtOrder.Text = rs!OrderLevel
       
       'txtFarmNo.Text = rs!bookfarmno & ""
       txtTitleFarm.Text = rs!remarks & ""
       
       txtBookRate.Text = rs!BookRate & ""
        
       txtFarmNo.Text = rs!BookFarmNo & ""
       txtTitleRate.Text = rs!TitleRate & ""

       txtHSNCODE.Text = rs!hsncode & ""
       
       
       id.Caption = rs.Fields("aouto").Value
       cmdDel.Enabled = True
       cmdModify.Enabled = True
       cmdSave.Enabled = False
       
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      'SendKeys "{tab}"
      
   End If
End Sub

Private Sub Form_Load()
  SetWidth
  If rs.State = 1 Then rs.Close
  rs.Open "select * from UnitMaster order by Name", con
  If rs.EOF = False Then
     While rs.EOF = False
         cboUnit.AddItem rs.Fields("Name").Value
         rs.MoveNext
     Wend
  End If
  
  fillgrid
  cboUnit.ListIndex = 0
  
  cmdModify.Enabled = False
  cmdDel.Enabled = False
  
  
    
  If rs.State = 1 Then rs.Close
  rs.Open "select * from ConsumeGp order by Name", con
  If rs.EOF = False Then
     While rs.EOF = False
         cbogp.AddItem rs.Fields("Name").Value
         rs.MoveNext
     Wend
  End If
    
    
  pos Me
  
End Sub

Private Sub txtBookRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtTitleRate.SetFocus

End Sub

Private Sub txtFarmNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtBookRate.SetFocus

End Sub

Private Sub txtItemName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtOpening.SetFocus
End Sub

Private Sub txtOpening_GotFocus()
    txtOpening.SelLength = 10
End Sub

Private Sub txtOpening_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtTitleFarm.SetFocus
End Sub

Private Sub txtOpening_KeyPress(KeyAscii As Integer)
   ' b = val_int(txtOpening, KeyAscii)
    'If b = False Then KeyAscii = 0
End Sub

Private Sub txtOrder_GotFocus()
txtOrder.SelLength = 10
End Sub
Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    b = val_int(txtOrder, KeyAscii)
    If b = False Then KeyAscii = 0
End Sub

Private Sub txtRate_GotFocus()
   txtRate.SelLength = 10
End Sub
Private Sub txtRate_KeyPress(KeyAscii As Integer)
    b = val_int(txtRate, KeyAscii)
    If b = False Then KeyAscii = 0
End Sub

Private Sub txtTitleFarm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtFarmNo.SetFocus
End Sub

Private Sub txtTitleRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSave.SetFocus
End Sub

Private Sub vs_Click()
   searchData
End Sub
Private Sub vs_SelChange()
   searchData
End Sub
