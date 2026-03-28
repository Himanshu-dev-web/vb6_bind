VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmWeatorWiseSale 
   Caption         =   "Weator Wise Sale"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   Icon            =   "frmWeatorWiseSale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin Crystal.CrystalReport cr 
      Left            =   360
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker FromDate 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   675
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   24510467
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   24510467
      UpDown          =   -1  'True
      CurrentDate     =   37701
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmWeatorWiseSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbostaff_Click()
  On Error Resume Next
  Dim ss As New ADODB.Recordset
  If ss.State = 1 Then ss.Close
  ss.Open "select BrokerName from staff where Brokercode=" & cbostaff.Text & "", con
  If ss.EOF = False Then
     name1.Caption = ss.Fields(0).Value
  End If

End Sub

Private Sub cmdPrint_Click()
   cr.ReportFileName = App.Path & "\WeatorWiseSale.rpt"
   cr.Connect = "FILEDSN=hotel;pwd=java;"
   cr.ReplaceSelectionFormula "{kot.date}>=datevalue('" & Format(FromDate.Value, "MM/dd/yy") & "') and {kot.date}<=datevalue('" & Format(todate.Value, "MM/dd/yy") & "') and {kot.Complimenty}=False"
   cr.Formulas(0) = "fromdate='" & FromDate.Value & "'"
   cr.Formulas(1) = "todate='" & todate.Value & "'"
   cr.WindowState = crptMaximized
   cr.Action = 1
End Sub
'Sub AddWeator()
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from staff where Des='Waitor'", con
'    cbostaff.Clear
'    If rs.EOF = False Then
'       While rs.EOF = False
'          cbostaff.AddItem rs(0)
'          rs.MoveNext
'       Wend
'    End If
'End Sub
Private Sub Form_Load()
   On Error Resume Next
   FromDate.Value = Date
   todate.Value = Date
   frmWeatorWiseSale.Left = 4000
   frmWeatorWiseSale.Top = 1000
   frmWeatorWiseSale.BackColor = &HDEFEF8
   'AddWeator
   'cbostaff.ListIndex = -1
End Sub


