VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSearchDispatche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Searching"
   ClientHeight    =   6750
   ClientLeft      =   2265
   ClientTop       =   1080
   ClientWidth     =   9480
   Icon            =   "frmSearchDispatche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9480
   Begin MSComctlLib.ListView ListView1 
      Height          =   6690
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   11800
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "Item Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Text            =   "Heating No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Text            =   "Geade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "d"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "f"
         Text            =   "Balance Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unit"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmSearchDispatche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ct
Dim k As Integer

Private Sub Form_Load()
k = 1
End Sub

Private Sub ListView1_Click()
 
 On Error Resume Next
 
 Dim litem
 
   For i = 1 To frmSales.vs.Rows - 1
       If frmSales.vs.TextMatrix(i, 0) <> "" Then
          
          If frmSales.vs.TextMatrix(i, 0) = ListView1.SelectedItem.Text Then
             MsgBox "Item Already Exist !!", vbCritical
             Exit Sub
          End If
          
       End If
   Next
   
   
   ct = ListView1.ColumnHeaders.Count
   If ct >= 1 Then frmSales.vs.TextMatrix(k, 0) = ListView1.SelectedItem.Text
   If ct >= 2 Then frmSales.txtHeatingNo.Text = ListView1.SelectedItem.SubItems(1)
   If ct >= 3 Then frmSales.vs.TextMatrix(k, 2) = ListView1.SelectedItem.SubItems(2)
   If ct >= 4 Then frmSales.vs.TextMatrix(k, 1) = ListView1.SelectedItem.SubItems(3)
   If ct >= 5 Then frmSales.vs.TextMatrix(k, 5) = ListView1.SelectedItem.SubItems(4)
   If ct >= 5 Then frmSales.vs.TextMatrix(k, 6) = ListView1.SelectedItem.SubItems(5)
   k = k + 1
   
End Sub
