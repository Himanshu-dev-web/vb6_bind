VERSION 5.00
Begin VB.Form frmTaxMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Master"
   ClientHeight    =   2070
   ClientLeft      =   4260
   ClientTop       =   3225
   ClientWidth     =   4005
   Icon            =   "frmTaxMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4005
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1410
      Width           =   1845
   End
   Begin VB.TextBox txtTaxPer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   1
      Top             =   765
      Width           =   510
   End
   Begin VB.TextBox txtTaxHead 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1185
      TabIndex        =   0
      Top             =   375
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Tax % :"
      Height          =   255
      Left            =   345
      TabIndex        =   3
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "Tax Head :"
      Height          =   270
      Left            =   315
      TabIndex        =   2
      Top             =   390
      Width           =   750
   End
End
Attribute VB_Name = "frmTaxMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub cmdOk_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from TaxHead", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
       rs.Fields(0).Value = txtTaxHead.Text
       rs.Fields(1).Value = txtTaxper.Text
       rs.Update
       MsgBox "Data Saved !! ", vbInformation
    End If
End Sub

Private Sub Form_Load()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from TaxHead", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
       txtTaxHead.Text = rs.Fields(0).Value
       txtTaxper.Text = rs.Fields(1).Value
    End If
End Sub

