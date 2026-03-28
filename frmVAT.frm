VERSION 5.00
Begin VB.Form frmVAT 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtvatOn 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1380
      Width           =   1095
   End
   Begin VB.TextBox txtvat 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "VAT ON"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   435
      Left            =   840
      TabIndex        =   4
      Top             =   900
      Width           =   1275
   End
End
Attribute VB_Name = "frmVAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
con.Execute "update setup1 set vatper=" & Val(txtvat.Text) & ",vatOn=" & Val(txtvatOn) & ""
End Sub
Sub fill()
   
   If rs.State = 1 Then rs.Close
   rs.Open "select vatper,vatOn from setup1", con
   If rs.EOF = False Then
      txtvat = rs(0) & ""
      txtvatOn = rs(1) & ""
   End If
   
End Sub
Private Sub Form_Load()
fill
End Sub
