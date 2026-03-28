VERSION 5.00
Begin VB.Form frmAddMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   ControlBox      =   0   'False
   Icon            =   "frmAddMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1005
      Width           =   2595
   End
   Begin VB.ComboBox cboGp 
      Height          =   315
      ItemData        =   "frmAddMaster.frx":0442
      Left            =   1200
      List            =   "frmAddMaster.frx":0464
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   615
      Width           =   2625
   End
   Begin VB.Label Label1 
      Caption         =   "Please Select Group Name "
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   3630
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Group "
      Height          =   285
      Left            =   165
      TabIndex        =   2
      Top             =   630
      Width           =   1650
   End
End
Attribute VB_Name = "frmAddMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
   Me.Visible = False
   frmIssue.Vs1Frame.Visible = False
   frmIssue.Vs3Frame.Visible = False
   
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
  fillcombo cboGp, "ItemGp", "ItemMaster", con
End Sub
