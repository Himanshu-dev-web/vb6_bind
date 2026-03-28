VERSION 5.00
Begin VB.Form UnitMaster 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folding Rate Master"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "UnitMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2820
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtid 
      Height          =   285
      Left            =   6195
      TabIndex        =   9
      Top             =   510
      Width           =   375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4890
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   4380
   End
   Begin VB.TextBox txtproduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   690
      Left            =   0
      TabIndex        =   2
      Top             =   5940
      Width           =   5220
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   495
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   495
         Left            =   1245
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1335
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   495
         Left            =   3915
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1215
      End
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Type Of Mode                             Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   180
      Width           =   3795
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E98A0A&
      BackStyle       =   0  'Transparent
      Caption         =   "Esc To Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4110
      TabIndex        =   8
      Top             =   6615
      Width           =   1335
   End
End
Attribute VB_Name = "UnitMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bb As Boolean
Private Sub cmdDel_Click()
On Error Resume Next
If rs.State = 1 Then rs.Close
         rs.Open "select * from UnitMaster where name='" & Trim(txtproduct) & "'", con, adOpenDynamic, adLockOptimistic
         If rs.EOF = False Then
         If MsgBox("Do U Want To Delete ?", vbInformation + vbYesNo, "Message") = vbYes Then
            rs.Delete
            addProduct
            txtproduct.Text = ""
            txtRate = ""
         End If
         End If
End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub
Sub addProduct()
      'On Error Resume Next
      If rs.State = 1 Then rs.Close
         rs.Open "select * from  UnitMaster", con, adOpenDynamic, adLockOptimistic
         List1.Clear
         If rs.EOF = False Then
            While rs.EOF = False
                List1.AddItem rs.Fields(0).Value & Space(60 - Len(rs.Fields(0).Value)) & rs(1)
                rs.MoveNext
            Wend
         End If
End Sub
Private Sub cmdRef_Click()
 txtproduct.Text = ""
 txtid.Text = ""
End Sub

Private Sub cmdSave_Click()
 save
End Sub
Sub save()
   On Error Resume Next
       
   If txtproduct.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If

         
         If rs.State = 1 Then rs.Close
         If txtid.Text <> "" Then
         rs.Open "select * from UnitMaster where Aouto=" & txtid.Text & "", con, adOpenDynamic, adLockOptimistic
         Else
         rs.Open "select * from UnitMaster where name='" & txtproduct & "'", con, adOpenDynamic, adLockOptimistic
         End If
         
         If rs.EOF = True Then
            rs.AddNew
            rs.Fields(0).Value = txtproduct.Text
            rs.Fields(1).Value = Val(txtRate.Text)
            
            rs.Update
            txtproduct.Text = ""
            addProduct
          Else
            rs.Fields(0).Value = txtproduct.Text
            rs.Fields(1).Value = Val(txtRate.Text)
            
            rs.Update
            txtproduct.Text = ""
            addProduct
        
         End If
         txtid.Text = ""
         txtproduct.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    'SendKeys "{tab}"
 End If
 If KeyCode = 27 Then
    Unload Me
 End If
End Sub
Private Sub Form_Load()
addProduct
Me.Left = 2500
pos Me
End Sub
Private Sub List1_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from UnitMaster where Name='" & Trim(Left(List1.Text, 20)) & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
txtid.Text = rs.Fields(2).Value

txtproduct.Text = rs(0)
txtRate.Text = rs(1)


End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
   On Error Resume Next
     
   If txtproduct.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If

   If KeyAscii = 13 Then
      If rs.State = 1 Then rs.Close
         rs.Open "select * from UnitMaster where Name='" & txtproduct & "'", con, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.AddNew
            rs.Fields(0).Value = txtproduct.Text
            rs.Update
            txtproduct.Text = ""
            addProduct
          Else
            rs.Fields(0).Value = txtproduct.Text
            rs.Update
            txtproduct.Text = ""
            addProduct
         End If
   End If

End Sub

Private Sub List2_Click()
txtproduct.Text = List1.Text
End Sub
Private Sub txtproduct_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      save
      txtproduct.SetFocus
   End If
End Sub

