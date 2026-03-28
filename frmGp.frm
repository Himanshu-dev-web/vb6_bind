VERSION 5.00
Begin VB.Form frmGp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Master"
   ClientHeight    =   6444
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5220
   Icon            =   "frmGp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6444
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   5220
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   495
         Left            =   3915
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1215
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
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   495
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1335
      End
   End
   Begin VB.TextBox txtproduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5235
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   5016
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5220
   End
   Begin VB.TextBox txtid 
      Height          =   285
      Left            =   5355
      TabIndex        =   0
      Top             =   90
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E98A0A&
      BackStyle       =   0  'Transparent
      Caption         =   "Esc To Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3990
      TabIndex        =   8
      Top             =   6075
      Width           =   1335
   End
End
Attribute VB_Name = "frmGp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bb As Boolean
Private Sub cmdDel_Click()
On Error Resume Next
If rs.State = 1 Then rs.Close
         rs.Open "select * from ConsumeGp where name='" & List1 & "'", con, adOpenDynamic, adLockOptimistic
         If rs.EOF = False Then
         If MsgBox("Do U Want To Delete ?", vbInformation + vbYesNo, "Message") = vbYes Then
            rs.Delete
            addProduct
            txtproduct.text = ""
         End If
         End If
End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub

Sub addProduct()
      'On Error Resume Next
      Set rs = New ADODB.Recordset
         rs.Open "select * from  ConsumeGp", con
         List1.clear
         If rs.EOF = False Then
            While rs.EOF = False
                List1.AddItem rs.Fields(0).Value
                rs.MoveNext
            Wend
         End If
End Sub
Private Sub cmdRef_Click()
 txtproduct.text = ""
 txtid.text = ""
End Sub

Private Sub cmdSave_Click()
 save
End Sub
Sub save()
   On Error Resume Next
       
   If txtproduct.text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If

            
         
         If rs.State = 1 Then rs.Close
         If txtid.text <> "" Then
         rs.Open "select * from ConsumeGp where Aouto=" & txtid.text & "", con, adOpenDynamic, adLockOptimistic
         Else
         rs.Open "select * from ConsumeGp where name='" & txtproduct & "'", con, adOpenDynamic, adLockOptimistic
         End If
         
         If rs.EOF = True Then
            rs.AddNew
            rs.Fields(0).Value = txtproduct.text
            rs.Update
            'con.Execute "update ItemMaster set ItemGp='" & txtproduct.text & "' where ItemGp='" & List1.text & "'"
            txtproduct.text = ""
            addProduct
          Else
            rs.Fields(0).Value = txtproduct.text
            rs.Update
            con.Execute "update ItemMaster set ItemGp='" & txtproduct.text & "' where ItemGp='" & List1.text & "'"
            txtproduct.text = ""
            addProduct
        
         End If
         
         
          
         
         
         txtid.text = ""
         
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
txtproduct.text = List1.text
Set rs = New ADODB.Recordset
rs.Open "select * from ConsumeGp where Name='" & txtproduct & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
txtid.text = rs.Fields(1).Value
End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
   On Error Resume Next
     
   If txtproduct.text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If

   If KeyAscii = 13 Then
      If rs.State = 1 Then rs.Close
         rs.Open "select * from ConsumeGp where Name='" & txtproduct & "'", con, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.AddNew
            rs.Fields(0).Value = txtproduct.text
            rs.Update
            txtproduct.text = ""
            addProduct
          Else
            rs.Fields(0).Value = txtproduct.text
            rs.Update
            txtproduct.text = ""
            addProduct
         End If
   End If

End Sub

Private Sub List2_Click()
txtproduct.text = List1.text
End Sub
Private Sub txtproduct_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      save
      txtproduct.SetFocus
   End If
End Sub


