VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frm_concession 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Concession Detail"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdreset 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   6
      Top             =   1245
      Width           =   3540
   End
   Begin VB.ComboBox cbohead 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2220
      TabIndex        =   3
      Top             =   2130
      Width           =   2010
   End
   Begin VB.TextBox txtvalue 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2220
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2535
      Width           =   2010
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   5
      Top             =   615
      Width           =   3540
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Main"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4320
      TabIndex        =   8
      Top             =   1845
      Width           =   3540
   End
   Begin VB.ComboBox con_cat 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2220
      TabIndex        =   2
      Top             =   1740
      Width           =   2010
   End
   Begin VB.TextBox txtadm 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2220
      MaxLength       =   10
      TabIndex        =   0
      Top             =   180
      Width           =   2010
   End
   Begin VB.ComboBox com_ino 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1350
      Width           =   2010
   End
   Begin VSFlex7LCtl.VSFlexGrid grid1 
      Height          =   3735
      Left            =   165
      TabIndex        =   7
      Top             =   3150
      Width           =   7770
      _cx             =   13705
      _cy             =   6588
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   0
      BackColorFixed  =   16761024
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   0
      BackColorBkg    =   12648447
      BackColorAlternate=   -2147483643
      GridColor       =   0
      GridColorFixed  =   16761024
      TreeColor       =   -2147483632
      FloodColor      =   0
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   4
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   20
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
      ExplorerBar     =   0
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Concession Head"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   195
      TabIndex        =   19
      Top             =   2160
      Width           =   2100
   End
   Begin VB.Shape Shape1 
      Height          =   3000
      Left            =   120
      Top             =   75
      Width           =   7830
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "For Percentage Based Concession Append % After Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   480
      Left            =   4290
      TabIndex        =   18
      Top             =   2460
      Width           =   3450
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Value   (Press Enter)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1860
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Concession Category:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   225
      TabIndex        =   16
      Top             =   1770
      Width           =   2100
   End
   Begin VB.Label student_name 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4260
      TabIndex        =   15
      Top             =   180
      Width           =   3615
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Admno:(Search By F5)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   225
      TabIndex        =   14
      Top             =   210
      Width           =   1950
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Installment No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   13
      Top             =   1395
      Width           =   1800
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   225
      TabIndex        =   12
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2220
      TabIndex        =   10
      Top             =   585
      Width           =   2010
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2220
      TabIndex        =   9
      Top             =   975
      Width           =   2010
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   225
      TabIndex        =   11
      Top             =   585
      Width           =   915
   End
End
Attribute VB_Name = "frm_concession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim rs As ADODB.Recordset
 Dim rs1 As ADODB.Recordset
 Dim cmd As New ADODB.Command
 Public Flag As Boolean
'
'Private Sub cbo_admno_Click()
'  'On Error Resume Next
'     ChkLateFees.Enabled = True
'      Dim RsIDetails As New ADODB.Recordset
'      Dim rs5 As New ADODB.Recordset
'           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           RsIDetails.Open "select InstallmentNo as INo, Head , Due, Remain, DueDate from studentwisefeesdues where admno = " & Val(Text1.Text) & " and installmentno <= " & Combo1.Text & " and remain <> " & 0 & "", Con
'           If Combo1.Text = "" Then
'           MsgBox "Please select Ino :", vbInformation
'         Exit Sub
'       txtAdmno.SetFocus
'     End If
'
'      Dim rr As New ADODB.Recordset
'      rr.Open "select admno from studentwisefeesdues where admno = " & Val(Text1.Text) & " and installmentno <= " & Combo1.Text & " and remain <> " & 0 & "", Con
'      grid1.Cols = RsIDetails.Fields.Count + 1
'      grid1.AllowUserResizing = flexResizeBoth
'      With grid1
'      .Rows = rr.RecordCount + 1
'      Do Until RsIDetails.EOF
'         s = "<INo|^*  Head  *|^ Dues |^ Remains |^ Recieved |^ Dues date "
'          grid1.FormatString = s
'             .TextMatrix(RsIDetails.AbsolutePosition, 0) = RsIDetails(0)
'                 .TextMatrix(RsIDetails.AbsolutePosition, 1) = RsIDetails(1)
'                    .TextMatrix(RsIDetails.AbsolutePosition, 2) = RsIDetails(2)
'                        .TextMatrix(RsIDetails.AbsolutePosition, 3) = RsIDetails(3)
'                          .TextMatrix(RsIDetails.AbsolutePosition, 5) = RsIDetails(4)
'                          RsIDetails.MoveNext
'                     Loop
'                 grid1.AutoSize 0, , , 50
' End With
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   rs5.Open "select feesetup from student where admno = " & CInt(Text1.Text) & "", Con, adOpenDynamic, adLockBatchOptimistic
'
'   Dim j As Integer
'    j = 1
'      cbo_ino.Clear
'              If rs5.Fields(0) = "Quarterly" Then
'                  cbo_ino.Clear
'                     For a = 1 To 4
'                        cbo_ino.AddItem j
'                            j = a + 1
'                              Next a
'                                 Exit Sub
'                                 ElseIf rs5.Fields(0) = "Monthly" Then
'
'                           cbo_ino.Clear
'                        For a = 0 To 11
'                      j = a + 1
'                  cbo_ino.AddItem j
'                Next a
'                Exit Sub
'           ElseIf rs5.Fields(0) = "Bi-Monthly" Then
'        cbo_ino.Clear
'    For a = 0 To 5
'        j = a + 1
'          cbo_ino.AddItem j
'             Next a
'               Exit Sub
'               ElseIf rs5.Fields(0) = "Quarterly" Then
'                  cbo_ino.Clear
'                     For a = 0 To 3
'                         j = a + 1
'                            cbo_ino.AddItem j
'                         Next a
'                        Exit Sub
'                      ElseIf rs5.Fields(0) = "Half Yearly" Then
'                   cbo_ino.Clear
'                For a = 0 To 1
'              j = a + 1
'          cbo_ino.AddItem j
'       Next a
'      Exit Sub
'    ElseIf rs5.Fields(0) = "Annually" Then
'        cbo_ino.Clear
'            cbo_ino.AddItem "1"
'                Exit Sub
'                End If
'
'rs5.Close
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'End Sub
'
'Private Sub cbo_admno_GotFocus()
'  SendKeys "%{down}"
'End Sub
'
'Private Sub cbo_admno_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then cbo_ino.SetFocus
'End Sub
'
'Private Sub cbo_admno_LostFocus()
'    '''''''''''''''''''''
'End Sub
'
'Private Sub cbo_cat_Click()
'
'   On Error Resume Next
'  ' code for show data in text20 , text14
' '''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim RS As New ADODB.Recordset
'
'If cboclass.Text = "All" And cbosec.Text <> "All" Then
'  MsgBox "Please class and sec........>", vbInformation, "Neotech"
'Exit Sub
'ElseIf cboclass.Text <> "All" And cbosec.Text = "All" Then
'  MsgBox "Please class and sec........>", vbInformation, "Neotech"
'Exit Sub
'ElseIf cboclass.Text = "All" And cbosec.Text = "All" Then
'MsgBox "Please class and sec........>", vbInformation, "Neotech"
'Exit Sub
'End If
'
' Text3.Text = ""
' RS.Open "select amount from headtransfer where class ='" & cboclass.Text & "' and sec = '" & cbosec.Text & "' and head ='" & cbo_head.Text & "'", Con
' Text3.Text = RS(0)
'
'
'End Sub
'
'Private Sub cbo_cat_GotFocus()
'   SendKeys "%{down}"
'End Sub
'
'Private Sub cbo_cat_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then Text3.SetFocus
'End Sub
'
'Private Sub cbo_cat_LostFocus()
'On Error Resume Next
'
'Dim RS As New ADODB.Recordset
'    RS.Open "select * from concession where admno=" & cbo_admno.Text & " and conhead ='" & cbo_head.Text & "' and ino =" & cbo_ino.Text & "", Con
'
'
'    cbo_admno.Text = RS.Fields(0)
'
'    cbo_cat.Text = RS.Fields(1)
'    cbo_head.Text = RS.Fields(2)
'    Text14.Text = RS.Fields(3)
'    Text20.Text = RS.Fields(4)
'    cbo_ino.Text = RS.Fields(5)
'    RS.Close
'''''
'End Sub
'
'Private Sub cbo_head_GotFocus()
'   SendKeys "%{down}"
'End Sub
'
'Private Sub cbo_head_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then cbo_cat.SetFocus
'End Sub

Private Sub cmdreset_Click()
If Not IsNumeric(txtadm) Then
 MsgBox "Invalid Admission Number.", vbInformation
 txtadm.SetFocus
 Exit Sub
End If
If rsdk.State = 1 Then rsdk.Close
rsdk.Open "select * from studentwisefeesdues where admno=" & txtadm.Text, con, adOpenDynamic, adLockOptimistic
If rsdk.RecordCount = 0 Then
 MsgBox "Dues record not available for this student." & vbNewLine & "Unable to Reset.", vbInformation
 Exit Sub
End If
busy
Do Until rsdk.EOF
If rsdk!recived = 0 Then
 rsdk!con_head = Null
 rsdk!remain = rsdk!due
 rsdk!con_amt = 0
 rsdk.Update
End If
rsdk.MoveNext
Loop
free
MsgBox "Process Complete.", vbInformation
End Sub

Private Sub com_ino_Click()
          Dim rs_cat As New ADODB.Recordset
          rs_cat.Open "select con_cat from con_cat", con, adOpenDynamic, adLockBatchOptimistic
          Dim RsIDetails As New ADODB.Recordset
          If RsIDetails.State = 1 Then RsIDetails.Close
          grid1.Refresh
          If com_ino.Text <> "All" Then
           RsIDetails.Open "select InstallmentNo as INo, Head , Due, Remain, DueDate,con_amt from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and installmentno = " & com_ino.Text & " and remain <> " & 0 & " order by installmentno", con
          Else
           RsIDetails.Open "select InstallmentNo as INo, Head , Due, Remain, DueDate,con_amt from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and remain <> " & 0 & " order by installmentno", con
          End If
          If RsIDetails.EOF = True Then
          'MsgBox "Dues are not present or no dues remains!", vbInformation, "Neotech S&S Ltd"
          'Exit Sub
          End If
          If com_ino.Text = "" Then
           MsgBox "Please select Ino :", vbInformation
         Exit Sub
     End If
      Dim r As New ADODB.Recordset
      Dim ss As String
      Dim ii As Integer
     r.Open "select con_cat from con_cat", con, adOpenDynamic, adLockBatchOptimistic
      While Not r.EOF
        If r.BOF = True Then
            ss = r(0)
        Else
            ss = ss & "|" & r(0)
        End If
      r.MoveNext
      Wend
      r.Close
      Dim rr As New ADODB.Recordset
      If com_ino.Text <> "All" Then
       rr.Open "select admno from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and installmentno = " & com_ino.Text & " and remain <> " & 0 & "", con
      Else
       rr.Open "select admno from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and remain <> " & 0 & "", con
      End If
      grid1.Cols = RsIDetails.fields.Count + 1
      grid1.AutoSizeMode = flexAutoSizeColWidth
      With grid1
      .Rows = rr.RecordCount + 1
      Dim s As String
      On Error Resume Next
      Do Until RsIDetails.EOF
         If rs_cat.RecordCount > 0 Then rs_cat.MoveFirst
         s = "<INo|<Head Name|> Dues |^Concession|^  Percent  |^ Remain Amount"
         grid1.FormatString = s
             .TextMatrix(RsIDetails.AbsolutePosition, 0) = RsIDetails(0)
             .TextMatrix(RsIDetails.AbsolutePosition, 1) = RsIDetails(1)
             .TextMatrix(RsIDetails.AbsolutePosition, 2) = RsIDetails(2)
             .TextMatrix(RsIDetails.AbsolutePosition, 3) = RsIDetails(5)
             .TextMatrix(RsIDetails.AbsolutePosition, 4) = .TextMatrix(RsIDetails.AbsolutePosition, 3) / .TextMatrix(RsIDetails.AbsolutePosition, 2) * 100
             .TextMatrix(RsIDetails.AbsolutePosition, 5) = .TextMatrix(RsIDetails.AbsolutePosition, 2) - .TextMatrix(RsIDetails.AbsolutePosition, 3)
       RsIDetails.MoveNext
     Loop
                 grid1.AutoSize 0, 5, , 50
 End With
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If RsIDetails.State = 1 Then RsIDetails.Close
If rsdk.State = 1 Then rsdk.Close
 If com_ino = "All" Then
  rsdk.Open "select distinct Head from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and remain <> " & 0 & " order by head", con
 Else
  rsdk.Open "select distinct Head from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and installmentno = " & com_ino.Text & " and remain <> " & 0 & " order by head", con
 End If
 cbohead.Clear
 Do Until rsdk.EOF
  cbohead.AddItem rsdk(0)
  rsdk.MoveNext
 Loop
 cbohead.AddItem "All"
End Sub


Private Sub com_ino_GotFocus()
 On Error Resume Next
    Dim rs As New ADODB.Recordset
    rs.Open "select class , sec,student_name from student where admno=" & txtadm.Text & "", con, adOpenDynamic, adLockBatchOptimistic
    Label4.Caption = rs(0)
    Label5.Caption = rs(1)
    student_name.Caption = "      " & rs(2)
  If rs.EOF = True Then
  Label4.Caption = ""
  Label5.Caption = ""
  If MsgBox("Admission no. not found!", vbInformation, "Neotech S&S Ltd") = vbOK Then
  End If
  Exit Sub
  End If

Dim x As Integer

Dim rs5 As New ADODB.Recordset

 rs5.Open "select feesetup from student where admno = " & CInt(txtadm.Text) & "", con, adOpenDynamic, adLockBatchOptimistic
   Dim j, a As Integer
    j = 1
      com_ino.Clear
              
              
              If rs5.fields(0) = "Monthly" Then
                 x = 12
              ElseIf rs5.fields(0) = "Bi-Monthly" Then
               x = 6
              ElseIf rs5.fields(0) = "Quarterly" Then
               x = 4
              ElseIf rs5.fields(0) = "Half Yearly" Then
               x = 2
              ElseIf rs5.fields(0) = "Annually" Then
               x = 1
              End If
If rs5.State = 1 Then rs5.Close
For j = 1 To x
 com_ino.AddItem j
Next j
com_ino.AddItem "All"


End Sub
'
'Private Sub cboclass_Click()
'  On Error Resume Next
'  Dim ss As New ADODB.Recordset
'  ss.Open "select sec from classsec where class ='" & cboclass.Text & "'", Con, adOpenDynamic, adLockBatchOptimistic
'   cbosec.Clear
'   While ss.EOF <> True
'    cbosec.AddItem ss(0)
'    cbosec.ListIndex = 0
'    ss.MoveNext
'   Wend
'   cbosec.AddItem "All"
'
'End Sub
'
'Private Sub cboclass_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'  cbosec.SetFocus
'  End If
'
'End Sub
'
'Private Sub cbosec_Click()
'On Error Resume Next
'Set RS = New ADODB.Recordset
'
'  student_list.Clear
'  stu_list.Clear
'  cbo_admno.Clear
'  RS.Open "select admno,student_name from student where class ='" & cboclass.Text & "' and sec ='" & cbosec.Text & "'", Con
'  While RS.EOF <> True
'  student_list.AddItem RS(1)
'  stu_list.AddItem RS(0)
'  cbo_admno.AddItem RS(0)
'  RS.MoveNext
'  Wend
'  student_list.AddItem "All"
'
' If cboclass = "All" And cbosec = "All" Then
'
'  student_list.Clear
'  stu_list.Clear
'  cbo_admno.Clear
'  Set RS = New ADODB.Recordset
'
'  RS.Open "select admno,student_name from student", Con
'
'  While RS.EOF <> True
'  student_list.AddItem RS(1)
'  stu_list.AddItem RS(0)
'  cbo_admno.AddItem RS(0)
'  RS.MoveNext
'  Wend
'
'
' End If
'
'
' If cboclass <> "All" And cbosec = "All" Then
'
'  student_list.Clear
'  stu_list.Clear
'  cbo_admno.Clear
'  Set RS = New ADODB.Recordset
'
'  RS.Open "select admno,student_name from student where class='" & cboclass.Text & "'", Con
'
'
'
'  While RS.EOF <> True
'  student_list.AddItem RS(1)
'  stu_list.AddItem RS(0)
'  cbo_admno.AddItem RS(0)
'  RS.MoveNext
'  Wend
'
'
' End If
'
'
'
'  RS.Close
'End Sub

Private Sub cbosec_GotFocus()
  SendKeys "%{down}"
End Sub
'
'Private Sub cbosec_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then student_list.SetFocus
'
'End Sub

Private Sub Command1_Click()
   If txtadm.Text = "" Then
    MsgBox "Please Select Admission No.", vbInformation
    txtadm.SetFocus
    Exit Sub
   End If
   If con_cat.Text = "" Then
      MsgBox "Please Select Concession Category ", vbInformation, "Neotech S&S Ltd"
      con_cat.SetFocus
      Exit Sub
   End If
   If cbohead.Text = "" Then
    MsgBox "Please select Fees Head.", vbInformation
    cbohead.SetFocus
    Exit Sub
   End If
   If grid1.Rows = 1 Then
    MsgBox "No concession record to save.", vbInformation
    Exit Sub
   End If
    Dim r1 As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set rs1 = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    
   'On Error Resume Next
      
   If txtadm.Text = "" And com_ino.Text = "" Then
   MsgBox "Enter Admission no. and Installment No.!", vbInformation, "Neotech S&S Ltd"
   txtadm.SetFocus
   Exit Sub
   End If
   If Flag = False Then
    MsgBox "Invalid concession declaration.", vbInformation
    txtvalue.SetFocus
    Exit Sub
   End If
   Dim k As Boolean
    For i = 1 To grid1.Rows - 1
       If grid1.TextMatrix(i, 3) <> "" Then
        k = True
       End If
   Next
    
  If k = False Then
     MsgBox "Please enter concession Amount/Percent", vbInformation, "Neotech S&S Ltd"
     Exit Sub
  End If
    
    
    For i = 1 To grid1.Rows - 1
    
    If rs1.State = 1 Or rs.State = 1 Then
     rs1.Close
     rs.Close
    End If
   busy
     
    
    rs1.Open "select * from concession where admno=" & txtadm.Text & " and conhead ='" & grid1.TextMatrix(i, 2) & "' and ino =" & grid1.TextMatrix(i, 0) & "", con, adOpenDynamic, adLockOptimistic
    rs.Open "select * from concession1 where admno=" & txtadm.Text & " and conhead ='" & grid1.TextMatrix(i, 2) & "' and ino =" & grid1.TextMatrix(i, 0) & "", con, adOpenDynamic, adLockOptimistic
    
    If grid1.TextMatrix(i, 4) <> "" And grid1.TextMatrix(i, 5) <> "" Then
    
    
    
    rs1.AddNew
    rs1.fields(0) = Val(txtadm.Text)
    rs1.fields(1) = grid1.TextMatrix(i, 1)
    rs1.fields(2) = grid1.TextMatrix(i, 4)
    rs1.fields(3) = Val(grid1.TextMatrix(i, 5))
    rs1.fields(4) = Val(grid1.TextMatrix(i, 0))
    rs1.fields(5) = CDate(Date)
    rs1.fields(6) = con_cat.Text
    rs1.Update
    
    rs.AddNew
    rs.fields(0) = Val(txtadm.Text)
    rs.fields(1) = grid1.TextMatrix(i, 1)
    rs.fields(2) = grid1.TextMatrix(i, 4)
    rs.fields(3) = Val(grid1.TextMatrix(i, 3))
    rs.fields(4) = Val(grid1.TextMatrix(i, 0))
    rs.fields(5) = CDate(Date)
    rs.fields(6) = con_cat.Text
    rs.Update
     
 
    Set cmd.ActiveConnection = con
    cmd.CommandText = "update studentwisefeesdues set Remain=" & Val(grid1.TextMatrix(i, 5)) & ",con_head='" & con_cat.Text & "',con_amt=" & Val(grid1.TextMatrix(i, 3)) & "  where admno =" & txtadm.Text & " and installmentno=" & grid1.TextMatrix(i, 0) & " and Head ='" & grid1.TextMatrix(i, 1) & "'"
    cmd.Execute
    
    End If
    Next i
   On Error GoTo 0
    
    MsgBox "Concession Record Saved successfully.", vbInformation, "NeoTech S&S Ltd."
    Command1.Enabled = False

'==============================================================
Dim C As New ADODB.Command
C.ActiveConnection = con
C.CommandText = "update student set isconcession = true where admno =" & txtadm.Text & " "
C.Execute
free
End Sub
'
'Private Sub Command2_Click()
'    On Error Resume Next
'    Set RS = New ADODB.Recordset
'    Set rs1 = New ADODB.Recordset
'
'    rs1.Open "delete from concession where admno=" & cbo_admno.Text & " and conhead ='" & cbo_head.Text & "' and ino =" & cbo_ino.Text & "", Con, adOpenDynamic, adLockOptimistic
'    RS.Open "delete from concession1 where admno=" & cbo_admno.Text & " and conhead ='" & cbo_admno.Text & "' and ino =" & cbo_ino.Text & "", Con, adOpenDynamic, adLockOptimistic
'
'
'       cbo_admno.Text = ""
'       cbo_cat.Text = ""
'       cbo_head.Text = ""
'       Text14.Text = ""
'       Text3.Text = ""
'       Text20.Text = ""
'       cbo_ino.Text = ""
'
'MsgBox "Delete successfully....", vbInformation, "NeoTech S&S Ltd."
'
'
'End Sub
'
'Private Sub Command3_Click()
'       Dim RS As New ADODB.Recordset
'       Set rs1 = New ADODB.Recordset
'
'    On Error Resume Next
'
'    rs1.Open "select * from concession where admno=" & cbo_admno.Text & " and conhead ='" & cbo_head.Text & "' and ino =" & cbo_ino.Text & "", Con, adOpenDynamic, adLockOptimistic
'    RS.Open "select * from concession1 where admno=" & cbo_admno.Text & " and conhead ='" & cbo_head.Text & "' and ino =" & cbo_ino.Text & "", Con, adOpenDynamic, adLockOptimistic
'
'
'    rs1.Fields(0) = cbo_admno.Text
'    rs1.Fields(1) = cbo_cat.Text
'    rs1.Fields(2) = cbo_head.Text
'    rs1.Fields(3) = Val(Trim(Text14.Text))
'    rs1.Fields(4) = Val(Trim(Text20.Text))
'    rs1.Fields(5) = Val(Trim(cbo_ino.Text))
'    rs1.Update
'
'
'    RS.Fields(0) = cbo_admno.Text
'    RS.Fields(1) = cbo_cat.Text
'    RS.Fields(2) = cbo_head.Text
'    RS.Fields(3) = 100 - Val(Trim(Text14.Text))
'    RS.Fields(4) = Val(Text3.Text) - Val(Text20.Text)
'    RS.Fields(5) = Val(Trim(cbo_ino.Text))
'    RS.Update
'
'     Set cmd.ActiveConnection = Con
'    cmd.CommandText = "update studentwisefeesdues set Due = " & Val(Text3.Text) - Val(Text20.Text) & " , Remain=" & Val(Text3.Text) - Val(Text20.Text) & "  where admno =" & cbo_admno.Text & " and installmentno=" & cbo_ino.Text & " and Head ='" & cbo_head.Text & "'"
'    cmd.Execute
'
'    MsgBox "Record for concession update successfully....", vbInformation, "NeoTech S&S Ltd."
'
'
'
'    cbo_admno_Click
'    grid1.Refresh
'End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_GotFocus()

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command4_LostFocus()
txtadm.SetFocus
End Sub

Private Sub Form_Activate()
Dim rs_con As New ADODB.Recordset
If rs_con.State = 1 Then rs_con.Close
rs_con.Open "select * from con_cat", con, adOpenDynamic, adLockOptimistic

If rs_con.EOF = False Then
While rs_con.EOF = False
 con_cat.AddItem rs_con(0)
 rs_con.MoveNext
Wend
End If

 txtadm.SetFocus
 'date_lab.Caption = Date
End Sub

Private Sub Form_GotFocus()
SendKeys "%{down}"
End Sub
'
'Private Sub stu_list_KeyPress(KeyAscii As Integer)
'       If KeyAscii = 13 Then cbo_admno.SetFocus
'End Sub
'
'Private Sub student_list_Click()
'On Error Resume Next
'  Dim rs1 As New ADODB.Recordset
'
' rs1.Open "select admno from student where class='" & cboclass.Text & "' and sec='" & cbosec.Text & "' and student_name='" & student_list.Text & "'", Con, adOpenDynamic, adLockBatchOptimistic
' Text1.Text = rs1(0)
' rs1.Close
'
'
'rs1.Open "select admno from student where student_name='" & student_list.Text & "'", Con, adOpenDynamic, adLockOptimistic
'stu_list.Text = rs1(0)
'rs1.Close
'End Sub
'
'Private Sub student_list_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then cbo_admno.SetFocus
'End Sub
'
'
'Private Sub Text14_Change()
' On Error Resume Next
'  If Text14.Text <= 100 Then
'  Text20.Text = Val(Text3.Text) * Val(Text14.Text) / 100
'  Exit Sub
' Else
' MsgBox "Enter value less then 100....>", vbInformation, "Neotech"
' End If
'End Sub
'
'
'Private Sub Text20_Change()
' On Error Resume Next
'    Text14.Text = Val(Text20.Text) * 100 / Val(Text3.Text)
'End Sub

'
'Private Sub Text14_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then Text20.SetFocus
'End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then Command1.SetFocus
End Sub
'
'Private Sub Text3_KeyPress(KeyAscii As Integer)
'      If KeyAscii = 13 Then Text14.SetFocus
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
       frmSplash.cdbHelp.ShowHelp
    End If
If KeyCode = vbKeyF5 Then
 Set frmfind.o = txtadm
 frmfind.Show 1
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
     Dim rr As New ADODB.Recordset
     rr.Open "select distinct con_cat from concession", con
     While rr.EOF <> True
      con_cat.AddItem rr(0)
      rr.MoveNext
     Wend
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim result, value As Long
    result = 0
    value = 1
    
    If Col = 3 Then
    
                         
              If CInt(grid1.TextMatrix(Row, 2) - grid1.TextMatrix(Row, 3)) < 0 Then
                 If MsgBox("Please insert valid value!", vbInformation, "Neotech S&S Ltd") = vbOK Then
                   Call com_ino_Click
                  Exit Sub
               End If
              End If
              
              grid1.TextMatrix(Row, 5) = grid1.TextMatrix(Row, 2) - grid1.TextMatrix(Row, 3)
              grid1.TextMatrix(Row, 4) = FormatNumber((grid1.TextMatrix(Row, 3) * 100 / grid1.TextMatrix(Row, 2)), 2)
              grid1.TextMatrix(Row, 2) = grid1.TextMatrix(Row, 2) - grid1.TextMatrix(Row, 3)
      
  End If

   If Col = 4 Then
              If grid1.TextMatrix(Row, 4) < 0 Or grid1.TextMatrix(Row, 4) > 100 Then
              If MsgBox("Please insert valid value!", vbInformation, "Neotech S&S Ltd") = vbOK Then
                   Call com_ino_Click
                  Exit Sub
              End If
              End If
              
              grid1.TextMatrix(Row, 3) = (Val(grid1.TextMatrix(Row, 2)) * Val(grid1.TextMatrix(Row, 4))) / 100
              grid1.TextMatrix(Row, 5) = (Val(grid1.TextMatrix(Row, 2)) - Val(grid1.TextMatrix(Row, 3)))
              grid1.TextMatrix(Row, 2) = grid1.TextMatrix(Row, 2) - grid1.TextMatrix(Row, 3)
      
End If



End Sub

Private Sub grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, j As Integer)
        
          
If Col = 3 Then
     If (j >= 65 And j <= 90) Or (j >= 97 And j <= 122) Or j = 13 Then
     j = 0
     End If
  End If
  
        
       
    
    
   End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtadm_KeyPress(KeyAscii As Integer)

 If Not IsNumeric(txtadm) Then
  Exit Sub
 End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 13) Then
    KeyAscii = 0
End If
     Command1.Enabled = True
     Dim rs_cat As New ADODB.Recordset
        
      If KeyAscii = 13 Then
       
          Set rs = New ADODB.Recordset
          If rs.State = 1 Then rs.Close
          If txtadm.Text = "" Then Exit Sub
          rs.Open "select student_name,class,sec from student where admno=" & txtadm.Text & "", con, adOpenDynamic, adLockOptimistic
          If rs.EOF = False Then
            student_name.Caption = " " & rs(0)
            Label4.Caption = " " & rs(1)
            Label5.Caption = " " & rs(2)
          End If
           
           rs_cat.Open "select con_cat from con_cat", con, adOpenDynamic, adLockBatchOptimistic
           Dim RsIDetails As New ADODB.Recordset
           If RsIDetails.State = 1 Then RsIDetails.Close
            grid1.Refresh
            grid1.Clear
              
               RsIDetails.Open "select InstallmentNo as INo, Head , Due, Remain, DueDate,con_amt from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and  remain <> " & 0 & " order by installmentno ", con
               If RsIDetails.EOF = True Then
               MsgBox "No Remaining Dues Available!", vbInformation, "Neotech S&S Ltd"
               Exit Sub
            End If
    
      
      Dim r As New ADODB.Recordset
        Dim ss As String
         Dim ii As Integer
           r.Open "select con_cat from con_cat", con, adOpenDynamic, adLockBatchOptimistic
             While Not r.EOF
               If r.BOF = True Then
             ss = r(0)
           Else
         ss = ss & "|" & r(0)
      End If
    r.MoveNext
 Wend
 r.Close
   Dim rr As New ADODB.Recordset
     rr.Open "select admno from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and remain <> " & 0 & "", con
      grid1.Cols = RsIDetails.fields.Count
      grid1.AutoSizeMode = flexAutoSizeColWidth
         With grid1
       .Rows = rr.RecordCount + 1
    Dim s As String
      Do Until RsIDetails.EOF
   
   If rs_cat.RecordCount > 0 Then
   rs_cat.MoveFirst
   End If
   s = "<INo|<Head Name|>Dues |^Concession|^  Percent  |^ Remain Amount"
        grid1.FormatString = s
          .TextMatrix(RsIDetails.AbsolutePosition, 0) = RsIDetails(0)
             
                .TextMatrix(RsIDetails.AbsolutePosition, 1) = RsIDetails(1)
                  .TextMatrix(RsIDetails.AbsolutePosition, 2) = RsIDetails(2)
                     .TextMatrix(RsIDetails.AbsolutePosition, 3) = RsIDetails(5)
                     .TextMatrix(RsIDetails.AbsolutePosition, 4) = .TextMatrix(RsIDetails.AbsolutePosition, 3) / .TextMatrix(RsIDetails.AbsolutePosition, 2) * 100
                       .TextMatrix(RsIDetails.AbsolutePosition, 5) = RsIDetails(3)
                     RsIDetails.MoveNext
                       Loop
                 grid1.AutoSize 0, 5, , 50
 End With
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            RsIDetails.Close
            
        com_ino.SetFocus
       End If
       If rsdk.State = 1 Then rsdk.Close
  rsdk.Open "select distinct Head from studentwisefeesdues where admno = " & Val(txtadm.Text) & " and remain <> " & 0 & " order by head", con
 cbohead.Clear
 cbohead.AddItem "All"
 Do Until rsdk.EOF
  cbohead.AddItem rsdk(0)
  rsdk.MoveNext
 Loop
 End Sub

Private Sub txtvalue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Dim value As Double
 Dim mystr As String
 Dim location As Integer
 location = InStr(1, txtvalue, "%")
 If location > 0 Then
  mystr = Mid(txtvalue, 1, location - 1)
 Else
  mystr = txtvalue
 End If
 If Val(mystr) = 0 Then
  MsgBox "Invalid Concession Value.", vbInformation
  txtvalue.SetFocus
  SendKeys "{Home}+{End}"
  Flag = False
  Exit Sub
 End If
 value = Val(mystr)
 Dim i As Integer
 If location > 0 And value > 100 Then
  MsgBox "Invalid Value.", vbInformation
  txtvalue.SetFocus
  Flag = False
  Exit Sub
 End If
 Flag = True
 For i = 1 To grid1.Rows - 1
  grid1.Row = i
  If cbohead <> "All" Then
   If grid1.TextMatrix(i, 1) = cbohead.Text Then
    If location = 0 Then
     grid1.TextMatrix(i, 3) = Format(value, "0.00")
     grid1.TextMatrix(i, 4) = Format((value / Val(grid1.TextMatrix(i, 2))) * 100, "0.00")
     grid1.TextMatrix(i, 5) = Format(Val(grid1.TextMatrix(i, 2)) - Val(grid1.TextMatrix(i, 3)), "0.00")
    Else
     grid1.TextMatrix(i, 3) = Format((Val(grid1.TextMatrix(i, 2)) / 100) * value, "0.00")
     grid1.TextMatrix(i, 4) = Format(value, "0.00")
     grid1.TextMatrix(i, 5) = Format(Val(grid1.TextMatrix(i, 2)) - Val(grid1.TextMatrix(i, 3)), "0.00")
    End If
    Else
     ' grid1.TextMatrix(i, 3) = Format(0, "0.00")
     ' grid1.TextMatrix(i, 4) = Format(0, "0.00")
     ' grid1.TextMatrix(i, 5) = Format(Val(grid1.TextMatrix(i, 2)), "0.00")
    End If
  Else
     If location = 0 Then
     grid1.TextMatrix(i, 3) = Format(value, "0.00")
     grid1.TextMatrix(i, 4) = Format((value / Val(grid1.TextMatrix(i, 2))) * 100, "0.00")
     grid1.TextMatrix(i, 5) = Format(Val(grid1.TextMatrix(i, 2)) - Val(grid1.TextMatrix(i, 3)), "0.00")
    Else
     grid1.TextMatrix(i, 3) = Format((Val(grid1.TextMatrix(i, 2)) / 100) * value, "0.00")
     grid1.TextMatrix(i, 4) = Format(value, "0.00")
     grid1.TextMatrix(i, 5) = Format(Val(grid1.TextMatrix(i, 2)) - Val(grid1.TextMatrix(i, 3)), "0.00")
    End If
  End If
  If Val(grid1.TextMatrix(i, 3)) > 0 Then
   For j = 0 To grid1.Cols - 1
    grid1.Col = j
    grid1.CellBackColor = vbYellow
   Next j
  End If
  If Val(grid1.TextMatrix(i, 3)) > Val(grid1.TextMatrix(i, 2)) Then
   MsgBox "Invalid Amount.", vbInformation
   Flag = False
   GoTo makezero
  End If
 Next i
End If
Exit Sub
makezero:
 For i = 1 To grid1.Rows - 1
   grid1.Row = i
   grid1.TextMatrix(i, 3) = Format(0, "0.00")
   grid1.TextMatrix(i, 4) = Format(0, "0.00")
   grid1.TextMatrix(i, 5) = Format(Val(grid1.TextMatrix(i, 2)), "0.00")
   For j = 0 To grid1.Cols - 1
    grid1.Col = j
    grid1.CellBackColor = vbWhite
   Next j
 Next i
End Sub
