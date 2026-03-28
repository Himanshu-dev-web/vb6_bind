VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form setup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7365
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   6195
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   9015
      Begin VB.TextBox txtExcise 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6060
         MaxLength       =   244
         TabIndex        =   59
         Top             =   3240
         Width           =   945
      End
      Begin VB.TextBox txtEducationCess 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1860
         MaxLength       =   244
         TabIndex        =   57
         Top             =   3240
         Width           =   945
      End
      Begin VB.TextBox txtvatper 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6060
         MaxLength       =   244
         TabIndex        =   55
         Top             =   2940
         Width           =   945
      End
      Begin VB.TextBox txtcstper 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6060
         MaxLength       =   244
         TabIndex        =   53
         Top             =   2640
         Width           =   945
      End
      Begin VB.TextBox txtInvoivcecondition 
         DataField       =   "FOOTNOTE3"
         DataSource      =   "Data1"
         Height          =   1245
         Left            =   1860
         MaxLength       =   244
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   4860
         Width           =   7035
      End
      Begin VB.TextBox txttin 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1830
         MaxLength       =   244
         TabIndex        =   14
         Top             =   2940
         Width           =   2985
      End
      Begin VB.TextBox txtcst 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1860
         MaxLength       =   244
         TabIndex        =   13
         Top             =   2640
         Width           =   2985
      End
      Begin VB.TextBox txtuptt 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5250
         MaxLength       =   244
         TabIndex        =   12
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtbankadvice 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   11
         Top             =   2280
         Width           =   1995
      End
      Begin VB.TextBox txtemail 
         DataField       =   "FOOTNOTE1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   8
         Top             =   1530
         Width           =   4845
      End
      Begin VB.TextBox txtfax 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6990
         MaxLength       =   244
         TabIndex        =   7
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtCname 
         DataField       =   "CLINIC_NAME"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   4905
      End
      Begin VB.TextBox add1 
         DataField       =   "ADDRESS1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   2
         Top             =   555
         Width           =   4935
      End
      Begin VB.TextBox add2 
         DataField       =   "ADDRESS2"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   3
         Top             =   885
         Width           =   3795
      End
      Begin VB.TextBox txtphone1 
         DataField       =   "PHONENO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1845
         MaxLength       =   244
         TabIndex        =   5
         Top             =   1215
         Width           =   2085
      End
      Begin VB.TextBox txtcourt 
         DataField       =   "FOOTNOTE1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1830
         MaxLength       =   244
         TabIndex        =   15
         Top             =   3570
         Width           =   7005
      End
      Begin VB.TextBox txtrem1 
         DataField       =   "FOOTNOTE2"
         DataSource      =   "Data1"
         Height          =   645
         Left            =   1830
         MaxLength       =   244
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   16
         Top             =   3885
         Width           =   6975
      End
      Begin VB.TextBox txtrem2 
         DataField       =   "FOOTNOTE3"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1830
         MaxLength       =   244
         TabIndex        =   17
         Top             =   4560
         Width           =   7035
      End
      Begin VB.TextBox txtphone2 
         DataField       =   "FAXNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3960
         MaxLength       =   244
         TabIndex        =   6
         Top             =   1200
         Width           =   1995
      End
      Begin VB.TextBox CITY 
         DataField       =   "CITY"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6420
         MaxLength       =   244
         TabIndex        =   4
         Top             =   870
         Width           =   2415
      End
      Begin MSMask.MaskEdBox txtyarfrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   1845
         TabIndex        =   9
         Top             =   1860
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtyarto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   1860
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Excise % :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   5040
         TabIndex        =   60
         Top             =   3300
         Width           =   885
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Education Cess % :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   58
         Top             =   3300
         Width           =   1650
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "V. A. T. % :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   5040
         TabIndex        =   56
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "C. S. T. % :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   5040
         TabIndex        =   54
         Top             =   2700
         Width           =   990
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Terms && Condition :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   52
         Top             =   4920
         Width           =   1665
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "T. I. N. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   900
         TabIndex        =   51
         Top             =   3000
         Width           =   750
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Company  Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   35
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Foot Note Report :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   34
         Top             =   3600
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Phone No  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   660
         TabIndex        =   33
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   885
         TabIndex        =   32
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Remark 1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   750
         TabIndex        =   31
         Top             =   3930
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Remark 2 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   750
         TabIndex        =   30
         Top             =   4620
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "From Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   705
         TabIndex        =   29
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1110
         TabIndex        =   28
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Bank Advice No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   27
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "City :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5790
         TabIndex        =   26
         Top             =   915
         Width           =   450
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Fax No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   6060
         TabIndex        =   25
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "To"
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
         Left            =   3240
         TabIndex        =   24
         Top             =   1890
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "U.P.T.T. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   4215
         TabIndex        =   23
         Top             =   2340
         Width           =   870
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "C. S. T. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   22
         Top             =   2700
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4425
      Left            =   840
      TabIndex        =   36
      Top             =   450
      Visible         =   0   'False
      Width           =   7965
      Begin VB.CheckBox cfooter 
         Caption         =   "Print Footer"
         Height          =   345
         Left            =   180
         TabIndex        =   45
         Top             =   2670
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox cheader 
         Caption         =   "Print Header"
         Height          =   345
         Left            =   180
         TabIndex        =   44
         Top             =   2220
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox footertext 
         Height          =   885
         Left            =   2220
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   2970
         Width           =   4785
      End
      Begin VB.TextBox headertext 
         Height          =   885
         Left            =   2220
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   1650
         Width           =   4785
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1290
         TabIndex        =   41
         Top             =   1710
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1290
         TabIndex        =   40
         Top             =   1320
         Width           =   555
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2550
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   3345
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   1905
         TabIndex        =   39
         Top             =   1710
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         AutoBuddy       =   -1  'True
         OrigLeft        =   2610
         OrigTop         =   570
         OrigRight       =   2850
         OrigBottom      =   1245
         Max             =   35
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1845
         TabIndex        =   46
         Top             =   1320
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "cfooter"
         BuddyDispid     =   196634
         OrigLeft        =   2610
         OrigTop         =   570
         OrigRight       =   2850
         OrigBottom      =   1245
         Max             =   35
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label flabel 
         Caption         =   "Footer Text"
         Height          =   195
         Left            =   2220
         TabIndex        =   50
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label hlabel 
         Caption         =   "Header Text"
         Height          =   195
         Left            =   2220
         TabIndex        =   49
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Bottom Margin"
         Height          =   195
         Left            =   150
         TabIndex        =   48
         Top             =   1710
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Top Margin"
         Height          =   195
         Left            =   150
         TabIndex        =   47
         Top             =   1290
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Report Name :-"
         Height          =   315
         Left            =   1050
         TabIndex        =   38
         Top             =   270
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin MSComctlLib.TabStrip SStab1 
      Height          =   6705
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   11827
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Setup"
            Key             =   "Tab1"
            Object.Tag             =   "Tab1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4500
      TabIndex        =   21
      Top             =   6840
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2670
      TabIndex        =   19
      Top             =   6840
      Width           =   1575
   End
End
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cfooter_Click()
    If cfooter.Value = 1 Then
        flabel.Enabled = True
        footertext.Enabled = True
    Else
        flabel.Enabled = False
        footertext.Enabled = False
    End If
End Sub

Private Sub cheader_Click()
    If cheader.Value = 0 Then
        hlabel.Enabled = False
        headertext.Enabled = False
    Else
        hlabel.Enabled = True
        headertext.Enabled = True
    End If
End Sub

Private Sub Command1_Click()
If SStab1.Tabs(1).Key = "Tab1" Then
   Set rs = New ADODB.Recordset
   rs.Open "SETUP1", con, adOpenDynamic, adLockOptimistic, adCmdTable
   rs!CNAME = Me.txtCname.Text
   rs!add1 = Me.add1.Text
   rs!add2 = Me.add2.Text
   rs!CITY = CITY.Text
   rs!phone1 = Me.txtphone1.Text
   rs!phone2 = Me.txtphone2.Text
   rs!fax = Me.txtfax.Text
   rs!yarfrom = txtyarfrom.Text
   rs!yarto = txtyarto.Text
   rs!email = txtemail.Text
   rs!COURT = txtcourt.Text
   rs!rem1 = txtrem1.Text
   rs!rem2 = txtrem2.Text
   rs!bankadviceno = txtbankadvice.Text
   rs!uptt = txtuptt.Text
   rs!cst = txtCST.Text
   rs!cstper = txtcstper.Text
   rs!tin = txttin.Text
   rs!vatper = txtvatper.Text
   rs!EducationCessPer = txtEducationCess.Text
   rs!ExcisePer = txtExcise.Text
   rs!Invoivcecondition = txtInvoivcecondition.Text
   rs.Update
End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
pos Me

Me.Top = 200
Me.Left = 300


End Sub

Private Sub SSTab1_GotFocus()
On Error Resume Next
If SStab1.Tabs(1).Key = "Tab1" Then
   Frame3.Visible = False
   Frame2.Visible = True
   Set rs = New ADODB.Recordset
   rs.Open "SETUP1", con, adOpenDynamic, adLockOptimistic, adCmdTable
   Me.txtCname.Text = rs!CNAME
   Me.add1.Text = rs!add1
   Me.add2.Text = rs!add2
   Me.CITY.Text = rs!CITY
   Me.txtphone1.Text = rs!phone1
   Me.txtphone2.Text = rs!phone2
   Me.txtfax = Trim(rs!fax)
   txtyarfrom.Text = Trim(rs!yarfrom)
   txtyarto.Text = Trim(rs!yarto)
   txtemail.Text = rs!email
   txtcourt.Text = rs!COURT
   txtrem1.Text = rs!rem1
   txtrem2.Text = rs!rem2
   txtbankadvice.Text = rs!bankadviceno
   txtuptt.Text = rs!uptt
   txtCST.Text = rs!cst
   txtvatper.Text = rs!vatper
   txtcstper.Text = rs!cstper
   txtEducationCess.Text = rs!EducationCessPer
   txtExcise.Text = rs!ExcisePer
   txttin.Text = rs!tin
   txtInvoivcecondition.Text = rs!Invoivcecondition
   rs.Update
End If
    
End Sub

