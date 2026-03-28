VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmsaleinvoice 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sale invoice"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   Icon            =   "frmsaleinvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Print &All"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   6780
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print &Extra Copy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   7260
      Width           =   1095
   End
   Begin VB.CommandButton Commandprinttriplicate 
      Caption         =   "Print &Triplicate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4290
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   7260
      Width           =   1095
   End
   Begin VB.CommandButton Commandprintduplicate 
      Caption         =   "Print &Duplicate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3195
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   7260
      Width           =   1095
   End
   Begin VB.CommandButton Commandprintoriginal 
      Caption         =   "Print &Original"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   7260
      Width           =   1095
   End
   Begin VB.TextBox txtTaxSaleInvoiceNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   21
      Top             =   5550
      Width           =   1965
   End
   Begin VB.TextBox txtTransportationDate 
      Height          =   315
      Left            =   6000
      TabIndex        =   22
      Top             =   5550
      Width           =   1635
   End
   Begin VB.TextBox tot1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   9360
      TabIndex        =   71
      Top             =   6540
      Width           =   1605
   End
   Begin VB.TextBox txtRoundoff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9360
      TabIndex        =   69
      Top             =   6825
      Width           =   1605
   End
   Begin VB.TextBox txtRoadPermitDate 
      Height          =   315
      Left            =   6000
      TabIndex        =   30
      Top             =   6300
      Width           =   1635
   End
   Begin VB.TextBox txtGrDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   1635
      Width           =   1965
   End
   Begin VB.CheckBox chkCartage 
      BackColor       =   &H80000013&
      Caption         =   "Cartage"
      Height          =   255
      Left            =   8430
      TabIndex        =   27
      Top             =   5970
      Width           =   915
   End
   Begin VB.CheckBox chkUnloading 
      BackColor       =   &H80000013&
      Caption         =   "Unloading"
      Height          =   255
      Left            =   7335
      TabIndex        =   26
      Top             =   5970
      Width           =   1095
   End
   Begin VB.CheckBox chkLoading 
      BackColor       =   &H80000013&
      Caption         =   "Loading"
      Height          =   255
      Left            =   6360
      TabIndex        =   25
      Top             =   5970
      Width           =   975
   End
   Begin VB.TextBox txttaxper 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8835
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   5400
      Width           =   525
   End
   Begin VB.TextBox txttaxamount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9360
      TabIndex        =   20
      Top             =   5400
      Width           =   1605
   End
   Begin VB.TextBox txtRoadPermitNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   29
      Top             =   6300
      Width           =   1965
   End
   Begin VB.TextBox txtModeOfTransportation 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   24
      Top             =   5940
      Width           =   1965
   End
   Begin VB.TextBox txtTimeOfRemovalOfGoods 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   2205
      Width           =   1965
   End
   Begin VB.TextBox txtSaleStockTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   1965
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   360
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4290
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6780
      Width           =   1095
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9360
      TabIndex        =   19
      Top             =   5115
      Width           =   1605
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   9360
      TabIndex        =   32
      Top             =   7110
      Width           =   1605
   End
   Begin VB.TextBox txtLoadingUnloading 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   9360
      TabIndex        =   28
      Top             =   5970
      Width           =   1605
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   9360
      TabIndex        =   23
      Top             =   5685
      Width           =   1605
   End
   Begin VB.TextBox txtOtherCharges 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FDD7&
      Height          =   285
      Left            =   9360
      TabIndex        =   31
      Top             =   6255
      Width           =   1605
   End
   Begin VB.TextBox txtBookNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   465
      Width           =   1005
   End
   Begin VB.TextBox txtinvoiceno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   180
      Width           =   1005
   End
   Begin VB.TextBox txtVehicleno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   1065
      Width           =   1965
   End
   Begin VB.TextBox txtGrNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   1350
      Width           =   1965
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3300
      TabIndex        =   18
      Top             =   5100
      Width           =   4905
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   2295
      Left            =   180
      TabIndex        =   17
      Top             =   2700
      Width           =   11115
      _cx             =   19606
      _cy             =   4048
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmsaleinvoice.frx":0442
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
      Editable        =   2
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
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   5340
      TabIndex        =   40
      Top             =   60
      Width           =   5955
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1620
         Width           =   4395
      End
      Begin VB.TextBox txtfax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1320
         Width           =   4395
      End
      Begin VB.TextBox txtphone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1020
         Width           =   4395
      End
      Begin VB.TextBox txtstate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   11
         Top             =   720
         Width           =   1755
      End
      Begin VB.TextBox txtaddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   420
         Width           =   4395
      End
      Begin VB.TextBox txtcity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   120
         Width           =   4395
      End
      Begin VB.TextBox txtTinNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   1920
         Width           =   4395
      End
      Begin VB.TextBox txttax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   2220
         Width           =   4395
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Email "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   46
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "Tin No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFCB97&
         BackStyle       =   0  'Transparent
         Caption         =   "CST No :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2220
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3195
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6780
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6780
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6780
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6780
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DtpInvoice 
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   750
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39498
   End
   Begin MSComCtl2.DTPicker dtpTransportationDate 
      Height          =   315
      Left            =   6420
      TabIndex        =   33
      Top             =   7440
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39498
   End
   Begin MSComCtl2.DTPicker dtpRoadPermitDate 
      Height          =   315
      Left            =   6420
      TabIndex        =   34
      Top             =   7740
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39498
   End
   Begin MSComCtl2.DTPicker DtpGr 
      Height          =   315
      Left            =   60
      TabIndex        =   68
      Top             =   1620
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39498
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax-Invoice No./Sale-Invoice No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   285
      TabIndex        =   74
      Top             =   5550
      Width           =   2865
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   73
      Top             =   5550
      Width           =   495
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   72
      Top             =   6600
      Width           =   1140
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Roundoff"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   70
      Top             =   6840
      Width           =   1065
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   66
      Top             =   6300
      Width           =   555
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Road Permit No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1695
      TabIndex        =   65
      Top             =   6300
      Width           =   1455
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Mode Of Transportation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1140
      TabIndex        =   64
      Top             =   5970
      Width           =   2010
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Of Removal Of Goods"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   63
      Top             =   2265
      Width           =   2340
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale / Stock Transfer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   615
      TabIndex        =   62
      Top             =   1980
      Width           =   1830
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Book No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1725
      TabIndex        =   61
      Top             =   465
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1530
      TabIndex        =   60
      Top             =   180
      Width           =   915
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "GR Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1725
      TabIndex        =   59
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Charges"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8040
      TabIndex        =   58
      Top             =   6300
      Width           =   1260
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading,Unloading Cartage"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8940
      TabIndex        =   57
      Top             =   7800
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "GR No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1905
      TabIndex        =   56
      Top             =   1380
      Width           =   540
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8220
      TabIndex        =   55
      Top             =   7140
      Width           =   1005
   End
   Begin VB.Label LBLTAX 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "CST %"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8220
      TabIndex        =   54
      Top             =   5460
      Width           =   600
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8490
      TabIndex        =   53
      Top             =   5160
      Width           =   810
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   52
      Top             =   5100
      Width           =   915
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1545
      TabIndex        =   51
      Top             =   1065
      Width           =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   50
      Top             =   750
      Width           =   405
   End
End
Attribute VB_Name = "frmsaleinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim ch As New ADODB.Recordset
Dim Str As String
Dim yrs As String
Dim FSO As New FileSystemObject
Dim a As Integer
Dim str1 As String
Dim H As Integer
Dim control1 As Control
Sub typeofcontrol()
'Dim O As Object
'For Each O In Me
'If TypeOf O Is VSFlexGrid Then
'O.Clear
'setVsWidth
'End If
'Next
'
'For Each O In Me
'If TypeOf O Is TextBox Then
'O.Text = ""
'End If
'Next
End Sub

Private Sub chkCartage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub chkLoading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub chkUnloading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub CmdPrint_Click()
    
If MsgBox("Do you want to Print", vbInformation + vbYesNo) = vbYes Then
                If txtinvoiceno.Text = "" Then
                MsgBox "enter Invoice No."
                txtinvoiceno.SetFocus
                Exit Sub
                
                ElseIf txtBookNo.Text = "" Then
                MsgBox "enter Book No."
                txtBookNo.SetFocus
                Exit Sub
                
                ElseIf txtname.Text = "" Then
                MsgBox "Select Customer Name"
                txtname.SetFocus
                Exit Sub
                
                Else
                save
                End If
    
    
    
    
    cr1.Reset
    Select Case INVOICETYPE
Case "SALEINVOICE"
  ' Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
   cr1.ReportFileName = App.Path & "\saleinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "TAXINVOICE"
   ' Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "TAX INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "BILL"
   ' Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "BILL" & "'"
   'cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\bill.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "CASH MEMO"
   ' Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "CASH MEMO" & "'"
   'cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\cashmemo.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "CHALLAN/TRANSFER INVOICE"
   ' Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "CHALLAN/TRANSFER INVOICE" & "'"
   'cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\CHALLAN.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "PERFORMAINVOICE"
   
   cr1.Formulas(0) = "INVOICENAME='" & "PERFORMA INVOICE" & "'"
   'cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\PERFORMA INVOICE.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno

End Select
   cr1.WindowState = crptMaximized
   cr1.Action = 1
   
   
End If
End Sub

Private Sub Command1_Click()
Dim O As Object
    For Each O In Me
          If TypeOf O Is TextBox Then
          O.Text = ""
          End If
      Next
      Record = ""
      txtname.SetFocus
      popuplist11 "Select name from Customer order by Name", con, , True, colmn

vs.Clear
setVsWidth
txttaxamount.Text = ""
txtSubTotal.Text = ""
txtTotal.Text = ""
txtLoadingUnloading.Text = ""
txtOtherCharges.Text = ""
txtTotalAmount.Text = ""
DtpInvoice.Value = Date
chkLoading.Value = 0
chkUnloading.Value = 0
chkCartage.Value = 0

Form_Load
maxNo


End Sub

Private Sub Command2_Click()
      
      If txtinvoiceno.Text = "" Then
         MsgBox "Please Search Records !!", vbInformation
         Exit Sub
      End If
      
If MsgBox("Do U Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
           
      Select Case INVOICETYPE
      
      Case "SALEINVOICE"
      con.Execute "delete from salesinvoice where invoiceno=" & txtinvoiceno.Text & ""
       
      Case "TAXINVOICE"
      con.Execute "delete from taxinvoice where invoiceno=" & txtinvoiceno.Text & ""
      
      Case "BILL"
      con.Execute "delete from bill where invoiceno=" & txtinvoiceno.Text & ""
      
      Case "CASH MEMO"
      con.Execute "delete from cashMemo where invoiceno=" & txtinvoiceno.Text & ""
      
      Case "CHALLAN/TRANSFER INVOICE"
      con.Execute "delete from ChallanTransferInvoice where invoiceno=" & txtinvoiceno.Text & ""
      
      Case "PERFORMAINVOICE"
      con.Execute "delete from PERFORMAInvoice where invoiceno=" & txtinvoiceno.Text & ""
      
      End Select
      'bookMaxNo
 Call Command1_Click
End If

'      If rs.RecordCount >= 1 Then
'        If MsgBox("Do U Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
'             rs.Delete adAffectCurrent
'        End If
'      End If
       
      maxNo
End Sub

Private Sub cmdSave_Click()
If MsgBox("Do U Want To Save !!!", vbInformation + vbYesNo) = vbYes Then
                If txtinvoiceno.Text = "" Then
                MsgBox "enter Invoice No."
                txtinvoiceno.SetFocus
                Exit Sub
                
                ElseIf txtBookNo.Text = "" Then
                MsgBox "enter Book No."
                txtBookNo.SetFocus
                Exit Sub
                
                ElseIf txtname.Text = "" Then
                MsgBox "Select Customer Name"
                txtname.SetFocus
                Exit Sub
                
                Else
                save
                End If
                
End If
End Sub


Sub save()

Set rs = New ADODB.Recordset
Select Case INVOICETYPE

Case "SALEINVOICE"
str1 = "SALEINVOICE"
rs.Open "Select * from salesinvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from salesinvoice where INVOICENO=" & txtinvoiceno.Text & ""
End If
Case "TAXINVOICE"
str1 = "TAXINVOICE"
rs.Open "Select * from taxinvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from taxinvoice where INVOICENO=" & txtinvoiceno.Text & ""
End If
Case "BILL"
str1 = "BILL"
rs.Open "Select * from bill where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from bill where INVOICENO=" & txtinvoiceno.Text & ""
End If
Case "CASH MEMO"
str1 = "CASH MEMO"
rs.Open "Select * from cashMemo where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from cashMemo where INVOICENO=" & txtinvoiceno.Text & ""
End If
Case "CHALLAN/TRANSFER INVOICE"
str1 = "CHALLAN/TRANSFER INVOICE"
rs.Open "Select * from ChallanTransferInvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from ChallanTransferInvoice where INVOICENO=" & txtinvoiceno.Text & ""
End If
Case "PERFORMAINVOICE"
str1 = "PERFORMAINVOICE"
rs.Open "Select * from performaInvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from performaInvoice where INVOICENO=" & txtinvoiceno.Text & ""
End If
End Select
For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 1) <> "" Then
'                     If MsgBox("Do U Want To Save !!!", vbInformation + vbYesNo) = vbYes Then

rs.AddNew

        rs!invoiceno = txtinvoiceno.Text
        rs!bookno = txtBookNo.Text
        rs!invoicedate = DtpInvoice.Value
        rs!vehicleno = txtVehicleno.Text
        rs!taxperc = Val(txttaxper.Text)
        rs!GRANDTOTAL = Val(txtTotalAmount.Text)
        rs!TOTALWRO = Val(tot1.Text)
        rs!AMTROUNDOFF = Val(txtRoundoff.Text)
        
        rs!remarks = txtRemarks.Text
        rs!custName = txtname.Text
        rs!Subtotal = Val(txtSubTotal.Text)
        rs!grno = txtGrNo.Text
        rs!grdate = txtGrDate.Text
        rs!tinno = txtTinNo.Text
        rs!othecharges = Val(txtOtherCharges.Text)
       
        
        If chkLoading.Value = 1 Or chkUnloading.Value = 1 Or chkCartage.Value = 1 Then
        rs!cartage = Val(txtLoadingUnloading.Text)
        End If
           
        
        rs!loadingYN = chkLoading.Value
        rs!UnloadingYN = chkUnloading.Value
        rs!CartageYN = chkCartage.Value
 
 
    If str1 = "CHALLAN/TRANSFER INVOICE" Then
       
        rs!SaleStockTransfer = txtSaleStockTransfer.Text
        rs!TimeOfRemovalOfGoods = txtTimeOfRemovalOfGoods.Text
        rs!TaxSaleInvoiceNo = txtTaxSaleInvoiceNo.Text
        rs!ModeOfTransportation = txtModeOfTransportation.Text
        rs!TransportationDate = txtTransportationDate.Text
        rs!RoadPermitNo = txtRoadPermitNo.Text
        rs!RoadPermitDate = txtRoadPermitDate.Text
    End If
    
    rs!itemname = vs.TextMatrix(i, 1)
    rs!itemdesc = vs.TextMatrix(i, 2)
    rs!qty = Val(vs.TextMatrix(i, 3))
    rs!itemamount = Val(vs.TextMatrix(i, 4))
    rs!itemtotal = Val(vs.TextMatrix(i, 5))

rs.Update
End If
'End If
Next
'""""""""""""ADD NEW CUSTOMER OR REPLACE THE EXISING ONE"""""""""""""""

 Set rs = New ADODB.Recordset

                
              If rs.State = 1 Then rs.Close
              rs.Open "Select * from Customer where Name='" & txtname.Text & "'", con, adOpenDynamic, adLockOptimistic
                  If rs.RecordCount >= 1 Then
                        rs!Name = txtname.Text
                        rs!Address = txtaddress.Text
                        rs!CITY = txtcity.Text
                        rs!State = txtstate.Text
                        rs!phone = txtphone.Text
                        rs!FAX = txtfax.Text
                        rs!email = txtemail.Text
                        rs!cst = txttax.Text
                        rs!tin = txtTinNo.Text
                        
                        rs.Update
                   Else
                        rs.AddNew
                        rs!Name = txtname.Text
                        rs!Address = txtaddress.Text
                        rs!CITY = txtcity.Text
                        rs!State = txtstate.Text
                        rs!phone = txtphone.Text
                        rs!FAX = txtfax.Text
                        rs!email = txtemail.Text
                        rs!cst = txttax.Text
                        rs!tin = txtTinNo.Text
                        
                        rs.Update
                     End If

End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub dp1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub dp2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub Command3_Click()
 cr1.Reset
      cr1.Formulas(2) = "pcopy='" & "EXTRA COPY" & "'"
    Select Case INVOICETYPE
Case "SALEINVOICE"
   'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
   cr1.ReportFileName = App.Path & "\saleinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "TAXINVOICE"
    'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "TAX INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "BILL"
   Exit Sub
Case "CASH MEMO"
    Exit Sub
Case "CHALLAN/TRANSFER INVOICE"
    'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "CHALLAN/TRANSFER INVOICE" & "'"
   'cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\CHALLAN.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "PERFORMAINVOICE"
Exit Sub
End Select
   cr1.WindowState = crptMaximized
   cr1.Action = 1
End Sub

Private Sub Command4_Click()
 
 If MsgBox("Do you want to Print", vbInformation + vbYesNo) = vbYes Then
                If txtinvoiceno.Text = "" Then
                MsgBox "enter Invoice No."
                txtinvoiceno.SetFocus
                Exit Sub
                
                ElseIf txtBookNo.Text = "" Then
                MsgBox "enter Book No."
                txtBookNo.SetFocus
                Exit Sub
                
                ElseIf txtname.Text = "" Then
                MsgBox "Select Customer Name"
                txtname.SetFocus
                Exit Sub
                
                Else
                save
                End If
    
    
    Select Case INVOICETYPE
    Case "SALEINVOICE"
   
       cr1.Reset
       cr1.Formulas(2) = "pcopy='" & "ORIGINAL" & "'"
       cr1.Formulas(0) = "INVOICENAME='" & "TAX INVOICE" & "'"
       cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
       cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
       cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
       cr1.Destination = crptToPrinter
       cr1.Action = 0

       cr1.Reset
       cr1.Formulas(2) = "pcopy='" & "DUPLICATE" & "'"
       cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
       cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
       cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
       cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
       cr1.Destination = crptToPrinter
       cr1.Action = 0

       cr1.Reset
       cr1.Formulas(2) = "pcopy='" & "TRIPLICATE" & "'"
       cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
       cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
       cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
       cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
       cr1.Destination = crptToPrinter
       cr1.Action = 0

   
Case "TAXINVOICE"
   
   
   cr1.Reset
   cr1.Formulas(2) = "pcopy='" & "ORIGINAL" & "'"
   cr1.Formulas(0) = "INVOICENAME='" & "TAX INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
   cr1.Destination = crptToPrinter
   cr1.Action = 0

   cr1.Reset
   cr1.Formulas(2) = "pcopy='" & "DUBLICATE" & "'"
   cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
   cr1.Destination = crptToPrinter
   cr1.Action = 0

   cr1.Reset
   cr1.Formulas(2) = "pcopy='" & "TRIPLICATE" & "'"
   cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
   cr1.Destination = crptToPrinter
   cr1.Action = 0
   
Case "BILL"
   Exit Sub
Case "CASH MEMO"
    Exit Sub
Case "CHALLAN/TRANSFER INVOICE"
   
   
   cr1.Reset
   cr1.Formulas(2) = "pcopy='" & "ORIGINAL" & "'"
   cr1.Formulas(0) = "INVOICENAME='" & "CHALLAN/TRANSFER INVOICE" & "'"
   cr1.ReportFileName = App.Path & "\CHALLAN.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
   cr1.Destination = crptToPrinter
   cr1.Action = 0
   
   cr1.Reset
   cr1.Formulas(2) = "pcopy='" & "DUBLICATE" & "'"
   cr1.Formulas(0) = "INVOICENAME='" & "CHALLAN/TRANSFER INVOICE" & "'"
   cr1.ReportFileName = App.Path & "\CHALLAN.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
   cr1.Destination = crptToPrinter
   cr1.Action = 0
   
   
   cr1.Reset
   cr1.Formulas(2) = "pcopy='" & "TRIPLICATE" & "'"
   cr1.Formulas(0) = "INVOICENAME='" & "CHALLAN/TRANSFER INVOICE" & "'"
   cr1.ReportFileName = App.Path & "\CHALLAN.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
   cr1.Destination = crptToPrinter
   cr1.Action = 0
   
 
Case "PERFORMAINVOICE"
Exit Sub
End Select

End If
End Sub

Private Sub Commandprintduplicate_Click()
 
 If MsgBox("Do you want to Print", vbInformation + vbYesNo) = vbYes Then
                If txtinvoiceno.Text = "" Then
                MsgBox "enter Invoice No."
                txtinvoiceno.SetFocus
                Exit Sub
                
                ElseIf txtBookNo.Text = "" Then
                MsgBox "enter Book No."
                txtBookNo.SetFocus
                Exit Sub
                
                ElseIf txtname.Text = "" Then
                MsgBox "Select Customer Name"
                txtname.SetFocus
                Exit Sub
                
                Else
                save
                End If
    
 
 
 cr1.Reset
      cr1.Formulas(2) = "pcopy='" & "DUPLICATE" & "'"
    Select Case INVOICETYPE
Case "SALEINVOICE"
   'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
   cr1.ReportFileName = App.Path & "\saleinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "TAXINVOICE"
    'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "TAX INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "BILL"
   Exit Sub
Case "CASH MEMO"
    Exit Sub
Case "CHALLAN/TRANSFER INVOICE"
    'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "CHALLAN/TRANSFER INVOICE" & "'"
   'cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\CHALLAN.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "PERFORMAINVOICE"
Exit Sub
End Select
   cr1.WindowState = crptMaximized
   cr1.Action = 1
   
   
End If

End Sub

Private Sub Commandprintoriginal_Click()
 
 If MsgBox("Do you want to Print", vbInformation + vbYesNo) = vbYes Then
                If txtinvoiceno.Text = "" Then
                MsgBox "enter Invoice No."
                txtinvoiceno.SetFocus
                Exit Sub
                
                ElseIf txtBookNo.Text = "" Then
                MsgBox "enter Book No."
                txtBookNo.SetFocus
                Exit Sub
                
                ElseIf txtname.Text = "" Then
                MsgBox "Select Customer Name"
                txtname.SetFocus
                Exit Sub
                
                Else
                save
                End If
    
    
    cr1.Reset
      cr1.Formulas(2) = "pcopy='" & "ORIGINAL" & "'"
    Select Case INVOICETYPE
Case "SALEINVOICE"
  ' Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
   cr1.ReportFileName = App.Path & "\saleinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "TAXINVOICE"
  '  Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "TAX INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "BILL"
   Exit Sub
Case "CASH MEMO"
    Exit Sub
Case "CHALLAN/TRANSFER INVOICE"
   cr1.Formulas(0) = "INVOICENAME='" & "CHALLAN/TRANSFER INVOICE" & "'"
   cr1.ReportFileName = App.Path & "\CHALLAN.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "PERFORMAINVOICE"
Exit Sub
End Select
   cr1.WindowState = crptMaximized
   cr1.Action = 1

End If
End Sub

Private Sub Commandprinttriplicate_Click()
If MsgBox("Do you want to Print", vbInformation + vbYesNo) = vbYes Then
                If txtinvoiceno.Text = "" Then
                MsgBox "enter Invoice No."
                txtinvoiceno.SetFocus
                Exit Sub
                
                ElseIf txtBookNo.Text = "" Then
                MsgBox "enter Book No."
                txtBookNo.SetFocus
                Exit Sub
                
                ElseIf txtname.Text = "" Then
                MsgBox "Select Customer Name"
                txtname.SetFocus
                Exit Sub
                
                Else
                save
                End If
     
 
 
 cr1.Reset
      cr1.Formulas(2) = "pcopy='" & "TRIPLICATE" & "'"
    Select Case INVOICETYPE
Case "SALEINVOICE"
   'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
   cr1.ReportFileName = App.Path & "\saleinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "TAXINVOICE"
    'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "TAX INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "BILL"
   Exit Sub
Case "CASH MEMO"
    Exit Sub
Case "CHALLAN/TRANSFER INVOICE"
    'Call cmdSave_Click
   cr1.Formulas(0) = "INVOICENAME='" & "CHALLAN/TRANSFER INVOICE" & "'"
   'cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\CHALLAN.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.Invoiceno} = " & frmsaleinvoice.txtinvoiceno
Case "PERFORMAINVOICE"
Exit Sub
End Select
   cr1.WindowState = crptMaximized
   cr1.Action = 1
   
   
End If
End Sub

Private Sub DtpGr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub DtpInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub dtpRoadPermitDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub dtpTransportationDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Activate()
'txtINVOICENO.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
       Unload Me
'    ElseIf KeyCode = 113 Then
'       popuplist11 "Select name from Customer order by Name", con, , True, colmn
'       popuplist1.Show
    End If
End Sub

Private Sub Form_Load()

Select Case INVOICETYPE
Case "SALEINVOICE"

txttaxper.Text = ""
If rs.State = 1 Then rs.Close
rs.Open "select cstper from setup1", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
txttaxper.Text = rs!cstper
End If

Case "TAXINVOICE"
If rs.State = 1 Then rs.Close
rs.Open "select vatper from setup1", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
txttaxper.Text = ""
txttaxper.Text = rs!vatper
End If

End Select

Dim s As String
s = ""
If rs.State = 1 Then rs.Close
rs.Open "SELECT distinct itemname FROM ItemMaster", con
While rs.EOF = False
If s = "" Then
s = rs.Fields(0).Value
Else
s = s & "|" & rs.Fields(0).Value
End If
rs.MoveNext
Wend
vs.ColComboList(1) = s

'''---------------------------------------------------
'Me.Top = 1500
'Me.Left = 2000

Record = ""
colmn = 1
popuplist11 "Select name from Customer order by Name", con, , True, colmn

'LblName.Caption = "Customer Details"
tbl = "Customer"
pos Me
setVsWidth
a = 0
sNo (a)

DtpInvoice.Value = Date
maxNo
'txtinvoiceno.SetFocus
End Sub


Private Sub txtaddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtBookNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtcity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txttaxno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub
Private Sub txtGrDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtinvoiceno_DblClick()
Select Case INVOICETYPE
Case "SALEINVOICE"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate  from salesinvoice", con
Case "TAXINVOICE"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from TAXinvoice", con
Case "BILL"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from bill", con
Case "CASH MEMO"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from cashMemo", con
Case "CHALLAN/TRANSFER INVOICE"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from ChallanTransferInvoice", con
Case "PERFORMAINVOICE"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from performaInvoice", con

End Select
End Sub

Private Sub txtinvoiceno_LostFocus()
count1 = 0
count1 = (Val(txtinvoiceno) / 50)
If count1 > Int(count1) Then
count1 = count1 + 1
End If
txtBookNo.Text = Int(count1)
End Sub

Private Sub txtLoadingUnloading_Change()
calc
End Sub

Private Sub txtLoadingUnloading_GotFocus()
txtLoadingUnloading.SelLength = 20

End Sub

Private Sub txtModeOfTransportation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtname_DblClick()
   popuplist2 "Select name from Customer order by Name", con

End Sub

Private Sub txtOtherCharges_Change()
calc
End Sub



Private Sub txtOtherCharges_GotFocus()
txtOtherCharges.SelLength = 20

End Sub

Private Sub txtRoadPermitDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtRoadPermitNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtSaleStockTransfer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtSubTotal_Change()
'calc
End Sub

Private Sub txttax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txttaxamount_Change()
calc
End Sub

Private Sub txttaxamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txttaxper_GotFocus()

If INVOICETYPE = "SALEINVOICE" Then

txttaxper.Text = ""
If rs.State = 1 Then rs.Close
rs.Open "select cstper from setup1", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
txttaxper.Text = rs!cstper
End If

ElseIf INVOICETYPE = "TAXINVOICE" Then
txttaxper.Text = ""
If rs.State = 1 Then rs.Close
rs.Open "select vatper from setup1", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
txttaxper.Text = rs!vatper
End If
End If
calc
End Sub

Private Sub txttaxper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtfax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtfinanYear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtGrNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtLoadingUnloading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtname_GotFocus()
If PopUpValue1 <> "" Then
txtname.Text = PopUpValue1
 SearchData
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
PopUpValue3 = ""
End If
 

End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{Tab}"
   SearchData
End If
If KeyCode = 113 Then
'       popuplist2 "Select name from Customer order by Name", con, , True, colmn
       popuplist2 "Select name from Customer order by Name", con
       
'       popuplist1.Show
    End If
End Sub

Sub SearchData()
If rs.State = 1 Then rs.Close
rs.Open "Select * from Customer where Name='" & txtname.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
   
   
   txtname.Text = rs!Name
   txtaddress.Text = rs!Address
   txtcity.Text = rs!CITY
   txtstate.Text = rs!State
   txtphone.Text = rs!phone
   txtfax.Text = rs!FAX
   txtemail.Text = rs!email
   txttax.Text = rs!cst & ""
   txtTinNo.Text = rs!tin & ""
  
End If

End Sub



Sub Search()
H = 0
vs.Clear
setVsWidth

Set rs = New ADODB.Recordset
Select Case INVOICETYPE
Case "SALEINVOICE"
str1 = "SALEINVOICE"
rs.Open "Select * from salesinvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
Case "TAXINVOICE"
str1 = "TAXINVOICE"
rs.Open "Select * from TAXinvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
Case "BILL"
str1 = "BILL"
rs.Open "Select * from bill where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
Case "CASH MEMO"
str1 = "CASH MEMO"
rs.Open "Select * from cashMemo where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
Case "CHALLAN/TRANSFER INVOICE"
str1 = "CHALLAN/TRANSFER INVOICE"
rs.Open "Select * from ChallanTransferInvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
Case "PERFORMAINVOICE"
str1 = "PERFORMAINVOICE"
rs.Open "Select * from performaInvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic

End Select
If rs.EOF = False Then
                txtinvoiceno.Text = rs!invoiceno
                txtBookNo.Text = rs!bookno
                DtpInvoice.Value = rs!invoicedate
                txtVehicleno.Text = rs!vehicleno
                txttaxper.Text = rs!taxperc
                txtTotalAmount.Text = rs!GRANDTOTAL
                txtRoundoff.Text = rs!AMTROUNDOFF
                tot1.Text = rs!TOTALWRO
                txtRemarks.Text = rs!remarks
                txtname.Text = rs!custName
                txtSubTotal.Text = rs!Subtotal
                txtGrNo.Text = rs!grno
                If rs!tinno <> "" Then txtTinNo.Text = rs!tinno
                txtGrDate.Text = rs!grdate
                'DtpGr.Value = rs!grdate
                txtLoadingUnloading.Text = rs!cartage
                txtOtherCharges.Text = rs!othecharges
                        
                
                If str1 <> "CHALLAN/TRANSFER INVOICE" Then
                If rs!loadingYN = True Then
                chkLoading.Value = 1
                Else
                chkLoading.Value = 0
                End If
                If rs!UnloadingYN = True Then
                chkUnloading.Value = 1
                Else
                chkUnloading.Value = 0
                End If
                 If rs!CartageYN = True Then
                 chkCartage.Value = 1
                Else
                chkCartage.Value = 0
                End If
      End If
      
      If str1 = "CHALLAN/TRANSFER INVOICE" Then
                txtSaleStockTransfer.Text = rs!SaleStockTransfer
                txtTimeOfRemovalOfGoods.Text = rs!TimeOfRemovalOfGoods
                txtTaxSaleInvoiceNo.Text = rs!TaxSaleInvoiceNo
                txtModeOfTransportation.Text = rs!ModeOfTransportation
                txtTransportationDate.Text = rs!TransportationDate
                txtRoadPermitNo.Text = rs!RoadPermitNo
                txtRoadPermitDate.Text = rs!RoadPermitDate
        End If
    
    
End If

Set rs = New ADODB.Recordset
Select Case INVOICETYPE
Case "SALEINVOICE"
rs.Open "Select * from salesinvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic

Case "TAXINVOICE"
rs.Open "Select * from TAXinvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic

Case "BILL"
rs.Open "Select * from bill where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic

Case "CASH MEMO"
rs.Open "Select * from cashMemo where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic

Case "CHALLAN/TRANSFER INVOICE"
rs.Open "Select * from ChallanTransferInvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
Case "PERFORMAINVOICE"
rs.Open "Select * from performaInvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic

End Select


For i = 1 To rs.RecordCount
                
                H = H + 1
                vs.TextMatrix(i, 0) = H
                If rs!itemname <> "" Then vs.TextMatrix(i, 1) = rs!itemname
               If rs!itemdesc <> "" Then vs.TextMatrix(i, 2) = rs!itemdesc
                If rs!qty <> "" Then vs.TextMatrix(i, 3) = rs!qty
                If rs!itemamount <> "" Then vs.TextMatrix(i, 4) = rs!itemamount
               If rs!itemtotal <> "" Then vs.TextMatrix(i, 5) = rs!itemtotal
                
rs.MoveNext

Next



If rs.State = 1 Then rs.Close
rs.Open "Select * from Customer where Name='" & txtname.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
   txtname.Text = rs!Name
   txtaddress.Text = rs!Address
   txtcity.Text = rs!CITY
   txtstate.Text = rs!State
   txtphone.Text = rs!phone
   txtfax.Text = rs!FAX
   txtemail.Text = rs!email
   txttax.Text = rs!cst
   txtTinNo.Text = rs!tin
  
End If
calc
str1 = ""
End Sub


Private Sub txtPen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtOtherCharges_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtphone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtreg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtservice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtINVOICENO_GotFocus()
If PopUpValue1 <> "" Then
txtinvoiceno.Text = PopUpValue1
Search
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
PopUpValue3 = ""
End If
End Sub

Private Sub txtINVOICENO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

If KeyCode = 113 Then
Select Case INVOICETYPE
Case "SALEINVOICE"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate  from salesinvoice", con
Case "TAXINVOICE"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from TAXinvoice", con
Case "BILL"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from bill", con
Case "CASH MEMO"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from cashMemo", con
Case "CHALLAN/TRANSFER INVOICE"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from ChallanTransferInvoice", con
Case "PERFORMAINVOICE"
popuplist2 "SELECT distinct Invoiceno,custName,bookno,invoicedate from performaInvoice", con

End Select

End If

End Sub

Private Sub txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Sub calc()

If txtSubTotal.Text = "" Then
txtSubTotal.Text = 0
End If

If txttaxper.Text = "" Then
txttaxper.Text = 0
End If
If txttaxamount.Text = "" Then
txttaxamount.Text = 0
End If
If txtTotal.Text = "" Then
txtTotal.Text = 0
End If
If txtLoadingUnloading.Text = "" Then
txtLoadingUnloading.Text = 0
End If
If txtOtherCharges.Text = "" Then
txtOtherCharges.Text = 0
End If
If tot1.Text = "" Then
tot1.Text = 0
End If
If txtRoundoff.Text = "" Then
txtRoundoff.Text = 0
End If
If txtTotalAmount.Text = "" Then
txtTotalAmount.Text = 0
End If

txtSubTotal.Text = 0
txtVehicleno.Text = UCase(txtVehicleno.Text)
For i = 1 To vs.Rows - 1
        txtSubTotal.Text = Val(txtSubTotal.Text) + Val(vs.TextMatrix(i, 5))
Next
If Val(txttaxper.Text) > 0 Then txttaxamount.Text = Format(((Val(txtSubTotal.Text) * Val(txttaxper.Text) / 100)), "0.00")
txtTotal.Text = Format(Val(txttaxamount.Text) + Val(txtSubTotal.Text), "0.00")
tot1 = Val(txtTotal.Text) + Val(txtLoadingUnloading.Text) + Val(txtOtherCharges.Text)
tot1 = Format(Val(tot1.Text), "0.00")
txtTotalAmount.Text = Round(Format(Val(tot1.Text), "0.00"))
txtRoundoff.Text = Format((Val(txtTotalAmount.Text) - Val(tot1.Text)), "0.00")

For i = 0 To vs.Rows - 1
If vs.TextMatrix(vs.RowSel, 1) <> "" Then
vs.TextMatrix(i, 3) = Format(vs.TextMatrix(i, 3), "0.000")
vs.TextMatrix(i, 4) = Format(vs.TextMatrix(i, 4), "0.00")
vs.TextMatrix(i, 5) = Format(vs.TextMatrix(i, 5), "0.00")
End If
Next

End Sub

Private Sub txtSubTotal_GotFocus()
calc
End Sub

Private Sub txtSubTotal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtTaxSaleInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtTimeOfRemovalOfGoods_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtTotal_Change()
calc
End Sub

Private Sub txtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtTotalAmount_Change()
calc
End Sub

Private Sub txtTotalAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtTransportationDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtVehicleno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtVehicleno_LostFocus()
calc
End Sub

Private Sub vs_ChangeEdit()
calc
End Sub


Private Sub vs_GotFocus()
vs.SetFocus

For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 1) = "" Then
SendKeys "{up}"
End If
Next


If vs.TextMatrix(1, 1) <> "" Then
For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 1) <> "" Then
SendKeys "{down}"
GoTo a1:
End If
Next

a1:
SendKeys "{down}"
End If
SendKeys "{up}"
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    For i = 1 To 5
        vs.TextMatrix(vs.RowSel, i) = ""
       Next
End If

If KeyCode = 13 Then
If vs.Col = 4 Then
SendKeys "{down}{home}"
sNo (a)
ElseIf vs.TextMatrix(vs.RowSel, 1) <> "" Then
SendKeys "{RIGHT}"
Else
SendKeys "{tab}"
End If
End If


If vs.Col = 1 Then

If KeyCode = 13 Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT rate FROM itemmaster where itemname='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
If rs.EOF = False Then
vs.TextMatrix(vs.RowSel, 4) = rs.Fields("rate")
calc
End If
End If
End If

If KeyCode = 13 Then
If vs.Col = 3 Then
vs.TextMatrix(vs.RowSel, 5) = Format(Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 4)), "0.00")
calc
End If
If vs.Col = 4 Then
vs.TextMatrix(vs.RowSel, 5) = Format(Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 4)), "0.00")
calc
End If
End If

If KeyCode = 113 Then
'popuplist2 "select itemname from itemmaster", con
End If
End Sub

Private Sub txtTinNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Sub setVsWidth()
vs.FormatString = "SL.NO.|DESCRIPTION|NO.& DESCRIPTION OF PACKAGE |QTY.|UNIT PRICE|VALUE OF GOODS"
vs.ColWidth(0) = 700
vs.ColWidth(1) = 3500
vs.ColWidth(2) = 3000
vs.ColWidth(3) = 1000
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 1500
End Sub

Sub sNo(q As Integer)
'a = a + 1
'vs.TextMatrix(a, 0) = a

If vs.TextMatrix(a + 1, 0) = "" Then
a = a + 1
vs.TextMatrix(a, 0) = a
End If
End Sub

Sub maxNo()

Set rs = New ADODB.Recordset
Select Case INVOICETYPE
Case "SALEINVOICE"
rs.Open "select max(invoiceno),max(bookno) from salesinvoice", con
If rs.EOF = False Then
End If
Case "TAXINVOICE"
rs.Open "select max(invoiceno),max(bookno)from taxinvoice", con
If rs.EOF = False Then
End If
Case "BILL"
rs.Open "select max(invoiceno),max(bookno)from bill", con
If rs.EOF = False Then
End If
Case "CASH MEMO"
rs.Open "select max(invoiceno),max(bookno) from cashMemo", con
If rs.EOF = False Then
End If
Case "CHALLAN/TRANSFER INVOICE"
rs.Open "select max(invoiceno),max(bookno) from ChallanTransferInvoice", con
If rs.EOF = False Then
End If
Case "PERFORMAINVOICE"
rs.Open "select max(invoiceno),max(bookno) from performaInvoice", con
If rs.EOF = False Then
End If
End Select


'If rs.State = 1 Then rs.Close
'rs.Open "select max(invoiceno) from salesinvoice", con
If IsNull(rs(0)) Then
txtinvoiceno.Text = 1
Else
txtinvoiceno.Text = (Val(rs(0)) + 1)
End If
End Sub


