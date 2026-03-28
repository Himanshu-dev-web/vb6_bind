VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmchallanTransferInvoice 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sale invoice"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtModeOfTransportation 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3420
      TabIndex        =   55
      Top             =   6180
      Width           =   1965
   End
   Begin VB.TextBox txtTaxSaleInvoiceNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7320
      TabIndex        =   53
      Top             =   5520
      Width           =   1965
   End
   Begin VB.TextBox txtTimeOfRemovalOfGoods 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2700
      TabIndex        =   51
      Top             =   1950
      Width           =   1965
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   60
      Top             =   6840
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
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txttaxamount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9900
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9900
      TabIndex        =   9
      Top             =   5475
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9900
      TabIndex        =   15
      Top             =   6900
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtLoadingUnloading 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9900
      TabIndex        =   13
      Top             =   6330
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9900
      TabIndex        =   12
      Top             =   6045
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtOtherCharges 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9900
      TabIndex        =   14
      Top             =   6615
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txttaxper 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtBookNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2700
      TabIndex        =   1
      Top             =   885
      Width           =   1965
   End
   Begin VB.TextBox txtinvoiceno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2700
      TabIndex        =   0
      Top             =   540
      Width           =   1965
   End
   Begin VB.TextBox txtSaleStockTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2700
      TabIndex        =   3
      Top             =   1605
      Width           =   1965
   End
   Begin VB.TextBox txtGrNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2100
      TabIndex        =   4
      Top             =   5505
      Width           =   1965
   End
   Begin VB.TextBox txtRoadPermitNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7260
      TabIndex        =   8
      Top             =   5880
      Width           =   2685
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   2595
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   11115
      _cx             =   19606
      _cy             =   4577
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      TabIndex        =   28
      Top             =   60
      Width           =   5955
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   1620
         Width           =   4395
      End
      Begin VB.TextBox txtfax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Top             =   1320
         Width           =   4395
      End
      Begin VB.TextBox txtphone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   1020
         Width           =   4395
      End
      Begin VB.TextBox txtstate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   23
         Top             =   720
         Width           =   1755
      End
      Begin VB.TextBox txtaddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Top             =   420
         Width           =   4395
      End
      Begin VB.TextBox txtcity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   120
         Width           =   4395
      End
      Begin VB.TextBox txtTinNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   1920
         Width           =   4395
      End
      Begin VB.TextBox txttax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   27
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
         TabIndex        =   37
         Top             =   1620
         Width           =   1095
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   120
         Width           =   1395
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   1920
         Width           =   1095
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
         TabIndex        =   29
         Top             =   2220
         Width           =   1095
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
      Left            =   2535
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6600
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
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
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
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DtpInvoice 
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   1230
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20250625
      CurrentDate     =   39498
   End
   Begin MSComCtl2.DTPicker DtpGr 
      Height          =   315
      Left            =   2100
      TabIndex        =   5
      Top             =   5790
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20250625
      CurrentDate     =   39498
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
      Left            =   1260
      TabIndex        =   56
      Top             =   6210
      Width           =   2010
   End
   Begin VB.Label Label17 
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
      Left            =   4305
      TabIndex        =   54
      Top             =   5550
      Width           =   2865
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   52
      Top             =   1980
      Width           =   2340
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
      Left            =   1830
      TabIndex        =   49
      Top             =   960
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
      Left            =   1695
      TabIndex        =   48
      Top             =   600
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
      Left            =   1230
      TabIndex        =   47
      Top             =   5880
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
      Left            =   8700
      TabIndex        =   46
      Top             =   6660
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading/Unloading Charges"
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
      Left            =   7500
      TabIndex        =   45
      Top             =   6360
      Visible         =   0   'False
      Width           =   2355
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
      Left            =   1410
      TabIndex        =   44
      Top             =   5580
      Width           =   540
   End
   Begin VB.Label Label18 
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
      Left            =   8700
      TabIndex        =   43
      Top             =   6960
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   8640
      TabIndex        =   42
      Top             =   5820
      Visible         =   0   'False
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
      Left            =   9000
      TabIndex        =   41
      Top             =   5520
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label15 
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
      Left            =   5820
      TabIndex        =   40
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Sale/Stock Transfer"
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
      Left            =   840
      TabIndex        =   39
      Top             =   1620
      Width           =   1710
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
      Left            =   2160
      TabIndex        =   38
      Top             =   1320
      Width           =   405
   End
End
Attribute VB_Name = "frmchallanTransferInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim Str As String
Dim yrs As String
Dim FSO As New FileSystemObject
Dim a As Integer


Private Sub cmdExit_Click()
 Unload Me
End Sub


Private Sub CmdPrint_Click()
    cr1.Reset
    Select Case INVOICETYPE
Case "SALEINVOICE"
   cr1.Formulas(0) = "INVOICENAME='" & "SALE INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "CST % " & "'"
   cr1.ReportFileName = App.Path & "\saleinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.saleInvoiceno} = " & frmsaleinvoice.txtinvoiceno

Case "TAXINVOICE"

   cr1.Formulas(0) = "INVOICENAME='" & "TAX INVOICE" & "'"
   cr1.Formulas(1) = "TAXNAME='" & "VAT % " & "'"
   cr1.ReportFileName = App.Path & "\taxinvoice.rpt"
   cr1.ReplaceSelectionFormula "{salesinvoice.saleInvoiceno} = " & frmsaleinvoice.txtinvoiceno

End Select
    
   cr1.WindowState = crptMaximized
   cr1.Action = 1
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
      End Select
      
End If
      
      
      
      
      
'
'      If rs.RecordCount >= 1 Then
'        If MsgBox("Do U Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
'             rs.Delete adAffectCurrent
'        End If
'      End If
      
      Call Command1_Click
      maxNo
End Sub

Private Sub cmdSave_Click()
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

Set rs = New ADODB.Recordset
Select Case INVOICETYPE
Case "SALEINVOICE"
rs.Open "Select * from salesinvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from salesinvoice where INVOICENO=" & txtinvoiceno.Text & ""
End If
Case "TAXINVOICE"
rs.Open "Select * from taxinvoice where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from taxinvoice where INVOICENO=" & txtinvoiceno.Text & ""
End If
Case "BILL"
rs.Open "Select * from bill where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from bill where INVOICENO=" & txtinvoiceno.Text & ""
End If
Case "CASH MEMO"
rs.Open "Select * from cashMemo where INVOICENO=" & txtinvoiceno.Text & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
con.Execute "delete from cashMemo where INVOICENO=" & txtinvoiceno.Text & ""
End If
End Select
For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 1) <> "" Then
rs.AddNew

        rs!INVOICENO = txtinvoiceno.Text
        rs!bookno = txtBookNo.Text
        rs!invoicedate = DtpInvoice.Value
        rs!vehicleno = txtVehicleno.Text
        rs!taxperc = Val(txttaxper.Text)
        rs!totalRate = Val(txtTotalAmount.Text)
        rs!remarks = txtRemarks.Text
        rs!custName = txtname.Text
        rs!Subtotal = Val(txtSubTotal.Text)
        rs!grno = txtGrNo.Text
        rs!tinno = txtTinNo.Text
        rs!grdate = DtpGr.Value
        rs!cartage = Val(txtLoadingUnloading.Text)
        rs!othecharges = Val(txtOtherCharges.Text)
    
    
    rs!itemname = vs.TextMatrix(i, 1)
    rs!itemdesc = vs.TextMatrix(i, 2)
    rs!qty = Val(vs.TextMatrix(i, 3))
    rs!itemamount = Val(vs.TextMatrix(i, 4))
    rs!Rate = Val(vs.TextMatrix(i, 5))

rs.Update
End If
Next

Call Command1_Click
vs.Clear
setVsWidth
             
End If
DtpInvoice.Value = Date
maxNo
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

a = 0
sNo (a)

DtpInvoice.Value = Date
maxNo
setVsWidth
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

Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txttaxno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtLoadingUnloading_Change()
calc
End Sub

Private Sub txtOtherCharges_Change()
calc
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
 SearchData
End Sub
Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{Tab}"
   SearchData
End If
If KeyCode = 113 Then
       popuplist11 "Select name from Customer order by Name", con, , True, colmn
       popuplist1.Show
    End If
End Sub

Sub SearchData()
If rs.State = 1 Then rs.Close
rs.Open "Select * from Customer where Name='" & Record & "'", con, adOpenDynamic, adLockOptimistic
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

vs.Clear
setVsWidth
 
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

End Select


If rs.EOF = False Then
 
  
                txtinvoiceno.Text = rs!INVOICENO
                txtBookNo.Text = rs!bookno
                DtpInvoice.Value = rs!invoicedate
                txtVehicleno.Text = rs!vehicleno
                txttaxper.Text = rs!taxperc
                txtTotalAmount.Text = rs!totalRate
                txtRemarks.Text = rs!remarks
                txtname.Text = rs!custName
                txtSubTotal.Text = rs!Subtotal
                txtGrNo.Text = rs!grno
                If rs!tinno <> "" Then txtTinNo.Text = rs!tinno
                DtpGr.Value = rs!grdate
                txtLoadingUnloading.Text = rs!cartage
                txtOtherCharges.Text = rs!othecharges
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

End Select


For i = 1 To rs.RecordCount

                vs.TextMatrix(i, 1) = rs!itemname
                vs.TextMatrix(i, 2) = rs!itemdesc
                vs.TextMatrix(i, 3) = rs!qty
                vs.TextMatrix(i, 4) = rs!itemamount
                vs.TextMatrix(i, 5) = rs!Rate

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
   txttax.Text = rs!cst & ""
   txtTinNo.Text = rs!tin & ""
  
End If
calc
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

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
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

End Select

End If

End Sub

Private Sub txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub


Sub calc()
txtSubTotal.Text = 0
For i = 1 To vs.Rows - 1
        txtSubTotal.Text = Val(txtSubTotal.Text) + Val(vs.TextMatrix(i, 5))
Next
If Val(txttaxper.Text) > 0 Then txttaxamount.Text = Format(((Val(txtSubTotal.Text) * Val(txttaxper.Text) / 100)), "0.00")
txtTotal.Text = Format(Val(txttaxamount.Text) + Val(txtSubTotal.Text), "0.00")
txtTotalAmount.Text = Format(Val(txtTotal.Text) + Val(txtLoadingUnloading.Text) + Val(txtOtherCharges.Text), "0.00")
End Sub

Private Sub txtSubTotal_GotFocus()
calc
End Sub

Private Sub txtSubTotal_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txtVehicleno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub txtVehicleno_LostFocus()
txtVehicleno.Text = UCase(txtVehicleno.Text)
End Sub

Private Sub vs_ChangeEdit()
calc
End Sub

Private Sub vs_GotFocus()

'If PopUpValue1 <> "" Then
'
'vs.TextMatrix(vs.roesel, 1) = PopUpValue1
'
'If rs.State = 1 Then rs.Close
'rs.Open "SELECT distinct itemname FROM ItemMaster", con
'If rs.EOF = False Then
'vs.TextMatrix(vs.RowSel, 1) = rs!itemname
'End If
'
'PopUpValue1 = ""
'PopUpValue2 = ""
'PopUpValue3 = ""
'PopUpValue3 = ""
'End If
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    For i = 0 To 5
        vs.TextMatrix(vs.RowSel, i) = ""
        a = a - 1
    Next
End If

If KeyCode = 13 Then
If vs.Col = 4 Then
SendKeys "{down}{home}"
sNo (a)
Else
SendKeys "{RIGHT}"
End If
End If

If KeyCode = 13 Then
If rs.State = 1 Then rs.Close
rs.Open "SELECT rate FROM itemmaster where itemname='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
If rs.EOF = False Then

vs.TextMatrix(vs.RowSel, 4) = rs.Fields("rate")
'vs.TextMatrix(vs.RowSel, 1) = rs.Fields("destinationcode")
calc
End If
End If

If KeyCode = 13 Then
If vs.Col = 3 Then

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
vs.FormatString = "SL.NO.|DESCRIPTION OF GOODS|QUANTITY |UNIT|ACTUAL / ASTIMATED VALUE OF GOODS"
vs.ColWidth(0) = 700
vs.ColWidth(1) = 4800
vs.ColWidth(2) = 1000
vs.ColWidth(3) = 900
vs.ColWidth(4) = 3300

End Sub

Sub sNo(q As Integer)
a = a + 1
vs.TextMatrix(a, 0) = a
End Sub

Sub maxNo()

Set rs = New ADODB.Recordset
Select Case INVOICETYPE
Case "SALEINVOICE"
rs.Open "select max(invoiceno) from salesinvoice", con
If rs.EOF = False Then
End If
Case "TAXINVOICE"
rs.Open "select max(invoiceno) from taxinvoice", con
If rs.EOF = False Then
End If
Case "BILL"
rs.Open "select max(invoiceno) from bill", con
If rs.EOF = False Then
End If
Case "CASH MEMO"
rs.Open "select max(invoiceno) from cashMemo", con
If rs.EOF = False Then
End If
Case "CHALLAN/TRANSFER INVOICE"
rs.Open "select max(invoiceno) from ChallanTransferInvoice", con
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
