VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmIssue 
   Caption         =   "Billing "
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13776
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10020
   ScaleWidth      =   13776
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDateOfSupp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   9108
      TabIndex        =   7
      Top             =   1620
      Width           =   2124
   End
   Begin VB.TextBox txtTransMode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5436
      TabIndex        =   6
      Top             =   1620
      Width           =   2448
   End
   Begin VB.TextBox txtVehiNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1872
      TabIndex        =   5
      Top             =   1620
      Width           =   2412
   End
   Begin VB.CheckBox Check1_igst 
      Caption         =   "Set IGST"
      Height          =   285
      Left            =   9135
      TabIndex        =   84
      Top             =   7296
      Width           =   1095
   End
   Begin VB.TextBox txtIGSTAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12300
      TabIndex        =   82
      Top             =   7272
      Width           =   1035
   End
   Begin VB.TextBox txtIGSTRate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11940
      TabIndex        =   81
      Top             =   7260
      Width           =   315
   End
   Begin VB.CheckBox Check1_ItemWiseTot 
      Caption         =   "Item Wise Total"
      Height          =   285
      Left            =   11655
      TabIndex        =   78
      Top             =   2904
      Width           =   1815
   End
   Begin VB.TextBox txtTotal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   9135
      TabIndex        =   76
      Text            =   "0"
      Top             =   6360
      Width           =   1035
   End
   Begin VB.TextBox txtTotalValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12300
      TabIndex        =   74
      Top             =   7896
      Width           =   1035
   End
   Begin VB.TextBox txtTGST 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12300
      TabIndex        =   72
      Top             =   7584
      Width           =   1035
   End
   Begin VB.TextBox txtAddLess 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12300
      TabIndex        =   70
      Top             =   8196
      Width           =   1035
   End
   Begin VB.ComboBox cboRCharge 
      Height          =   300
      ItemData        =   "frmIssue.frx":0000
      Left            =   5880
      List            =   "frmIssue.frx":000A
      TabIndex        =   3
      Top             =   936
      Width           =   735
   End
   Begin VB.TextBox txtPlaceofsupp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Text            =   "MEERUT"
      Top             =   936
      Width           =   2370
   End
   Begin VB.TextBox txtPAN_Shippto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7920
      TabIndex        =   65
      Top             =   2928
      Width           =   3390
   End
   Begin VB.TextBox txtGST_Shippto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5820
      TabIndex        =   64
      Top             =   2928
      Width           =   2070
   End
   Begin VB.TextBox txtParty_Shippto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5820
      MaxLength       =   60
      TabIndex        =   9
      Top             =   2280
      Width           =   5505
   End
   Begin VB.TextBox txtAdd_Shippto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5820
      MaxLength       =   60
      TabIndex        =   62
      Top             =   2604
      Width           =   5490
   End
   Begin VB.TextBox txtPan_billto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2340
      TabIndex        =   61
      Top             =   2928
      Width           =   3390
   End
   Begin VB.TextBox txtGST_billto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   60
      Top             =   2928
      Width           =   2070
   End
   Begin VB.TextBox TXTpaN 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   11220
      TabIndex        =   59
      Top             =   936
      Width           =   2505
   End
   Begin VB.TextBox txtSGSTRate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11940
      TabIndex        =   54
      Top             =   6936
      Width           =   315
   End
   Begin VB.TextBox txtCGSTRate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11940
      TabIndex        =   53
      Top             =   6636
      Width           =   315
   End
   Begin VB.TextBox txtSGSTAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12300
      TabIndex        =   52
      Top             =   6960
      Width           =   1035
   End
   Begin VB.TextBox txtCGSTAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12300
      TabIndex        =   51
      Top             =   6660
      Width           =   1035
   End
   Begin VB.TextBox txtVatOn 
      Height          =   315
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   48
      Top             =   8364
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtBindAmt 
      Height          =   315
      Left            =   1380
      TabIndex        =   47
      Top             =   8784
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   12300
      TabIndex        =   41
      Text            =   "0"
      Top             =   8520
      Width           =   1035
   End
   Begin VB.TextBox txtvatamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   46
      Text            =   "0"
      Top             =   8424
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox txtvatRate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2820
      TabIndex        =   45
      Text            =   "0"
      Top             =   8424
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Mvprv 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Move &Previous"
      Height          =   525
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7704
      Width           =   1170
   End
   Begin VB.CommandButton Mvnxt 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Move &Next"
      Height          =   525
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7704
      Width           =   1080
   End
   Begin VB.CommandButton Mvlst 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Move &Last"
      Height          =   525
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7704
      Width           =   1080
   End
   Begin VB.CommandButton mvfrst 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Move &First"
      Height          =   525
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7704
      Width           =   1080
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   33
      Top             =   2604
      Width           =   5490
   End
   Begin VB.TextBox txtLoose 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4860
      TabIndex        =   32
      Top             =   8784
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3360
      TabIndex        =   31
      Top             =   8784
      Visible         =   0   'False
      Width           =   1455
   End
   Begin Crystal.CrystalReport CR 
      Left            =   12540
      Top             =   9444
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   6120
      TabIndex        =   22
      Top             =   8916
      Width           =   7290
      Begin VB.CommandButton cmdPrint_GST 
         Caption         =   "&Print (Tax Inv.)"
         Height          =   600
         Left            =   5340
         TabIndex        =   68
         Top             =   120
         Width           =   1005
      End
      Begin VB.CommandButton cmdExit_12 
         Caption         =   "E&xit"
         Height          =   600
         Left            =   6360
         TabIndex        =   29
         Top             =   135
         Width           =   825
      End
      Begin VB.CommandButton cmdPrint_7 
         Caption         =   "&Print"
         Height          =   600
         Left            =   4440
         TabIndex        =   28
         Top             =   135
         Width           =   885
      End
      Begin VB.CommandButton cmdUndo_5 
         Caption         =   "&Undo"
         Height          =   600
         Left            =   3588
         TabIndex        =   27
         Top             =   135
         Width           =   852
      End
      Begin VB.CommandButton cmdEdit_4 
         Caption         =   "&Edit"
         Height          =   600
         Left            =   2664
         TabIndex        =   26
         Top             =   135
         Width           =   885
      End
      Begin VB.CommandButton cmdDelete_3 
         Caption         =   "&Delete"
         Height          =   600
         Left            =   1785
         TabIndex        =   25
         Top             =   135
         Width           =   885
      End
      Begin VB.CommandButton cmdSave_2 
         Caption         =   "&Save"
         Height          =   600
         Left            =   960
         TabIndex        =   24
         Top             =   135
         Width           =   825
      End
      Begin VB.CommandButton cmdAdd_1 
         Caption         =   "&Add"
         Height          =   600
         Left            =   75
         TabIndex        =   23
         Top             =   135
         Width           =   885
      End
   End
   Begin MSComCtl2.DTPicker Dates 
      Height          =   312
      Left            =   1920
      TabIndex        =   1
      Top             =   576
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      Format          =   150077441
      CurrentDate     =   39500
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8244
      TabIndex        =   21
      Top             =   924
      Width           =   2190
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8220
      TabIndex        =   20
      Top             =   624
      Width           =   5505
   End
   Begin VB.Frame Frame3 
      Caption         =   "Output Weight"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   135
      TabIndex        =   15
      Top             =   9876
      Visible         =   0   'False
      Width           =   465
      Begin VB.TextBox txtRawAndCasting 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3690
         TabIndex        =   16
         Text            =   "0"
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Shape Shape2 
         Height          =   585
         Left            =   75
         Top             =   1365
         Width           =   3150
      End
      Begin VB.Label Label9 
         Caption         =   "Receiving Semi Finish &&  Finish"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   19
         Top             =   1515
         Width           =   3060
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   0
         Top             =   1590
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Receiving  from casting"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   18
         Top             =   885
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Raw Issue Weight"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   570
         Width           =   1635
      End
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   12300
      TabIndex        =   43
      Text            =   "0"
      Top             =   6336
      Width           =   1035
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   216
      TabIndex        =   8
      Top             =   2280
      Width           =   5505
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8220
      TabIndex        =   4
      Top             =   336
      Width           =   5505
   End
   Begin VB.TextBox txtHeating 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   252
      Width           =   1680
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3060
      Left            =   216
      TabIndex        =   40
      Top             =   3516
      Width           =   13308
      _cx             =   23474
      _cy             =   5397
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   16711680
      BackColorFixed  =   13041606
      ForeColorFixed  =   0
      BackColorSel    =   13888387
      ForeColorSel    =   16711680
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   8388608
      GridColorFixed  =   8454016
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
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
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIssue.frx":0017
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      ExplorerBar     =   7
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
   Begin VB.ComboBox cboItem 
      Appearance      =   0  'Flat
      Height          =   2184
      ItemData        =   "frmIssue.frx":00DE
      Left            =   264
      List            =   "frmIssue.frx":00E0
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   79
      Top             =   3516
      Width           =   4785
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DateOfSupp :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   16
      Left            =   7956
      TabIndex        =   87
      Top             =   1620
      Width           =   1128
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TransMode :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   15
      Left            =   4356
      TabIndex        =   86
      Top             =   1656
      Width           =   1128
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "VehicleNo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   14
      Left            =   288
      TabIndex        =   85
      Top             =   1584
      Width           =   1128
   End
   Begin VB.Label Label12 
      Caption         =   "Total IGST @"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   10260
      TabIndex        =   83
      Top             =   7332
      Width           =   1728
   End
   Begin VB.Label lblbal 
      BackStyle       =   0  'Transparent
      Caption         =   "Bending Qty For Bill :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   276
      TabIndex        =   80
      Top             =   7032
      Width           =   3312
   End
   Begin VB.Label Label11 
      Caption         =   "Total :"
      Height          =   240
      Left            =   8412
      TabIndex        =   77
      Top             =   6396
      Width           =   648
   End
   Begin VB.Label Label10 
      Caption         =   "Total Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   10260
      TabIndex        =   75
      Top             =   7956
      Width           =   1968
   End
   Begin VB.Label Label8 
      Caption         =   "Total Amount : GST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   10260
      TabIndex        =   73
      Top             =   7644
      Width           =   1968
   End
   Begin VB.Label Label7 
      Caption         =   "Add/Less Rounding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   10260
      TabIndex        =   71
      Top             =   8256
      Width           =   1968
   End
   Begin VB.Label lblEmail 
      Height          =   252
      Left            =   8280
      TabIndex        =   69
      Top             =   1236
      Width           =   2892
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Reverse Charge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   13
      Left            =   4320
      TabIndex        =   67
      Top             =   936
      Width           =   1608
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Place of Supply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   12
      Left            =   276
      TabIndex        =   66
      Top             =   936
      Width           =   1608
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Shipped To :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   11
      Left            =   5856
      TabIndex        =   63
      Top             =   2016
      Width           =   1128
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   10
      Left            =   7080
      TabIndex        =   58
      Top             =   936
      Width           =   1128
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   9
      Left            =   10440
      TabIndex        =   57
      Top             =   936
      Width           =   768
   End
   Begin VB.Label Label21 
      Caption         =   "Total SGST @"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   10260
      TabIndex        =   56
      Top             =   7020
      Width           =   1728
   End
   Begin VB.Label Label20 
      Caption         =   "Total CGST @"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   10260
      TabIndex        =   55
      Top             =   6720
      Width           =   1608
   End
   Begin VB.Label Label4 
      Caption         =   "VAT On :"
      Height          =   252
      Left            =   240
      TabIndex        =   50
      Top             =   8424
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.Label Label6 
      Caption         =   "Binding Amt."
      Height          =   312
      Left            =   240
      TabIndex        =   49
      Top             =   8844
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Total Amount After Tax :"
      Height          =   192
      Index           =   4
      Left            =   10260
      TabIndex        =   44
      Top             =   8580
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "VAT         %"
      Height          =   192
      Index           =   8
      Left            =   2460
      TabIndex        =   42
      Top             =   8484
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   192
      Left            =   240
      TabIndex        =   35
      Top             =   3264
      Width           =   13212
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Billed To :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   276
      TabIndex        =   34
      Top             =   2016
      Width           =   1128
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search Bill"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1920
      TabIndex        =   30
      Top             =   24
      Width           =   2508
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search Firm"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   8220
      TabIndex        =   14
      Top             =   84
      Width           =   2508
   End
   Begin VB.Label Label1 
      Caption         =   "Taxable Value :"
      Height          =   252
      Index           =   6
      Left            =   10260
      TabIndex        =   13
      Top             =   6336
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   276
      TabIndex        =   12
      Top             =   576
      Width           =   1128
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Firm Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   7080
      TabIndex        =   11
      Top             =   384
      Width           =   1128
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   276
      TabIndex        =   10
      Top             =   264
      Width           =   1128
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemgp As String
Dim rates As Double
Dim i As Integer
Dim Status As String
Dim Item_Name As String
Dim unit As String
Dim qty As Integer
Dim iitem1 As String
Dim StockFlag As String
Private Sub cmdMain_Click()
   Unload Me
End Sub
Sub viewData(item_ As String, gp_ As String)

Dim str1 As String
str1 = ""

If rs.State = 1 Then rs.Close
rs.Open "select  ItemName,Itemgp from ItemMaster  where Itemgp='" & gp_ & "' and ItemName='" & item_ & "'", con
While rs.EOF = False
     con.Execute "update ITC set bgp='" & rs!itemgp & "' where BookName= '" & rs!itemname & "'"
     con.Execute "update BookReceive set bgp= '" & rs!itemgp & "' where BookName= '" & rs!itemname & "'"
rs.MoveNext
Wend




If (item_ <> "") Then
   str1 = " bookname='" & item_ & "'"
End If

If (gp_ <> "") Then
   str1 = str1 & " and bgp='" & gp_ & "'"
End If


Dim qty
qty = 0
  
If rs.State = 1 Then rs.Close
rs.Open "select EntDate,BOOKName,qty,status from PendingBill  where " & str1 & " order by BOOKName,entdate", con
For k1 = 1 To rs.RecordCount
 
 If rs!Status = "R" Then
    qty = qty + rs!qty
 ElseIf rs!Status = "D" Then
    qty = qty - rs!qty
 Else
    qty = qty - rs!qty
 End If

 
 rs.MoveNext

Next


lblbal.Caption = "Bending Qty For Bill : " & qty

End Sub

Sub cellposi()
 'VsFrame.Top = vs.Top + ((vs.CellTop)) - 1850
 'VsFrame.Left = (vs.Left) + 4500
End Sub
Sub Total()
Dim sum_total As Double
sum_total = 0

For j = 1 To vs.Rows - 1
    If vs.TextMatrix(j, 0) <> "" Then
    sum1 = 0
    If vs.TextMatrix(j, 1) = "Book" Then
        sum1 = ((Val(vs.TextMatrix(j, 5)) / 1000) * Val(vs.TextMatrix(j, 4)))
        vs.TextMatrix(j, 7) = (sum1 * Val(vs.TextMatrix(j, 6)))
    Else
        vs.TextMatrix(j, 7) = (Val(vs.TextMatrix(j, 5)) * Val(vs.TextMatrix(j, 6)))
    End If
    End If
Next


'txtTotal1.Text = 0
txtTotal2.text = 0

For j = 1 To vs.Rows - 1
If vs.TextMatrix(j, 0) <> "" Then
   'txtTotal1.Text = (Val(txtTotal1.Text) + Val(vs.TextMatrix(j, 7)))
   txtTotal2.text = (Val(txtTotal2.text) + Val(vs.TextMatrix(j, 7)))
End If
Next

'txtTotal1.Text = Format(Round(txtTotal1.Text, 2), "0.00")
txtTotal2.text = Format(Round(txtTotal2.text, 2), "0.00")

End Sub
Sub calc()

Dim sum_total As Double
sum_total = 0

txtTotal1.text = 0
txtLoose.text = 0

For j = 1 To vs.Rows - 1
If vs.TextMatrix(j, 7) <> "" Then
   txtTotal1.text = (Val(txtTotal1.text) + Val(vs.TextMatrix(j, 7)))
End If
Next



txtTotal1.text = Format(Round(txtTotal1.text, 2), "0.00")
sum_total = txtTotal1.text



'===============================================
Dim a1, a2 As Double
Dim sum_ As Double
sum_ = 0
'''cgst

If Check1_igst.Value = 0 Then

    If Val(txtCGSTRate) > 0 Then
        a1 = Val(sum_total)
        a2 = (a1 * CDbl(txtCGSTRate)) / 100
        
        If a2 > 0 Then
          'a1 = (a2 + 0.0001)
          txtCGSTAmt.text = Round(a2, 2)
        End If
    End If
    
    'sgst
    If Val(txtSGSTRate) > 0 Then
        a1 = Val(sum_total)
        a2 = (a1 * CDbl(txtSGSTRate)) / 100
        
        If a2 > 0 Then
          'a1 = (a2 + 0.0001)
          txtSGSTAmt.text = Round(a2, 2)
        End If
    End If
    
    txtIGSTAmt = 0
    txtIGSTRate = 0

Else

    If Val(txtIGSTRate) > 0 Then
        a1 = Val(sum_total)
        a2 = (a1 * CDbl(txtIGSTRate)) / 100
        
        If a2 > 0 Then
          txtIGSTAmt.text = Round(a2, 2)
        End If
    End If
    
    txtSGSTAmt = 0
    txtCGSTAmt = 0
    txtSGSTRate = 0
    txtCGSTRate = 0
    

End If



txtTGST.text = (Val(txtCGSTAmt) + Val(txtSGSTAmt) + Val(txtIGSTAmt))


txtTotalValue.text = (sum_total + Val(txtTGST.text))

t1 = Int(CDbl(txtTotalValue.text))
a1 = Int(CDbl(txtTotalValue.text))
a2 = CDbl(txtTotalValue.text)

a1 = Int(a1)
a1 = Round(a2 - a1, 2)

If a1 > 0 Then
   
If a1 <= 0.49 Then
   a2 = -1 * a1
Else
   a2 = 1 - Abs(a1)
End If

Else

a2 = 0
End If


txtAddLess.text = a2




txtNet.text = Round(t1 + Val(txtAddLess.text), 0)


End Sub
Sub calculate()

Dim sum_total As Double
sum_total = 0

txtTotal1.text = 0
txtLoose.text = 0

Set rs1 = New ADODB.Recordset
rs1.Open "select  sum(amt) from InvoicebGST where (invoiceno=" & txtHeating & " and firm='" & txtParty.text & "')", con
If rs1.EOF = False Then

 If Not IsNull(rs1(0)) Then
     txtTotal1.text = rs1(0)
 End If

End If


txtTotal1.text = Format(Round(txtTotal1.text, 2), "0.00")
sum_total = txtTotal1.text



'===============================================
Dim a1, a2 As Double
Dim sum_ As Double
sum_ = 0
'''cgst

If Check1_igst.Value = 0 Then

    If Val(txtCGSTRate) > 0 Then
        a1 = Val(sum_total)
        a2 = (a1 * CDbl(txtCGSTRate)) / 100
        
        If a2 > 0 Then
          'a1 = (a2 + 0.0001)
          txtCGSTAmt.text = Round(a2, 2)
        End If
    End If
    
    'sgst
    If Val(txtSGSTRate) > 0 Then
        a1 = Val(sum_total)
        a2 = (a1 * CDbl(txtSGSTRate)) / 100
        
        If a2 > 0 Then
          'a1 = (a2 + 0.0001)
          txtSGSTAmt.text = Round(a2, 2)
        End If
    End If
    
    txtIGSTAmt = 0
    txtIGSTRate = 0

Else

    If Val(txtIGSTRate) > 0 Then
        a1 = Val(sum_total)
        a2 = (a1 * CDbl(txtIGSTRate)) / 100
        
        If a2 > 0 Then
          txtIGSTAmt.text = Round(a2, 2)
        End If
    End If
    
    txtSGSTAmt = 0
    txtCGSTAmt = 0
    txtSGSTRate = 0
    txtCGSTRate = 0
    

End If


txtTGST.text = (IIf(txtCGSTAmt = "", 0, Val(txtCGSTAmt)) + IIf(txtSGSTAmt = "", 0, Val(txtSGSTAmt)) + IIf(txtIGSTAmt = "", 0, Val(txtIGSTAmt)))
txtTotalValue.text = (sum_total + Val(txtTGST.text))

t1 = Int(CDbl(txtTotalValue.text))
a1 = Val(txtTotalValue.text)
a2 = Val(txtTotalValue.text)


a1 = Int(a1)
a1 = Round(a2 - a1, 2)

If a1 > 0 Then
   
If a1 <= 0.49 Then
   a2 = -1 * a1
Else
   a2 = 1 - Abs(a1)
End If

Else

a2 = 0
End If






txtAddLess.text = Round(a2, 2)
txtNet.text = Round(Val(txtTotalValue.text) + Val(txtAddLess.text), 0)


End Sub


Sub cellposiVs()
 Vs1Frame.Width = 2500
 Vs1Frame.Top = vs1.Top + ((vs1.CellTop))
 Vs1Frame.Left = (vs1.Left) + 550
End Sub
Sub AddItemInGrid1()
'
'    Dim rs_4 As New ADODB.Recordset
'
'    rs_4.Open "select * from ItemMaster order by BKDESC", con, adOpenDynamic, adLockOptimistic
'
'    Set cboitemvs1.RowSource = rs_4
'    cboitemvs1.ListField = "BKDESC"
'    cboitemvs1.BoundColumn = "BKCODE"
'    cboitemvs1.ReFill
    
End Sub
Sub AddItemInGrid3()
    'Adodc1.ConnectionString = "filedsn=Saru"
    'Adodc1.CommandType = adCmdText
'
'    Dim rs_3 As New ADODB.Recordset
'
'    'rs_3.Open "select * from ItemMaster where (ItemGp='Finish Item' or ItemGp='Scrap' or ItemGp='Losses' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') or  Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
'    rs_3.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
'
'    'Adodc1.Refresh
'    Set cboItemVs3.RowSource = rs_3
'    cboItemVs3.ListField = "ItemName"
'    cboItemVs3.BoundColumn = "ItemName"
'    cboItemVs3.ReFill
    
End Sub
Sub AddItemInGrid2()
'    'Adodc1.ConnectionString = "filedsn=Saru"
'    'Adodc1.CommandType = adCmdText
'    Dim rs_2 As New ADODB.Recordset
'
'    'rs_2.Open "select * from ItemMaster where (ItemGp='Semi Finish (R/D)' or ItemGp= 'Semi Finish (Store)' or Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
'    rs_2.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
'
'    Set cboItemVs2.RowSource = rs_2
'    cboItemVs2.ListField = "ItemName"
'    cboItemVs2.BoundColumn = "ItemName"
'    cboItemVs2.ReFill
    
End Sub
Sub AddItemInGrid()
    
    Dim rs_1 As New ADODB.Recordset
    
    'rs_1.Open "select * from ItemMaster where ItemGp='Raw Item' or ItemGp='Scrap' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') order by ItemName", con, adOpenDynamic, adLockOptimistic
    
    If vs.TextMatrix(vs.RowSel, 2) <> "" Then
    rs_1.Open "select * from ItemMaster where ItemGp='" & vs.TextMatrix(vs.RowSel, 2) & "' order by ItemName", con, adOpenDynamic, adLockOptimistic
    Else
    rs_1.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
    End If
    
    cboItem.clear
    While rs_1.EOF = False
    cboItem.AddItem rs_1!itemname
    rs_1.MoveNext
    Wend
    
    
    'Set cboItem.RowSource = rs_1
    'cboItem.ListField = "ItemName"
    'cboItem.BoundColumn = "ItemName"
    'cboItem.ReFill
    
End Sub
Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQty.SetFocus
End Sub


Private Sub cbogodown_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then txtRemarks.SetFocus
End Sub

Private Sub cbogodown_LostFocus()
If cbogodown = "" Then
   MsgBox "Select Godown Name ..", vbCritical
   cbogodown.SetFocus
   Exit Sub
End If
End Sub

Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then

        vs.TextMatrix(vs.RowSel, 3) = cboItem.text

        If rs.State = 1 Then rs.Close
        rs.Open "select BookFarmNo,BookRate,TitleRate from [ItemMaster] where ItemName='" & Trim(vs.TextMatrix(vs.RowSel, 3)) & "'", con

        If rs.EOF = False Then

            '================ UI handling =================
            cboItem.Visible = False
            vs.SetFocus
            sendkeys "{RIGHT}"
            '================================================

            '================ FIX START =================
            If vs.TextMatrix(vs.RowSel, 1) = "Book" Then

                Dim farmVal As String
                farmVal = rs.Fields(0).Value & ""

                If IsNumeric(farmVal) Then
                    vs.TextMatrix(vs.RowSel, 4) = Val(farmVal)
                Else
                    vs.TextMatrix(vs.RowSel, 4) = 0   'fallback to avoid crash
                End If

                vs.TextMatrix(vs.RowSel, 5) = rs!BookRate & ""

            ElseIf (vs.TextMatrix(vs.RowSel, 1) = "Title" Or vs.TextMatrix(vs.RowSel, 1) = "Stich") Then

                vs.TextMatrix(vs.RowSel, 5) = rs!TitleRate & ""

            Else

                vs.TextMatrix(vs.RowSel, 5) = rs!TitleRate & ""

            End If
            '================ FIX END =================

        Else
            Exit Sub
        End If

        vs.SetFocus

    ElseIf KeyCode = 27 Then

        cboItem.Visible = False

    End If

End Sub
Sub saveInMaster()
         On Error Resume Next
      
         If rs.State = 1 Then rs.Close
         rs.Open "select * from ItemMaster where ItemName='" & iitem1 & "'", con, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.AddNew
            rs.Fields("ItemGp").Value = frmAddMaster.cbogp.text
            rs.Fields("ItemName").Value = iitem1
            rs.Fields("Unit").Value = "Kg"
            rs.Update
         Else
            MsgBox "This Item Already Exist !!", vbCritical
            Exit Sub
         End If
         frmAddMaster.Visible = False
  
End Sub
Private Sub cboitemvs1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        cellposiVs
        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.text
        
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where ItemName='" & cboitemvs1.text & "'", con
        If rs.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                frmAddMaster.Show 1
                iitem1 = cboitemvs1.text
                saveInMaster
                cboitemvs1.text = ""
                Vs1Frame.Visible = False
                vs1.SetFocus
             End If
        End If
        vs1.SetFocus
     ElseIf KeyCode = 27 Then
        Vs1Frame.Visible = False
     End If
End Sub


Private Sub cboItemVs2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
        
        'cellposiVs2
        vs3.TextMatrix(vs3.RowSel, 0) = cboItemVs2.text
        Set rs = New ADODB.Recordset
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where ItemName='" & cboItemVs2.text & "'", con
        If rs.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                frmAddMaster.Show 1
                iitem1 = cboItemVs2.text
                saveInMaster
                
                cboItemVs2.text = ""
             End If
        End If
        vs3.SetFocus
        
ElseIf KeyCode = 27 Then
         FrameVs2.Visible = False
End If
End Sub

Private Sub cboItemVs3_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        'cellposiVs3
        
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.text
        Set rs = New ADODB.Recordset
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where ItemName='" & cboItemVs3.text & "'", con
        If rs.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                Vs3Frame.Visible = False
                iitem1 = cboItemVs3.text
                frmAddMaster.Show 1
                saveInMaster
                cboItemVs3.text = ""
                vs2.SetFocus
             End If
        End If
        Vs3Frame.Visible = False
        'cboItemVs3.Visible = False
        vs2.SetFocus
     ElseIf KeyCode = 27 Then
        Vs3Frame.Visible = False
     End If

End Sub

Private Sub cmdadd_Click()
 If rs.State = 1 Then rs.Close
 rs.Open "select HeatingNo from IssueMaster where HeatingDate >=datevalue('" & fromDate.Value & "') and HeatingDate <=datevalue('" & todate.Value & "') order by HeatingNo", con
 ListHeatingNo.clear
 If rs.EOF = False Then
    While rs.EOF = False
       ListHeatingNo.AddItem rs(0)
       rs.MoveNext
    Wend
 End If
End Sub
Private Sub cmdDelete_Click()
   
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
      
      DeleteRecord txtHeating.text, "HeatingNo", "IssueMaster"
      DeleteRecord txtHeating.text, "HeatingNo", "IssueRawMetrial"
      Call cmdRef_Click
      
   End If
End Sub
Sub DeleteStock()
    
'''Dim rr As New ADODB.Recordset
'''Dim rs_u As New ADODB.Recordset
'''Dim openning As Double
'''
'''
'''
'''
'''
''''================ Issue For Casting
'''
'''
''' If StockFlag = "1" Then
'''
'''    If rs_u.State = 1 Then rs_u.Close
'''    rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''    If rs_u.EOF = False Then
'''        rs_u!qty = rs_u!qty + qty
'''        rs_u.Update
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Receive For Casting
'''
''' If StockFlag = "2" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''            rs_u!qty = rs_u!qty - qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Receive For Finish
'''
''' If StockFlag = "3" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''             rs_u!qty = rs_u!qty - qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Issue For Finish
'''
''' If StockFlag = "4" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''             rs_u!qty = rs_u!qty + qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
''' '====================================
'''
'''
'''
'''
'''
    
   
   
    
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFatch_Click()
AddSemifinish
'Total4

End Sub

Private Sub cmdFind_Click()
 Frame1.Visible = True
 fromDate.SetFocus
End Sub

Private Sub cmdModify_Click()
   If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
      
        
      DeleteRecord txtHeating.text, "HeatingNo", "IssueMaster"
      DeleteRecord txtHeating.text, "HeatingNo", "IssueRawMetrial"
      
      
      
      'SaveData
      
      UpdateIssue
      
      Call cmdRef_Click
   End If
End Sub
Sub UpdateIssue()

Dim rss As New ADODB.Recordset
Dim Search As New ADODB.Recordset
    
If Search.State = 1 Then Search.Close
Search.Open "select ItemName,qty from Invoice where HeatNo='" & txtHeating.text & "'", con
If Search.EOF = False Then
While Search.EOF = False

    If rss.State = 1 Then rss.Close
    rss.Open "select * from IssueRawMetrial where HeatingNo=" & txtHeating.text & " and ItemName='" & Search.Fields(0).Value & "'", con, adOpenDynamic, adLockOptimistic
    If rss.EOF = False Then
       rss.Fields("Issue").Value = (CDbl(rss.Fields("Issue").Value) + CDbl(Search.Fields("qty").Value))
       rss.Update
    End If
    
    Search.MoveNext
    
Wend
  
End If
  
End Sub

Private Sub cmdRef_Click()
      txtHeating.text = ""
      txtParty.text = ""
      
      txtRemarks.text = ""
      
      
      txtTotal1.text = 0
      txtTotal2.text = 0
      txtTotal3.text = 0
      txtTotal4.text = 0
      
      txtSize.text = ""
      txtGrade.text = ""
      txtRawAndCasting.text = 0
      
      vs.clear
      vs1.clear
      vs2.clear
      vs3.clear
      
      SetWidth
      txtHeating.SetFocus
      cmdDelete.Enabled = False
      cmdModify.Enabled = False
      cmdSave.Enabled = True
      
      Record = ""
      
End Sub


Private Sub Command4_Click()
   Unload Me
End Sub
Private Sub cmdSave_Click()
    
    
    
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from IssueMaster where HeatingNo=" & txtHeating.text & "", con
    If rs.EOF = False Then
       MsgBox "Heating No. Already Exist !!", vbInformation
       Exit Sub
    End If
    
    If txtHeating.text = "" Then
       MsgBox "Please Enter Heating No !!", vbCritical
       txtHeating.SetFocus
       Exit Sub
    End If
    
    
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from IssueMaster where HeatingNo=" & txtHeating.text & "", con
    If rs.EOF = True Then
       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
        '  SaveData
       End If
    Else
          MsgBox "Dublicate Heating No !!", vbCritical
    End If
End Sub
Sub ItemGpSearch(Str As String)
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ItemGp,Rate from ItemMaster where ItemName='" & Str & "'", con
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).Value
       rates = rs1.Fields(1).Value
    End If
    
End Sub
Sub UpdateStock()
    Dim rr As New ADODB.Recordset
    Dim rs_u As New ADODB.Recordset
    Dim openning As Double
    
 
    
    
 '================ Issue For Casting
 
 
 If StockFlag = "1" Then
    
    If rs_u.State = 1 Then rs_u.Close
    rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
    If rs_u.EOF = True Then
        rs_u.AddNew
        rs_u!itemname = Item_Name
        ItemGpSearch Item_Name
        rs_u!itemgp = itemgp
        rs_u!unit = unit
        rs_u!Rate = rates
        rs_u!qty = (-1 * qty)
        rs_u.Update
     Else
        rs_u!qty = rs_u!qty - qty
        rs_u.Update
    End If
    
 End If
 
 '====================================
 
 
 '================ Receive For Casting
 
 If StockFlag = "2" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!Rate = rates
            rs_u!qty = qty
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty + qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
 
 
 '================ Receive For Finish
 
 If StockFlag = "3" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!Rate = rates
            rs_u!qty = qty
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty + qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
 
 
  '================ Issue For Finish
 
 If StockFlag = "4" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!Rate = rates
            rs_u!qty = (-1 * qty)
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty - qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
    
    
    
    
End Sub
 
Private Sub Check1_igst_Click()
showGST
calc
End Sub

Private Sub cmdAdd_1_Click()
   
   
   
   txtVehiNo.text = ""
txtTransMode.text = ""
txtDateOfSupp.text = ""

   
    txtHeating.text = ""
    Dates.Value = Date
    'txtParty.Text = ""
    txtRemarks.text = ""
    'Text1.Text = ""
    'Text2.Text = ""
    txtTotal.text = ""
    txtLoose.text = ""
    txtAdd = ""
    
    txtTotal1 = "0"
    txtTotal2.text = 0
    txtvatamt = 0
    txtNet = 0
    
    lblbal.Caption = ""
    
    txtBindAmt.text = 0
    txtCGSTAmt.text = 0
    txtSGSTAmt.text = 0
    
    txtTGST = 0
    txtTotalValue = 0
    txtAddLess = 0
    
    txtPan_billto.text = ""
    txtGST_billto.text = ""
    txtParty_Shippto.text = ""
    txtAdd_Shippto.text = ""
    txtPAN_Shippto.text = ""
    txtGST_Shippto.text = ""
    
    cboRCharge.ListIndex = 1
    txtPlaceofsupp.text = "MEERUT"
    
   
   'RefData Me
   vs.clear
   SetWidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   'cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   txtHeating.SetFocus
   showGST
   
   If txtParty = "" Then
      MsgBox "select firm name ....", vbInformation
      Exit Sub
   End If
   
   
   Set rs = New ADODB.Recordset
   rs.Open "select max(invoiceno) from invoicea where (SUBLEDGER='" & Trim(txtParty) & "')", con
   If Not IsNull(rs(0)) Then
     txtHeating.text = rs(0) + 1
   Else
     txtHeating.text = 1
   End If
   
   
   
'cbogodown.ListIndex = 0
   
Set rs = New ADODB.Recordset
rs.Open "select * from invoicea", con, adOpenForwardOnly, adLockReadOnly
If rs.EOF = False Then
   rs.MoveLast
   Dates.Value = rs!invoicedate
End If


 If rs.State = 1 Then rs.Close
 rs.Open "select vatper from setup1", con
 If rs.EOF = False Then
    txtvatRate.text = rs(0)
 End If


End Sub
Sub clear()
    txtHeating.text = ""
    Dates.Value = Date
    'txtParty.Text = ""
    txtRemarks.text = ""
    'Text1.Text = ""
    'Text2.Text = ""
    txtTotal.text = ""
    txtLoose.text = ""
    txtAdd = ""
    
    txtTotal1 = ""
   
   'RefData Me
   vs.clear
   SetWidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   'cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   
   
   If txtParty = "" Then
      MsgBox "select firm name ....", vbInformation
      Exit Sub
   End If
   
   
   Set rs = New ADODB.Recordset
   rs.Open "select max(invoiceno) from invoicea where (SUBLEDGER='" & Trim(txtParty) & "')", con
   If Not IsNull(rs(0)) Then
     'txtHeating.Text = MaxSNo("invoicea", "INVOICENO")
     txtHeating.text = rs(0) + 1
   Else
     txtHeating.text = 1
   End If
   
   
   
   'cbogodown.ListIndex = 0
   
Set rs = New ADODB.Recordset
rs.Open "select * from invoicea", con, adOpenForwardOnly, adLockReadOnly
If rs.EOF = False Then
   rs.MoveLast
   Dates.Value = rs!invoicedate
End If
End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
con.Execute "delete from invoiceb where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "')"
con.Execute "delete from invoicea where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "')"
Call cmdAdd_1_Click
End If
End Sub

Private Sub cmdEdit_4_Click()
   cmdDelete_3.Enabled = True
   cmdEdit_4.Enabled = False
   cmdPrint_7.Enabled = True
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = False
   cmdExit_12.Enabled = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_7_Click()

On Error GoTo err1

CR.Reset
CR.ReportFileName = App.Path & "/CHALLAN.rpt"

' ? ADD THIS (check full path)
' MsgBox App.Path & "\" & frmPassword.cboyrs.text & "\Data.mdb"
' CR.Connect = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\" & frmPassword.cboyrs.text & "\Data.mdb;"

CR.ReplaceSelectionFormula "{invoiceA.invoiceno}=" & txtHeating.text & " and {invoiceA.subledger}='" & txtParty.text & "'"



If (txtvatamt) > 0 Then
   wrd = toword(txtTotal1.text)
   CR.Formulas(0) = "wordsamt='" & wrd & "'"
Else
   wrd = toword(txtTotal1)
   CR.Formulas(0) = "wordsamt='" & wrd & "'"
End If

If rs.State = 1 Then rs.Close
rs.Open "select Phone1,Phone2 from setup1", con
If rs.EOF = False Then
   CR.Formulas(1) = "phone='" & rs(0) & ", " & rs(1) & "'"
End If

If rs.State = 1 Then rs.Close
rs.Open "select Phone,tin from firm where name='" & txtParty & "'", con
If rs.EOF = False Then
   CR.Formulas(2) = "tin='" & rs(1) & "'"
End If



Dim s1_ As String
s1_ = billFormat(txtHeating.text)
If s1_ <> "" Then

CR.Formulas(3) = "bill_='" & s1_ & "'"
End If


CR.WindowShowPrintSetupBtn = True
CR.WindowState = crptMaximized

CR.Action = 1

Exit Sub

err1:
MsgBox "ERROR: " & Err.Description
End Sub

Private Sub cmdPrint_GST_Click()
On Error GoTo err1
FNAME = txtParty.text

CR.Reset
CR.ReportFileName = App.Path & "/CHALLANGST.rpt"


' ? ADD THIS (check full path)
' MsgBox App.Path & "\" & frmPassword.cboyrs.text & "\Data.mdb"
' CR.Connect = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\" & frmPassword.cboyrs.text & "\Data.mdb;"
CR.ReplaceSelectionFormula "{invoiceA.invoiceno}=" & txtHeating.text & " and {InvoicebGST.firm}='" & txtParty.text & "'"
'If (txtvatamt) > 0 Then
wrd = toword(txtNet.text)
CR.Formulas(0) = "wordsamt='" & wrd & "'"
'Else
'wrd = toword(txtTotal1)
'cr.Formulas(0) = "wordsamt='" & wrd & "'"
'End If
Set rs = New ADODB.Recordset
rs.Open "select Phone,tin from firm where name='" & txtParty & "'", con
If rs.EOF = False Then
   'cr.Formulas(1) = "phone='" & rs(0) & "'"
   CR.Formulas(2) = "tin='" & rs(1) & "'"
End If

If rs.State = 1 Then rs.Close
rs.Open "select phone2 from setup1", con
If rs.EOF = False Then
   CR.Formulas(1) = "phone='" & rs(0) & "'"
End If

If rs.State = 1 Then rs.Close
rs.Open "select bank,account,ifsc from firm where name='" & txtParty.text & "'", con
If rs.EOF = False Then
   CR.Formulas(3) = "bank='" & rs(0) & "'"
   CR.Formulas(4) = "acc='" & rs(1) & "'"
   CR.Formulas(5) = "ifsc='" & rs(2) & "'"
End If


Dim s1_ As String
s1_ = billFormat(txtHeating.text)
If s1_ <> "" Then
CR.Formulas(6) = "bill_='" & s1_ & "'"
End If



CR.WindowShowPrintSetupBtn = True
CR.WindowState = crptMaximized
CR.Action = 1


Exit Sub

err1:
MsgBox "ERROR: " & Err.Description
End Sub
Private Sub cmdSave_2_Click()

SaveData

'===========================================================
calculate

If rs.State = 1 Then rs.Close
rs.Open "select * from invoicea where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "')", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then

    If Check1_ItemWiseTot.Value = 1 Then
    rs!AgentName = "y"
    Else
    rs!AgentName = "n"
    End If


    rs.Fields("VehicleNo").Value = txtVehiNo.text
    rs.Fields("TransMode").Value = txtTransMode.text
    rs.Fields("DateOfSupp").Value = txtDateOfSupp.text
     
    rs.Fields("INVOICENO").Value = txtHeating.text
    rs.Fields("INVOICEDATE").Value = Dates.Value
    rs.Fields("SUBLEDGER").Value = txtParty.text
    rs.Fields("GENLEDGER").Value = "Sundry Debtors"
    rs.Fields("Remarks").Value = txtRemarks.text
    rs.Fields("add1").Value = Text1.text
    rs.Fields("add2").Value = Text2.text
    rs.Fields("godown").Value = txtAdd.text
    rs.Fields("PAN").Value = TXTpaN.text
    rs.Fields("email").Value = lblEmail.Caption
    rs.Fields("GAMOUNT").Value = Val(txtTotal1.text)
    rs.Fields("netamount").Value = Val(txtNet.text)
    rs.Fields("Vatper").Value = Val(txtvatRate.text)
    rs.Fields("Vatamt").Value = Val(txtvatamt.text)
    rs.Fields("Vaton").Value = Val(txtVatOn.text)
    rs.Fields("bindingAmt").Value = Val(txtBindAmt.text)
    rs.Fields("CGSTAmt").Value = Val(txtCGSTAmt.text)
    rs.Fields("CGSTrate").Value = Val(txtCGSTRate.text)
    
    rs.Fields("SGSTAmt").Value = Val(txtSGSTAmt.text)
    rs.Fields("SGSTRate").Value = Val(txtSGSTRate.text)
    
    rs.Fields("IGSTAmt").Value = Val(txtIGSTAmt.text)
    rs.Fields("IGSTRate").Value = Val(txtIGSTRate.text)


   rs.Fields("TotalValue").Value = Val(txtTotalValue.text)
   rs.Fields("TotalGST").Value = Val(txtTGST.text)
   rs.Fields("addless").Value = Val(txtAddLess.text)
    
    rs.Fields("PAN_Billto").Value = txtPan_billto.text
    rs.Fields("GSTIN_Billto").Value = txtGST_billto.text
    rs.Fields("Party_shippto").Value = txtParty_Shippto.text
    rs.Fields("Add_shippto").Value = txtAdd_Shippto.text
    rs.Fields("PAN_shippto").Value = txtPAN_Shippto.text
    rs.Fields("GSTIN_shippto").Value = txtGST_Shippto.text
    rs.Fields("placeofsupp").Value = txtPlaceofsupp.text
    rs.Fields("rcharge").Value = cboRCharge.text
    rs!baa = Val(txtTotal2.text)
    rs.Update
End If





End Sub
Sub SaveData()
On Error GoTo aa1

Dim date1 As Date



If rs.State = 1 Then rs.Close
rs.Open "select invoicedate from invoicea where (SUBLEDGER='" & Trim(txtParty) & "') order by invoiceno desc", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
   date1 = rs(0)
   con.Execute "delete from invoiceb nvoiceb where INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "'"
End If


If rs.State = 1 Then rs.Close
rs.Open "select max(invoiceno) from invoicea where (SUBLEDGER='" & Trim(txtParty) & "')", con, adOpenDynamic, adLockOptimistic
If Not IsNull(0) Then
 
   If Val(txtHeating) > rs(0) Then
   
      If (DateValue(Dates.Value) < date1) Then
         MsgBox "Not Valid Date.....", vbCritical
         Exit Sub
      End If
   
   End If
 

End If




If txtParty.text = "" Then
MsgBox "Please Enter Firm Name !!", vbInformation
txtParty.SetFocus
Exit Sub
End If


If txtRemarks.text = "" Then
MsgBox "Please Enter Party Name !!", vbInformation
txtRemarks.SetFocus
Exit Sub
End If





If MsgBox("Want to Save ?", vbYesNo + vbQuestion) = vbYes Then

If rs.State = 1 Then rs.Close
rs.Open "select * from invoicea where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "')", con, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then
    rs.AddNew
End If

    
    rs.Fields("VehicleNo").Value = txtVehiNo.text
    rs.Fields("TransMode").Value = txtTransMode.text
    rs.Fields("DateOfSupp").Value = txtDateOfSupp.text


    rs.Fields("INVOICENO").Value = txtHeating.text
    rs.Fields("INVOICEDATE").Value = Dates.Value
    rs.Fields("SUBLEDGER").Value = txtParty.text
    rs.Fields("GENLEDGER").Value = "Sundry Debtors"
    rs.Fields("Remarks").Value = txtRemarks.text
    rs.Fields("add1").Value = Text1.text
    rs.Fields("add2").Value = Text2.text
    rs.Fields("godown").Value = txtAdd.text
    rs.Fields("PAN").Value = TXTpaN.text
    
    rs.Fields("email").Value = lblEmail.Caption

    rs.Fields("GAMOUNT").Value = Val(txtTotal1.text)
    rs.Fields("netamount").Value = Val(txtNet.text)
    
    rs.Fields("Vatper").Value = Val(txtvatRate.text)
    rs.Fields("Vatamt").Value = Val(txtvatamt.text)
    rs.Fields("Vaton").Value = Val(txtVatOn.text)
    rs.Fields("bindingAmt").Value = Val(txtBindAmt.text)
    
    
    rs.Fields("CGSTAmt").Value = Val(txtCGSTAmt.text)
    rs.Fields("CGSTrate").Value = Val(txtCGSTRate.text)
    
    rs.Fields("SGSTAmt").Value = Val(txtSGSTAmt.text)
    rs.Fields("SGSTRate").Value = Val(txtSGSTRate.text)

   rs.Fields("TotalValue").Value = Val(txtTotalValue.text)
   rs.Fields("TotalGST").Value = Val(txtTGST.text)
   rs.Fields("addless").Value = Val(txtAddLess.text)
    
    rs.Fields("PAN_Billto").Value = txtPan_billto.text
    rs.Fields("GSTIN_Billto").Value = txtGST_billto.text
    rs.Fields("Party_shippto").Value = txtParty_Shippto.text
    rs.Fields("Add_shippto").Value = txtAdd_Shippto.text
    rs.Fields("PAN_shippto").Value = txtPAN_Shippto.text
    rs.Fields("GSTIN_shippto").Value = txtGST_Shippto.text
    
    rs.Fields("placeofsupp").Value = txtPlaceofsupp.text
    rs.Fields("rcharge").Value = cboRCharge.text
    
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select  [state],TaxHead from Customer where Name='" & txtRemarks.text & "'", con
    If rs1.EOF = False Then
       rs.Fields("Stcode_Billto").Value = rs1!TaxHead & ""
       rs.Fields("State_Billto").Value = rs1(0) & ""
    End If

    If rs1.State = 1 Then rs1.Close
    rs1.Open "select  [state],TaxHead from Customer where Name='" & txtParty_Shippto.text & "'", con
    If rs1.EOF = False Then
       rs.Fields("Stcode_shippto").Value = rs1!TaxHead & ""
       rs.Fields("State_shippto").Value = rs1(0) & ""
    End If


    rs!baa = Val(txtTotal2.text)
    rs.Update


cmdSave_2.Enabled = False
cmdPrint_7.SetFocus



Dim rs2 As New ADODB.Recordset
Dim sum1_ As Double
Dim newAmt_ As Double
Dim newrate_ As Double

Dim Qty_, Amt_



If rs.State = 1 Then rs.Close
rs.Open "select * from invoiceb where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "')", con, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then


For i = 1 To vs.Rows - 1

If vs.TextMatrix(i, 2) <> "" Then

rs.AddNew
rs.Fields("INVOICENO").Value = txtHeating.text
rs.Fields("INVOICEDATE").Value = Dates.Value
rs.Fields("SUBLEDGER").Value = txtParty.text
rs.Fields("GENLEDGER").Value = "Sundry Debtors"

If vs.TextMatrix(i, 0) <> "" Then
rs.Fields("sno1").Value = vs.TextMatrix(i, 0)
Else
rs.Fields("sno1").Value = i
End If



rs.Fields("type").Value = vs.TextMatrix(i, 1)
rs.Fields("bookgp").Value = vs.TextMatrix(i, 2)
rs.Fields("BOOKCODE").Value = vs.TextMatrix(i, 3)

rs.Fields("TBook").Value = vs.TextMatrix(i, 4)


rs.Fields("rate").Value = vs.TextMatrix(i, 5)
rs.Fields("qty").Value = vs.TextMatrix(i, 6)
rs.Fields("amt").Value = vs.TextMatrix(i, 7)

If rs2.State = 1 Then rs2.Close
rs2.Open "select * from itemmaster where ItemName='" & vs.TextMatrix(i, 3) & "'", con, adOpenDynamic, adLockOptimistic
If rs2.EOF = False Then
 If vs.TextMatrix(i, 1) = "Book" Then
    rs!remarks = rs2!remarks & ""
 End If
 
 rs!hsncode = rs2!hsncode & ""
 
End If


rs.Update

End If

DoEvents
DoEvents
DoEvents


Next

Else
con.Execute "delete from invoiceb where (INVOICENO=" & txtHeating.text & "  and SUBLEDGER='" & Trim(txtParty) & "')"

For i = 1 To vs.Rows - 1

If vs.TextMatrix(i, 2) <> "" Then

rs.AddNew
rs.Fields("INVOICENO").Value = txtHeating.text
rs.Fields("INVOICEDATE").Value = Dates.Value
rs.Fields("SUBLEDGER").Value = txtParty.text
rs.Fields("GENLEDGER").Value = "Sundry Debtors"

rs.Fields("sno1").Value = vs.TextMatrix(i, 0)
rs.Fields("type").Value = vs.TextMatrix(i, 1)
rs.Fields("bookgp").Value = vs.TextMatrix(i, 2)
rs.Fields("BOOKCODE").Value = vs.TextMatrix(i, 3)

rs.Fields("TBook").Value = vs.TextMatrix(i, 4) & ""


rs.Fields("rate").Value = vs.TextMatrix(i, 5)
rs.Fields("qty").Value = vs.TextMatrix(i, 6)
rs.Fields("amt").Value = vs.TextMatrix(i, 7)


If rs2.State = 1 Then rs2.Close
rs2.Open "select * from itemmaster where ItemName='" & vs.TextMatrix(i, 3) & "'", con, adOpenDynamic, adLockOptimistic
If rs2.EOF = False Then
If vs.TextMatrix(i, 1) = "Book" Then
 rs!remarks = rs2!titlefarmno & ""
End If
End If


rs.Update

End If

Next


End If

End If

''''
'''''=================================================================
'''''=======================New Code For GST==
''''DoEvents
''''DoEvents
''''DoEvents
''''DoEvents
''''Dim k1 As Integer
''''Dim q1 As Integer
''''Dim rs_ As New ADODB.Recordset
''''
''''k1 = 1
''''con.Execute "delete from InvoicebGST where invoiceno=" & txtHeating & " and firm='" & txtParty.Text & "'"
''''If rs1.State = 1 Then rs1.Close
''''rs1.Open "select distinct BOOKCODE from invoiceb  where (invoiceno=" & txtHeating & "  and SUBLEDGER='" & txtParty.Text & "')", con
''''While rs1.EOF = False
''''   Set rs = New ADODB.Recordset
''''   str1 = "SELECT INVOICEB.INVOICENO, INVOICEB.BOOKCODE, sum(INVOICEB.qty) as qty, sum(INVOICEB.amt) as amt FROM INVOICEB  where (invoiceno=" & txtHeating.Text & " and BOOKCODE='" & rs1(0) & "' and SUBLEDGER='" & txtParty.Text & "')  group by  INVOICEB.INVOICENO, INVOICEB.BOOKCODE"
''''   ''str2 = "SELECT  count(qty) FROM INVOICEB  where (invoiceno=" & txtHeating.Text & " and BOOKCODE='" & rs1(0) & "' and SUBLEDGER='" & txtParty.Text & "') "
''''
''''   Set rs_ = New ADODB.Recordset
''''   rs_.Open "SELECT  count(qty) FROM INVOICEB  where (invoiceno=" & txtHeating.Text & " and BOOKCODE='" & rs1(0) & "' and SUBLEDGER='" & txtParty.Text & "') ", con
''''   If rs_.EOF = False Then
''''      q1 = rs_(0)
''''   End If
''''
''''
''''   rs.Open str1, con
''''   If rs.EOF = False Then
''''
''''
''''    Amt_ = rs(3)
''''    Qty_ = rs(2)
''''
''''    newrate_ = Round(Amt_ / Qty_, 3)
''''    newAmt_ = Round((newrate_ * Qty_), 2)
''''     Qty_ = Qty_ / q1
''''    newrate_ = newrate_ * q1
''''
''''
''''    con.Execute "insert into InvoicebGST(INVOICENO,ItemName,HSNCode,rate,qty,amt,sno1,firm) " & _
''''    " values (" & rs(0) & ",'" & rs(1) & "','" & "******" & "'," & newrate_ & "," & Qty_ & "," & newAmt_ & "," & k1 & ",'" & txtParty.Text & "') "
''''    k1 = k1 + 1
''''   End If
''''
''''rs1.MoveNext
''''Wend
'''''==========================================================
''''



'=================================================================
'=======================New Code For GST==
DoEvents
DoEvents
DoEvents
DoEvents
Dim k1 As Integer
k1 = 1
con.Execute "delete from InvoicebGST where invoiceno=" & txtHeating & " and firm='" & txtParty.text & "'"
If rs1.State = 1 Then rs1.Close
If Check1_ItemWiseTot.Value = 0 Then
   rs1.Open "select distinct BOOKCODE,Qty from invoiceb  where (invoiceno=" & txtHeating & "  and SUBLEDGER='" & txtParty.text & "')", con
Else
   rs1.Open "select distinct BOOKCODE from invoiceb  where (invoiceno=" & txtHeating & "  and SUBLEDGER='" & txtParty.text & "')", con
End If

While rs1.EOF = False
   Set rs = New ADODB.Recordset
   
   If Check1_ItemWiseTot.Value = 0 Then
      str1 = "SELECT INVOICEB.INVOICENO, INVOICEB.BOOKCODE, INVOICEB.qty, sum(INVOICEB.amt) as amt,INVOICEB.hsncode FROM INVOICEB  where (invoiceno=" & txtHeating.text & " and BOOKCODE='" & rs1(0) & "' and SUBLEDGER='" & txtParty.text & "' and Qty=" & rs1!qty & ")  group by  INVOICEB.INVOICENO, INVOICEB.BOOKCODE,INVOICEB.qty,invoiceb.hsncode"
   Else
       str1 = "SELECT INVOICEB.INVOICENO, INVOICEB.BOOKCODE, sum(INVOICEB.qty) as qty, sum(INVOICEB.amt) as amt,hsncode FROM INVOICEB  where (invoiceno=" & txtHeating.text & " and BOOKCODE='" & rs1(0) & "' and SUBLEDGER='" & txtParty.text & "')  group by  INVOICEB.INVOICENO, INVOICEB.BOOKCODE,invoiceb.hsncode"
   End If
   
   rs.Open str1, con
   If rs.EOF = False Then
   
    Amt_ = rs(3)
    Qty_ = rs(2)

    newrate_ = Round(Amt_ / Qty_, 3)
    newAmt_ = Round((newrate_ * Qty_), 2)
   
    con.Execute "insert into InvoicebGST(INVOICENO,ItemName,HSNCode,rate,qty,amt,sno1,firm) " & _
    " values (" & rs(0) & ",'" & rs(1) & "','" & rs!hsncode & "'," & newrate_ & "," & Qty_ & "," & newAmt_ & "," & k1 & ",'" & txtParty.text & "') "
    k1 = k1 + 1
   End If

rs1.MoveNext
Wend
'==========================================================



cmdSave_2.Enabled = False


Exit Sub
aa1:
MsgBox Err.Description


End Sub
Sub updateNewAmt()

Dim sum1_ As Double
Dim sum2_ As Double
Dim sum3_ As Double
 
If rs1.State = 1 Then rs1.Close
rs1.Open "select BOOKCODE,sum(qty),sum(amt) from invoiceb where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "') group by BOOKCODE", con, adOpenDynamic, adLockOptimistic
For i = 1 To rs1.RecordCount
  If rs1.EOF = False Then
        sum1_ = rs1(2) / rs1(1)
        sum3_ = Round(rs1(2) / rs1(1), 3)
        sum2_ = (rs1(1) * sum1_)
        
        con.Execute "update invoiceb set newrate=" & sum3_ & ",newamt=" & sum2_ & ",newQty=" & rs1(1) & " where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "' and BOOKCODE='" & rs1(0) & "')"
        rs1.MoveNext
   End If
Next

End Sub

Sub searchData()

vs.clear

If txtHeating.text = "" Then Exit Sub

If rs.State = 1 Then rs.Close
rs.Open "select * from invoicea where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "')", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then

cmdEdit_4.Enabled = True
cmdSave_2.Enabled = False
cmdDelete_3.Enabled = False


If rs!AgentName <> "" Then

If rs!AgentName = "y" Then
    Check1_ItemWiseTot.Value = 1
Else
    Check1_ItemWiseTot.Value = 0
End If

End If


txtVehiNo.text = rs.Fields("VehicleNo").Value & ""
txtTransMode.text = rs.Fields("TransMode").Value & ""
txtDateOfSupp.text = rs.Fields("DateOfSupp").Value & ""




Me.Dates.Value = Format(rs!invoicedate, "dd/mm/yyyy")

lblEmail.Caption = rs.Fields("email").Value & ""

txtParty.text = rs.Fields("SUBLEDGER").Value
txtRemarks.text = rs.Fields("Remarks").Value & ""
Text1.text = rs.Fields("add1").Value & ""
Text2.text = rs.Fields("add2").Value & ""
'If Not IsNull(rs.Fields("godown").Value) Then
txtAdd = rs.Fields("godown").Value & ""
'Else
'cbogodown.ListIndex = -1

txtTotal1.text = Format(rs.Fields("GAMOUNT").Value, "0.00") & ""
txtNet.text = rs.Fields("NETAMOUNT").Value

txtvatRate.text = rs.Fields("Vatper").Value & ""
txtvatamt.text = rs.Fields("VatAmt").Value & ""


txtVatOn.text = rs.Fields("Vaton").Value & ""
txtBindAmt.text = rs.Fields("bindingAmt").Value & ""

txtCGSTAmt.text = rs.Fields("CGSTAmt").Value & ""
txtCGSTRate.text = rs.Fields("CGSTrate").Value & ""

txtSGSTAmt.text = rs.Fields("SGSTAmt").Value & ""
txtSGSTRate.text = rs.Fields("SGSTRate").Value & ""

txtIGSTAmt.text = rs.Fields("iGSTAmt").Value & ""
txtIGSTRate.text = rs.Fields("iGSTRate").Value & ""

 

TXTpaN.text = rs.Fields("PAN").Value & ""



txtPan_billto.text = rs.Fields("PAN_Billto").Value & ""
txtGST_billto.text = rs.Fields("GSTIN_Billto").Value & ""
txtParty_Shippto.text = rs.Fields("Party_shippto").Value & ""
txtAdd_Shippto.text = rs.Fields("Add_shippto").Value & ""
txtPAN_Shippto.text = rs.Fields("PAN_shippto").Value & ""
txtGST_Shippto.text = rs.Fields("GSTIN_shippto").Value & ""

txtPlaceofsupp.text = rs.Fields("placeofsupp").Value & ""
cboRCharge.text = rs.Fields("rcharge").Value & ""

 txtTotalValue.text = rs.Fields("TotalValue").Value & ""
 txtTGST.text = rs.Fields("TotalGST").Value & ""
txtAddLess.text = rs.Fields("addless").Value & ""

txtTotal2.text = rs!baa & ""



 



If rs.State = 1 Then rs.Close
rs.Open "select * from invoiceb where (INVOICENO=" & txtHeating.text & " and SUBLEDGER='" & Trim(txtParty) & "') order by sno1", con, adOpenDynamic, adLockOptimistic
For i = 1 To rs.RecordCount
If rs.EOF = False Then



vs.TextMatrix(i, 0) = rs.Fields("sno1").Value
vs.TextMatrix(i, 1) = rs.Fields("type").Value
vs.TextMatrix(i, 2) = rs.Fields("bookgp").Value
vs.TextMatrix(i, 3) = rs.Fields("BOOKCODE").Value

vs.TextMatrix(i, 4) = rs.Fields("TBook").Value & ""

vs.TextMatrix(i, 5) = rs.Fields("rate").Value
vs.TextMatrix(i, 6) = rs.Fields("qty").Value
vs.TextMatrix(i, 7) = rs.Fields("amt").Value



rs.MoveNext
End If
Next


If Val(txtIGSTAmt.text) > 0 Then
       Check1_igst.Value = 1
    Else
       Check1_igst.Value = 0
    End If
End If

SetWidth

Total

End Sub
Private Sub cmdUndo_5_Click()
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = False
   cmdPrint_7.Enabled = True
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = False
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
End Sub



Private Sub Dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtParty.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 
If KeyCode = 27 Then
     If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
     End If
 End If
End Sub
Sub showGST()
 Dim rs_ As New ADODB.Recordset
 If rs_.State = 1 Then rs_.Close
 rs_.Open "select vatper from setup1", con
 If rs_.EOF = False Then
    If Not IsNull(rs_(0)) Then
       
       If Check1_igst.Value = 0 Then
            txtCGSTRate.text = (rs_(0) / 2)
            txtSGSTRate.text = (rs_(0) / 2)
            txtIGSTRate.text = 0
       Else
            txtCGSTRate.text = 0
            txtSGSTRate.text = 0
            txtIGSTRate.text = rs_(0)

       End If
       
       
    End If
 End If

End Sub
Sub TotalFinal()
   If txtTotal3.text = "" Then
      txtTotal3.text = 0
   End If
   
   If txtTotal2.text = "" Then
      txtTotal2.text = 0
   End If
   
   
    txtRawAndCasting.text = (CDbl(txtTotal2.text) + CDbl(txtTotal3.text))
    txtRawAndCasting.text = Format(txtRawAndCasting.text, "#,###.000")
End Sub
Private Sub Form_Load()
On Error Resume Next
 
 If rs.State = 1 Then rs.Close
 rs.Open "select vatper,vaton from setup1", con
 If rs.EOF = False Then
    txtvatRate.text = rs(0)
    txtVatOn.text = rs(1) & ""
 End If
 
 
 
 'SetWidth
 AddItemInGrid
 'AddItemInGrid1
 'AddItemInGrid2
 'AddItemInGrid3
 SetWidth
 
 'Dates.Value = Date
 
 
 Dim s As String
 
 s = ""
 
 
 'If rs.State = 1 Then rs.Close
 Set rs = New ADODB.Recordset
 rs.Open "select * from ConsumeGp order by name", con
 While rs.EOF = False
 If s = "" Then
 s = rs(0)
 Else
 s = s & "|" & rs(0)
 End If
 rs.MoveNext
 Wend
 
 vs.ColComboList(2) = s
 
 
If rs.State = 1 Then rs.Close
rs.Open "select max(INVOICENO),subledger from invoicea group by subledger", con, adOpenForwardOnly, adLockReadOnly
If rs.EOF = False Then
   'rs.MoveLast
   'Dates.Value = rs!INVOICEdate
   
   txtParty = rs(1)
   txtHeating.text = rs(0)   'MaxSNo("invoicea", "INVOICENO")
   'txtHeating.Text = Val(txtHeating.Text) - 1
   searchData
   
End If

showGST
cboRCharge.ListIndex = 1


vs.ColComboList(1) = "Book|Title|Inner|HB|I+P|I+C+P|I+C|P+C|Paper|Stich|CD|Supp.|Cutting|Booklet|Poster"
 
End Sub
Sub SetWidth()
vs.Cols = 8
vs.FormatString = "S.No|Type|Book Group|Books Name|Farm|>Rate|>Qty|>Amount"
vs.ColWidth(0) = 600
vs.ColWidth(1) = 1000
vs.ColWidth(2) = 2700
vs.ColWidth(3) = 4500

vs.ColWidth(4) = 1000
vs.ColWidth(5) = 1000
vs.ColWidth(6) = 1000
vs.ColWidth(7) = 1000
'vs.ColWidth(7) = 1100
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then todate.SetFocus
End Sub
Private Sub HeatingDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtParty.SetFocus
End Sub

Private Sub ListHeatingNo_Click()
  Call cmdRef_Click
  searchData
  TotalFinal
  'Frame1.Visible = False
End Sub
Private Sub ToDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then Call cmdadd_Click
End Sub
Private Sub txtGrade_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub mvfrst_Click()

Mode = ""
Dim rsbilmast As ADODB.Recordset
Set rsbilmast = New ADODB.Recordset

If txtParty = "" Then Exit Sub

sq = "Select * from INVOICEA where subledger='" & txtParty & "' order by invoiceNo asc"
rsbilmast.Open sq, con, adOpenKeyset, adLockReadOnly
If rsbilmast.RecordCount > 0 Then
rsbilmast.MoveFirst
txtHeating = rsbilmast.Fields("invoiceNo")
searchData
End If



End Sub

Private Sub Mvlst_Click()
Mode = ""


Dim rsbilmast As ADODB.Recordset
Set rsbilmast = New ADODB.Recordset

If txtParty = "" Then Exit Sub

sq = "Select * from INVOICEA where subledger='" & txtParty & "' order by invoiceNo asc"
rsbilmast.Open sq, con, adOpenKeyset, adLockReadOnly

If rsbilmast.RecordCount > 0 Then
    rsbilmast.MoveLast
    Me.txtHeating.text = rsbilmast.Fields("INVOICENO")
    searchData
Else
    MsgBox "No record", vbInformation
End If
End Sub

Private Sub Mvnxt_Click()
Dim rsbilmast As ADODB.Recordset
Set rsbilmast = New ADODB.Recordset

If txtHeating = "" Then Exit Sub

If txtParty = "" Then Exit Sub

sq = "Select * from invoicea where INVOICENO >= " + txtHeating + " and subledger='" & txtParty & "' order by INVOICENO asc"
tmpbillno = txtHeating
rsbilmast.Open sq, con, adOpenKeyset, adLockReadOnly

If rsbilmast.RecordCount > 1 Then
If Not rsbilmast.EOF Then

Do While Val(tmpbillno) = rsbilmast.Fields("INVOICENO")
   rsbilmast.MoveNext
   If rsbilmast.EOF Then
      MsgBox " There is no more record", vbInformation
      Exit Sub
End If
Loop

Me.txtHeating.text = rsbilmast.Fields("INVOICENO")
searchData

End If
Else
   rsbilmast.MoveLast
   Me.txtHeating.text = rsbilmast.Fields("INVOICENO")
   searchData
   
End If

End Sub

Private Sub Mvprv_Click()

Dim rsbilmast As ADODB.Recordset
Set rsbilmast = New ADODB.Recordset

If txtParty = "" Then Exit Sub

If txtHeating = "" Then Exit Sub
sq = "Select * from invoicea where subledger='" & txtParty & "'  order by INVOICENO asc"
rsbilmast.Open sq, con, adOpenKeyset, adLockReadOnly
If rsbilmast.RecordCount >= 1 Then

If Not rsbilmast.EOF Then


rsbilmast.MoveFirst
rsbilmast.Find "invoiceno=" & txtHeating & ""
If rsbilmast.EOF = False Then
    rsbilmast.MovePrevious
    
  If (rsbilmast.EOF = False) Then
  
   If rsbilmast.BOF = False Then
    Me.txtHeating.text = rsbilmast.Fields("INVOICENO")
    searchData
   End If
  
  End If
    
End If


End If

Else
    rsbilmast.MoveFirst
    Me.txtHeating.text = rsbilmast.Fields("INVOICENO")
    searchData

End If


End Sub

Private Sub txtAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
vs.SetFocus
vs.Col = 1
End If

End Sub

Private Sub txtDateOfSupp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

 txtRemarks.SetFocus

End If

End Sub

Private Sub txtHeating_GotFocus()

txtHeating.SelLength = 10

If PopUpValue1 <> "" Then

txtHeating.text = PopUpValue1
txtParty = PopUpValue3
Dates.Value = PopUpValue2

vs.clear
SetWidth
searchData
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
End If
End Sub

Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 113 Or KeyCode = 255) Then
popuplist2 "select INVOICENO,INVOICEDATE,SUBLEDGER,remarks as Party from invoicea where subledger='" & txtParty & "' order by INVOICENO", con
End If

If KeyCode = 13 Then

searchData
End If

End Sub

Private Sub txtHeating_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
        
   Dates.SetFocus

        
  End If
  

End Sub
Private Sub txtParty_GotFocus()
If PopUpValue1 <> "" Then

    txtParty.text = PopUpValue1
    FNAME = PopUpValue1
    Text1.text = PopUpValue2
    Text2.text = PopUpValue3
    TXTpaN.text = PopUpValue4
    lblEmail.Caption = PopUpValue5
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    PopUpValue4 = ""
    PopUpValue5 = ""
    
    clear
    
End If
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 113 Or KeyCode = 255) Then
popuplist2 "select Name,Address,TIN,CST AS PAN,Email from firm order by name", con
End If
If KeyCode = 13 Then

If txtParty <> "" Then
 txtDateOfSupp.SetFocus
End If

End If

End Sub

Private Sub txtParty_LostFocus()
   Record = ""
End Sub
Private Sub txtQty_GotFocus()
     txtQty.SelLength = 10
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtParty_Shippto_GotFocus()
If PopUpValue1 <> "" Then
    txtParty_Shippto.text = PopUpValue1
    txtAdd_Shippto.text = PopUpValue2
    
    txtGST_Shippto.text = PopUpValue3
    txtPAN_Shippto.text = PopUpValue4
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    PopUpValue4 = ""
End If

End Sub

Private Sub txtParty_Shippto_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 113 Or KeyCode = 255) Then
popuplist2 "select Name,Address,Tin as GSTIN,CST As PAN from customer order by name", con
End If

If KeyCode = 13 Then
    
     If txtParty_Shippto = "" Then Exit Sub
     vs.SetFocus
     vs.Col = 1
     
     If vs.Row = 1 Then
        vs.TextMatrix(vs.RowSel, 0) = 1
     End If
     
     
    For i = 1 To vs.Rows - 1
       If vs.TextMatrix(i, 1) <> "" Then
          vs.Row = i
          s1 = 1
       End If
    Next
    
    If s1 = 0 Then
       vs.Col = 1
    End If

End If


End Sub

Private Sub txtRemarks_GotFocus()
If PopUpValue1 <> "" Then
    txtRemarks.text = PopUpValue1
    txtAdd.text = PopUpValue2
    txtGST_billto.text = PopUpValue3
    txtPan_billto.text = PopUpValue4
    
    txtParty_Shippto.SetFocus
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    PopUpValue4 = ""
End If
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 113 Or KeyCode = 255) Then
popuplist2 "select Name,Address,Tin as GSTIN,CST As PAN from customer order by name", con
End If

If KeyCode = 13 Then


'txtRemarks.SetFocus
End If

End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)

s1 = 0

If KeyAscii = 13 Then
 
 
If txtRemarks = "" Then
   txtRemarks.SetFocus
   Exit Sub
End If
 
 'SendKeys "{tab}"
' vs.SetFocus
' vs.Col = 1
'
' If vs.Row = 1 Then
'    vs.TextMatrix(vs.RowSel, 0) = 1
' End If
'
'
'For i = 1 To vs.Rows - 1
'   If vs.TextMatrix(i, 1) <> "" Then
'      vs.Row = i
'      s1 = 1
'   End If
'Next
'
'If s1 = 0 Then
'   vs.Col = 1
'End If

End If

End Sub
Private Sub txtSize_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtGrade.SetFocus
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If (vs.Col = 0 Or vs.Col = 1) Then
        vs.TextMatrix(vs.RowSel, 0) = vs.Row
        
     ElseIf vs.Col = 3 Then
        cellposi
        If cboItem.text <> "" Then
        vs.TextMatrix(vs.RowSel, 0) = cboItem.text
        End If
     End If
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
  If (KeyCode = 115 Or KeyCode = 46) Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    
    For i = 1 To vs.Rows - 1
       If vs.TextMatrix(i, 1) <> "" Then
          vs.TextMatrix(i, 0) = i
       End If
    Next
    
    Total
    calc
  End If
  End If
  
  If KeyCode = 13 Then
     
     If vs.Col = 3 Then
        vs.Editable = flexEDNone
        cboItem.Visible = True
        cboItem.SetFocus
     Else
        vs.Editable = flexEDKbdMouse
        cellposi
     End If
     
     If vs.Col = 4 Then
       sendkeys "{right}"
     End If
     
     If vs.Col = 5 Then
       sendkeys "{right}"
     End If

     

  End If
  
  
  
  
  
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
          
 If vs.Col = 1 Then
    If vs.TextMatrix(vs.RowSel, 1) <> "" Then
       '  SendKeys "{right}"
       
        If vs.Row >= 2 Then
        If vs.TextMatrix(vs.RowSel - 1, 1) = "Book" Then
           vs.TextMatrix(vs.RowSel, 2) = vs.TextMatrix(vs.RowSel - 1, 2)
        End If
        End If
       
       vs.Col = 2
    End If
 ElseIf vs.Col = 2 Then
    If vs.TextMatrix(vs.RowSel, 2) <> "" Then
         AddItemInGrid
         'SendKeys "{right}"
         
         
 
         
         vs.Col = 3
    End If
 End If
 
 '=============================================
 
 If vs.Col = 3 Then
 
      
      vs.Editable = flexEDNone
      cboItem.Visible = True
      cboItem.ZOrder
      
      If vs.Row >= 2 Then
        If vs.TextMatrix(vs.RowSel - 1, 1) = "Book" Then
           cboItem.text = vs.TextMatrix(vs.RowSel - 1, 3)
        End If
      End If
      
      cboItem.Top = vs.Top + vs.CellTop
      cboItem.Left = vs.CellLeft + 250
      cboItem.Width = vs.ColWidth(vs.Col)
      cboItem.SetFocus
          
  If vs.TextMatrix(vs.RowSel, 0) <> "" Then
       'SendKeys "{right}"
  End If
       
  
    
 End If
 
 
 If vs.Col = 4 Then
 
    If (vs.TextMatrix(vs.RowSel, 2) <> "" And vs.TextMatrix(vs.RowSel, 3) <> "") Then
     viewData vs.TextMatrix(vs.RowSel, 3), vs.TextMatrix(vs.RowSel, 2)
    End If

    If Val(vs.TextMatrix(vs.RowSel, 4)) > 0 Then
       sendkeys "{right}"
    End If
 End If
 
 If vs.Col = 5 Then
        'SendKeys "{right}"
        'SendKeys "{right}"
        If Val(vs.TextMatrix(vs.RowSel, 5)) > 0 Then
        vs.Col = 6
        End If
        
        'Exit Sub
 End If
    
If vs.Col = 6 Then
    
 If Val(vs.TextMatrix(vs.RowSel, 5)) > 0 Then
    'SendKeys "{right}"
 '   vs.Col = 7
 End If
'    SendKeys "{right}"
'    vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, vs.TextMatrix(vs.RowSel, 1)) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))
'
End If


If vs.Col = 6 Then

If vs.TextMatrix(vs.RowSel, 6) = "" Then
   vs.Col = 6
   vs.Editable = flexEDKbdMouse
   Exit Sub
End If

If vs.TextMatrix(vs.RowSel, 1) = "Book" Then
sum1 = ((Val(vs.TextMatrix(vs.RowSel, 5)) / 1000) * Val(vs.TextMatrix(vs.RowSel, 4)))
vs.TextMatrix(vs.RowSel, 7) = (sum1 * Val(vs.TextMatrix(vs.RowSel, 6)))
Else
vs.TextMatrix(vs.RowSel, 7) = (Val(vs.TextMatrix(vs.RowSel, 5)) * Val(vs.TextMatrix(vs.RowSel, 6)))
End If


cboItem.text = ""

vs.Col = 1
vs.Row = vs.Row + 1

'SendKeys "{home}"
'SendKeys "{down}"

vs.TextMatrix(vs.RowSel, 0) = vs.Row
vs.Col = 1

Total
calc

End If
    
       
Total

End If

End Sub



Private Sub vs_LeaveCell()
  Total
End Sub

Private Sub vs1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs1.Col = 0 Then
        cellposiVs
        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.text
     End If

End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
    vs1.RemoveItem (vs1.RowSel)
    'Total1
    TotalFinal
  End If
  
  If KeyCode = 13 Then
     If vs1.Col = 0 Then
        vs1.Editable = flexEDNone
        Vs1Frame.Visible = True
        cboitemvs1.Visible = True
        cboitemvs1.SetFocus
     Else
        vs1.Editable = flexEDKbdMouse
        cellposiVs
     End If
  End If
End Sub

Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
        
 If vs1.Col = 0 Then
    vs1.Editable = flexEDNone
    Vs1Frame.Visible = True
    cboitemvs1.SetFocus
          
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", con
    If rs.EOF = False Then
       vs1.TextMatrix(vs1.RowSel, 1) = rs.Fields("Unit").Value
       sendkeys "{right}"
       sendkeys "{right}"
       Vs1Frame.Visible = False
       vs1.Editable = flexEDKbdMouse
       vs1.SetFocus
    Else
       vs1.TextMatrix(vs1.RowSel, 1) = "Kg"
       sendkeys "{right}"
       sendkeys "{right}"
       Vs1Frame.Visible = False
       vs1.Editable = flexEDKbdMouse
       vs1.SetFocus
    
    End If
    
 End If
    
 If vs1.Col = 2 Then
           
    sendkeys "{home}"
    sendkeys "{down}"
    
    AddItemInGrid1
 End If
    
    

 'Total1
 TotalFinal

End If


End Sub
Sub AddSemifinish()
   Dim j As Integer
   
   j = 1
    
   vs3.clear
   For i = 1 To vs1.Rows - 1
    
   If vs1.TextMatrix(i, 0) <> "" Then
      If rs.State = 1 Then rs.Close
      rs.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(i, 0) & "'", con
      If rs.Fields("itemgp").Value = "Semi Finish (R/D)" Or rs.Fields("itemgp").Value = "Semi Finish (Store)" Then
         vs3.TextMatrix(j, 0) = vs1.TextMatrix(i, 0)
         vs3.TextMatrix(j, 1) = vs1.TextMatrix(i, 1)
         vs3.TextMatrix(j, 2) = vs1.TextMatrix(i, 2)
         j = j + 1
      End If
   End If
        
   Next
    
    
End Sub
Private Sub vs1_LeaveCell()
   'Total1
End Sub

Private Sub vs2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs2.Col = 0 Then
        'cellposiVs3
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.text
     End If

End Sub

'''Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
''''' If KeyCode = 46 Then
'''''    vs2.RemoveItem (vs2.RowSel)
'''''    'Total2
'''''    TotalFinal
''''' End If
'''''
'''''
'''''  If KeyCode = 13 Then
'''''
'''''     If vs2.Col = 0 Then
'''''        vs2.Editable = flexEDNone
'''''        Vs3Frame.Visible = True
'''''        cboItemVs3.Visible = True
'''''        cboItemVs3.SetFocus
'''''     Else
'''''        vs2.Editable = flexEDKbdMouse
'''''        'cellposiVs3
'''''     End If
'''''
'''''  End If
'''
'''End Sub

''''Private Sub vs2_KeyUp(KeyCode As Integer, Shift As Integer)
''''If KeyCode = 13 Then
''''
''''
'''' If vs2.Col = 0 Then
''''
''''      vs2.Editable = flexEDNone
''''      Vs3Frame.Visible = True
''''
''''
''''
''''    Set rs = New ADODB.Recordset
''''    If rs.State = 1 Then rs.Close
''''    rs.Open "select * from ItemMaster where ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", con
''''    If rs.EOF = False Then
''''       vs2.TextMatrix(vs2.RowSel, 1) = rs.Fields("Unit").Value
''''       SendKeys "{right}"
''''       SendKeys "{right}"
''''       Vs3Frame.Visible = False
''''       vs2.Editable = flexEDKbdMouse
''''       vs2.SetFocus
''''    Else
''''       vs2.TextMatrix(vs2.RowSel, 1) = "Kg"
''''       SendKeys "{right}"
''''       SendKeys "{right}"
''''       Vs3Frame.Visible = False
''''       vs2.Editable = flexEDKbdMouse
''''       vs2.SetFocus
''''
''''    End If
''''
'''' End If
''''
''''
''''    If vs2.Col = 2 Then
''''
''''           SendKeys "{home}"
''''           SendKeys "{down}"
''''           Vs3Frame.Top = Vs3Frame.Top + 170
''''    End If
''''
''''
''''   'Total2
''''
''''End If
''''
''''End Sub
''''Private Sub vs2_LeaveCell()
''''   'Total2
''''End Sub
''''Private Sub vs3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
''''     If vs3.Col = 0 Then
''''        'cellposiVs2
''''        'vs3.TextMatrix(vs3.RowSel, 0) = cboitemvscboItemVs2.Text
''''     End If
''''
''''End Sub
''''
''''Private Sub vs3_KeyDown(KeyCode As Integer, Shift As Integer)
''''
''''  If KeyCode = 46 Then
''''    vs3.RemoveItem (vs3.RowSel)
''''    'Total4
''''  End If
''''
''''  If KeyCode = 13 Then
''''     If vs3.Col = 0 Then
''''
''''        vs3.Editable = flexEDNone
''''        FrameVs2.Visible = True
''''        cboItemVs2.Visible = True
''''        cboItemVs2.SetFocus
''''     Else
''''
''''        vs3.Editable = flexEDKbdMouse
''''
''''     End If
''''  End If
''''
''''End Sub
''''
''''Private Sub vs3_KeyUp(KeyCode As Integer, Shift As Integer)
''''
''''If KeyCode = 13 Then
''''
'''' If vs3.Col = 0 Then
''''    vs3.Editable = flexEDNone
''''    FrameVs2.Visible = True
''''    cboItemVs2.SetFocus
''''
''''
''''
''''
''''    Set rs = New ADODB.Recordset
''''    If rs.State = 1 Then rs.Close
''''    rs.Open "select * from ItemMaster where ItemName='" & vs3.TextMatrix(vs3.RowSel, 0) & "'", con
''''    If rs.EOF = False Then
''''       vs3.TextMatrix(vs3.RowSel, 1) = rs.Fields("Unit").Value
''''       SendKeys "{right}"
''''       SendKeys "{right}"
''''       FrameVs2.Visible = False
''''       vs3.Editable = flexEDKbdMouse
''''       vs3.SetFocus
''''    Else
''''       vs3.TextMatrix(vs3.RowSel, 1) = "Kg"
''''       SendKeys "{right}"
''''       SendKeys "{right}"
''''       FrameVs2.Visible = False
''''       vs3.Editable = flexEDKbdMouse
''''       vs3.SetFocus
''''
''''    End If
''''
'''' End If
''''
'''' If vs3.Col = 2 Then
''''
''''   If rs.State = 1 Then rs.Close
''''   rs.Open "select  OpeningStock from ItemMaster where ItemName='" & vs3.TextMatrix(vs.RowSel, 0) & "'", con
''''   If rs.EOF = False Then
''''      If Val(rs.Fields(0).Value) < Val(vs3.TextMatrix(vs3.RowSel, 2)) Then
''''         MsgBox "Stock Less !!", vbInformation
''''
''''      End If
''''   End If
''''
''''
''''    SendKeys "{home}"
''''    SendKeys "{down}"
''''
''''    FrameVs2.Top = FrameVs2.Top + 170
''''    'AddItemInGrid2
'''' End If
''''
''''
''''
'''' 'Total4
''''
''''End If
''''
''''End Sub
Private Sub vs_SelChange()

If vs.Col = 5 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 4 Then
   vs.Editable = flexEDKbdMouse
End If

End Sub

