VERSION 5.00
Begin VB.Form frmreset 
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreset 
      Caption         =   "Reset Databse"
      Height          =   510
      Left            =   1215
      TabIndex        =   0
      Top             =   765
      Width           =   2175
   End
End
Attribute VB_Name = "frmreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdreset_Click()

a = InputBox("Enter Code :")

If a = 999731681 Then


con.Execute "delete from BookReceive"
con.Execute "delete from FoldingDetails"
con.Execute "delete from Invoice"
con.Execute "delete from INVOICEA"

con.Execute "delete from INVOICEB"
con.Execute "delete from InvoicebGST"

con.Execute "delete from ITC"

con.Execute "delete from payment"
con.Execute "delete from Receipt"

con.Execute "delete from Title"

con.Execute "delete from TitleDel"

con.Execute "delete from titleStatent"

MsgBox "Database is clear succesfully", vbInformation

Else

MsgBox "Data is not set"

End If




End Sub
