Attribute VB_Name = "ConnectionModule"
Option Explicit
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public cmd As New ADODB.Command
'Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Public rel_tab As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public q As Integer
Public fyear As String
Public PopUpValue1 As String
Public PopUpValue2 As String
Public PopUpValue3 As String
Public PopUpValue4 As String
Public PopUpValue5 As String
Public Pay_rec As String
Public colmn As Integer
Public fs1 As New FileSystemObject
Public ss, ss1 As String
Public gst As Double
Public bill_format As String
Public ch_ As Boolean
Public chk_done As Boolean

Public Record

Public tbl As String
Public f1 As File
Public INVOICETYPE As String
Public txt1 As TextStream
Public COMPNAME As String
Public add1 As String
Public add2 As String
Public ADD3 As String
Public FNAME As String
Public Sub sendkeys(text As Variant, Optional wait As Boolean = False)
Dim wshshell As Object
Set wshshell = CreateObject("wscript.shell")
wshshell.sendkeys CStr(text), wait
Set wshshell = Nothing
End Sub
Public Function billFormat(bill_ As String) As String

Dim yrs_ As String
Dim BILL1
BILL1 = bill_

If rs1.State = 1 Then rs1.Close
rs1.Open "select Invoivcecondition from setup1", con, adOpenDynamic, adLockOptimistic
If rs1.EOF = False Then
   yrs_ = rs1(0) & ""
End If



If Len(bill_) > 0 Then
   bill_ = "B" & Format(Trim(bill_), "0000") & "/" & yrs_
   billFormat = bill_
End If


If FNAME = "DEVA BOOK BINDING HOUSE" Then
   bill_ = "A" & Format(Trim(BILL1), "0000") & "/" & yrs_
   billFormat = bill_
   
End If

End Function
' =========================
' ?? CHANGE DATE HERE ONLY
' =========================
Public Function GetCurrentDate() As Date

    ' ?? NORMAL MODE (production)
    GetCurrentDate = Date

    ' ?? TEST CASES (uncomment ONE at a time)

    'GetCurrentDate = Date + 350   ' far from expiry
    'GetCurrentDate = Date + 360   ' near expiry (warning should come)
    'GetCurrentDate = Date + 364   ' last day warning
    'GetCurrentDate = Date + 366   ' expired ? should ask key

    ' ?? OR fixed testing date
    'GetCurrentDate = #03/25/2027#
    
End Function


' =========================
' ?? MAIN LICENSE CHECK
' =========================
Public Function check_session() As String

    Dim expDateStr As String
    Dim expDate As Date
    Dim d1 As Integer
    Dim userKey As String

    Call ReadLicense(expDateStr)
    
    ' ?? Safe conversion
    If IsDate(expDateStr) Then
        expDate = CDate(expDateStr)
    Else
        expDate = CDate("03/31/2027")
        SaveLicense Format(expDate, "DD-MMM-YYYY")
    End If

    d1 = DateDiff("d", GetCurrentDate(), expDate)

   

    ' ?? WARNING
    If d1 >= 0 And d1 <= 15 Then
        MsgBox "Software will expire in " & d1 & " days!", vbExclamation
    End If

    ' ? EXPIRED
    If d1 < 0 Then
        
        userKey = InputBox("License Expired!" & vbNewLine & _
                           "Enter Renewal Key:", "Renewal Required")

        If Trim(userKey) = "" Then
            MsgBox "License required!", vbExclamation
            End
        End If

        If VerifyRenewalKey(userKey) Then
            
            expDate = DateAdd("d", 365, GetCurrentDate())
            Call SaveLicense(Format(expDate, "DD-MMM-YYYY"))
            
            MsgBox "Renewed! Valid till: " & expDate, vbInformation
            
        Else
            MsgBox "Invalid Key!", vbCritical
            End
        End If

    End If

End Function


' =========================
' ?? KEY LOGIC
' =========================
Public Function VerifyRenewalKey(inputKey As String) As Boolean

    Dim expectedKey As String
    
    expectedKey = "RENEW" & Year(GetCurrentDate()) & Format(Month(GetCurrentDate()), "00")
    
    If UCase(Trim(inputKey)) = expectedKey Then
        VerifyRenewalKey = True
    Else
        VerifyRenewalKey = False
    End If

End Function


' =========================
' ?? SAVE LICENSE
' =========================
Public Sub SaveLicense(expDate As String)

    Dim fs As New FileSystemObject
    Dim txt As TextStream
    Dim licPath As String

    licPath = App.Path & "\lic.dat"

    On Error Resume Next
    SetAttr licPath, vbNormal
    On Error GoTo 0

    Set txt = fs.CreateTextFile(licPath, True)
    txt.WriteLine EncryptStr(expDate)
    txt.Close

    SetAttr licPath, vbHidden

End Sub


' =========================
' ?? READ LICENSE
' =========================
Public Sub ReadLicense(ByRef expDate As String)

    Dim fs As New FileSystemObject
    Dim txt As TextStream
    Dim licPath As String

    licPath = App.Path & "\lic.dat"

    If fs.FileExists(licPath) Then
        
        On Error Resume Next
        Err.clear
        
        SetAttr licPath, vbNormal
        Set txt = fs.OpenTextFile(licPath, ForReading)
        
        If Not txt.AtEndOfStream Then
            expDate = DecryptStr(txt.ReadLine)
        Else
            expDate = ""
        End If
        
        txt.Close
        SetAttr licPath, vbHidden
        
        ' ?? fallback (file corrupted or empty)
        If Err.Number <> 0 Or expDate = "" Or Not IsDate(expDate) Then
            expDate = Format(DateAdd("d", 360, Date), "DD-MMM-YYYY")
            SaveLicense expDate
        End If
        
        On Error GoTo 0
        
    Else
        ' ?? FIRST RUN ? set expiry after 30 days
        expDate = Format(DateAdd("d", 360, Date), "DD-MMM-YYYY")
        SaveLicense expDate
    End If

End Sub


' =========================
' ?? ENCRYPT / DECRYPT
' =========================
Public Function EncryptStr(s As String) As String
    Dim i As Integer, result As String
    For i = 1 To Len(s)
        result = result & Chr(Asc(Mid(s, i, 1)) + 3)
    Next i
    EncryptStr = result
End Function

Public Function DecryptStr(s As String) As String
    Dim i As Integer, result As String
    For i = 1 To Len(s)
        result = result & Chr(Asc(Mid(s, i, 1)) - 3)
    Next i
    DecryptStr = result
End Function
Public Function formatstr(s As String, spaces As Integer, l_r As String, maxlen As Integer) As String
    Dim i As Integer
    Dim tempstr As String
    On Error GoTo last
    If Len(s) > maxlen Then
    s = Mid(s, 1, maxlen)
    End If
    If l_r = "R" Or l_r = "r" Then
    tempstr = tempstr & Space(maxlen - Len(s))
    tempstr = tempstr & s & Space(2)
    Else
    tempstr = s & Space(maxlen - Len(s))
    End If
    formatstr = tempstr
    Exit Function
last:
    formatstr = Space(maxlen)
End Function
Public Sub busy()
Screen.MousePointer = vbHourglass
End Sub
Public Sub free()
 Screen.MousePointer = vbDefault
End Sub
Public Sub pos(frm As Form)
'frm.Top = 860
'frm.Left = 4680
frm.BackColor = &HC0E0FF
End Sub
Sub Main()
  ch_ = True      ' ? ADD THIS
    chk_done = False ' ? ADD THIS
    
    
    Dim licResult As String
    licResult = check_session()
    If licResult <> "" Then
        MsgBox licResult, vbCritical
        End
    End If
    ' *** END LICENSE CHECK ***
    
    
    On Error GoTo ConnError
    
    Dim connStr As String
    Dim appPath As String
    Dim yearFolder As String
    Dim fs As New FileSystemObject
    
    ' Clean app path (remove trailing backslash if present)
    appPath = App.Path
    If Right(appPath, 1) = "\" Then
        appPath = Left(appPath, Len(appPath) - 1)
    End If
    
    ' Get year from combo box
    yearFolder = Trim(frmPassword.cboyrs.text)
    
    ' Validate year is not empty
    If Len(yearFolder) = 0 Then
        MsgBox "No financial year selected!", vbCritical
        End
    End If
    
    ' Build full path
    connStr = appPath & "\" & yearFolder & "\Data.mdb"
    
    ' Check if year folder exists
    If Not fs.FolderExists(appPath & "\" & yearFolder) Then
        MsgBox "Year folder not found:" & vbNewLine & appPath & "\" & yearFolder, vbCritical
        End
    End If
    
    ' Check if Data.mdb exists inside that folder
    If Not fs.FileExists(connStr) Then
        MsgBox "Data.mdb not found:" & vbNewLine & connStr, vbCritical
        End
    End If
    
    ' Open connection
    If con.State = 1 Then con.Close
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & connStr
    con.Open
    con.CursorLocation = adUseClient
    
    ii
    
    
    
    '-----------------------------
    On Error Resume Next
    
    con.Execute "alter table Receipt add DrCr text(10)"
    con.Execute "alter table Receipt add PayAmt double"
    con.Execute "alter table INVOICEA add VehicleNo text(30)"
    con.Execute "alter table INVOICEA add TransMode text(30)"
    con.Execute "alter table INVOICEA add DateOfSupp text(30)"
    con.Execute "alter table Winrpt add issueno Long Integer"
    con.Execute "alter table ItemMaster add remarks text(100)"
    con.Execute "alter table ItemMaster add hsncode text(10)"
    con.Execute "alter table invoiceb add hsncode text(10)"
    con.Execute "alter table InvoicebGST add firm text(150)"
    con.Execute "alter table itc add bgp text(30)"
    con.Execute "alter table BookReceive add bgp text(30)"
    con.Execute "alter table ItemMaster ALTER COLUMN TitleFarmNo TEXT(200)"
    con.Execute "alter table titleStatent add des1 text(100)"
    con.Execute "alter table titleStatent add binder text(100)"
    con.Execute "alter table INVOICEA add pan text(20)"
    con.Execute "alter table INVOICEA add email text(50)"
    con.Execute "alter table setup1 add GST double"
    con.Execute "alter table firm add bank text(40),Account text(40),ifsc text(40)"
    con.Execute "alter table INVOICEA add IGSTAmt double"
    con.Execute "alter table INVOICEA add IGSTrate double"
    con.Execute "alter table INVOICEA add CGSTAmt double"
    con.Execute "alter table INVOICEA add CGSTrate double"
    con.Execute "alter table INVOICEA add SGSTAmt double"
    con.Execute "alter table INVOICEA add SGSTrate double"
    con.Execute "alter table INVOICEA add TotalValue double"
    con.Execute "alter table INVOICEA add TotalGST double"
    con.Execute "alter table INVOICEA add addless double"
    con.Execute "alter table INVOICEA add Stcode_Billto text(5)"
    con.Execute "alter table INVOICEA add State_Billto text(40)"
    con.Execute "alter table INVOICEA add Stcode_shippto text(5)"
    con.Execute "alter table INVOICEA add State_shippto text(40)"
    con.Execute "alter table INVOICEA add PAN_Billto text(20)"
    con.Execute "alter table INVOICEA add GSTIN_Billto text(20)"
    con.Execute "alter table INVOICEA add PAN_shippto text(20)"
    con.Execute "alter table INVOICEA add GSTIN_shippto text(20)"
    con.Execute "alter table INVOICEA add Party_shippto text(60)"
    con.Execute "alter table INVOICEA add Add_shippto text(60)"
    con.Execute "alter table INVOICEA add placeofsupp text(40)"
    con.Execute "alter table INVOICEA add rcharge text(10)"
    con.Execute "alter table INVOICEB add NewRate double"
    con.Execute "alter table INVOICEB add NewAmt double"
    con.Execute "alter table INVOICEB add NewQty double"
    
    Exit Sub

ConnError:
    MsgBox "Connection Error " & Err.Number & ": " & Err.Description & vbNewLine & _
           "Path: " & connStr, vbCritical
    Open appPath & "\error.log" For Append As #1
    Print #1, Now & " - Error " & Err.Number & ": " & Err.Description & " | Path: " & connStr
    Close #1
    End
End Sub
Sub DSN()
On Error GoTo DSNError

Dim FSO As FileSystemObject
Dim f As File
Dim txt As TextStream
Dim matter As String
Dim Total As String
Dim s(1, 2) As String
Set FSO = New FileSystemObject
Dim ss As String

If FSO.FolderExists("C:\Program Files\Common Files\ODBC\Data Sources") = True Then
ss = "C"
ElseIf FSO.FolderExists("D:\Program Files\Common Files\ODBC\Data Sources") = True Then
ss = "D"
ElseIf FSO.FolderExists("E:\Program Files\Common Files\ODBC\Data Sources") = True Then
ss = "E"
ElseIf FSO.FolderExists("F:\Program Files\Common Files\ODBC\Data Sources") = True Then
ss = "F"
End If

If ss = "" Then
ss = "C"
End If

Dim op_system As String
Dim Dusername As String
Dim dstrpath As String

Open App.Path & "\soft.mdb" For Input As #1
Line Input #1, op_system
Close #1

Open App.Path & "\user.ini" For Input As #1
Line Input #1, Dusername
Close #1

If Right(op_system, 1) = "x" Then
   Set txt = FSO.CreateTextFile(ss & ":\Progra~1\Common~1\ODBC\DataSo~1\JKpayment.dsn")
Else
   dstrpath = "C:\Users\" & Dusername & "\Documents\JKpayment.dsn"
   Set txt = FSO.CreateTextFile(dstrpath)
End If

matter = matter & "[ODBC]" & vbNewLine
matter = matter & "DRIVER=Microsoft Access Driver (*.mdb)" & vbNewLine
matter = matter & "UID = admin" & vbNewLine
matter = matter & "UserCommitSync = Yes" & vbNewLine
matter = matter & "Threads = 3" & vbNewLine
matter = matter & "afeTransactions = 0" & vbNewLine
matter = matter & "PageTimeout = 5" & vbNewLine
matter = matter & "MaxScanRows = 8" & vbNewLine
matter = matter & "MaxBufferSize = 2048" & vbNewLine
matter = matter & "FIL=MS Access" & vbNewLine
matter = matter & "DriverId = 25" & vbNewLine
matter = matter & "DefaultDir=" & App.Path & vbNewLine
matter = matter & "DBQ=" & App.Path & "\" & frmPassword.cboyrs.text & "\Data.mdb"

txt.Write matter
txt.Close

Call Main

Exit Sub    ' <-- THIS stops normal flow from falling into error handler

DSNError:   ' <-- THIS must be INSIDE the sub, before End Sub
    Dim errMsg As String
    errMsg = "Error in DSN()" & vbNewLine & _
             "Error No : " & Err.Number & vbNewLine & _
             "Description : " & Err.Description & vbNewLine & vbNewLine & _
             "App.Path = " & App.Path & vbNewLine & _
             "Year Selected = " & frmPassword.cboyrs.text & vbNewLine & _
             "Username = " & Dusername & vbNewLine & _
             "OS flag = " & op_system
    MsgBox errMsg, vbCritical, "Startup Failed"
    Open App.Path & "\error.log" For Append As #2
    Print #2, Now & " | " & errMsg
    Close #2
    End

End Sub     ' <-- only ONE End Sub at the very bottom
Public Function str_val(j As TextBox, i As Integer) As Boolean
 
 If (i >= 65 And i <= 90) Or (i >= 97 And i <= 122) Or (i = 13) Or (i = 32) Or (i = 8) Then
  str_val = True
  Else
  str_val = False
  End If
End Function
Public Function val_int(i As TextBox, j As Integer) As Boolean
Dim a As Boolean
If j >= 48 And j <= 57 Or j = 8 Or j = 13 Then

val_int = True
Else
val_int = False
End If

End Function
            'Developer: Dinesh Saini
            'Get Max + 1 Number from Perticular Number Field from a table
Public Function MaxSNo(tbl As String, fld As String) As Double
    Dim rs As New Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "Select max(" & fld & ") from " & tbl, con
    If IsNull(rs(0)) Then
        MaxSNo = 1
    Else
        MaxSNo = Val(rs(0)) + 1
    End If
    rs.Close
End Function

            'Developer: Dinesh Saini
            'Get Max + 1 Number from Perticular Number Field from a table
Public Function DuplicateString(tbl As String, fld As String, ValStr As String) As Boolean
    Dim rs As New Recordset
    rs.Open "Select * from " & tbl & " Where " & fld & " ='" & ValStr & "'", con
    If rs.RecordCount = 0 Then
         DuplicateString = False
    Else
         DuplicateString = True
    End If
    rs.Close
End Function


            'Developer: Dinesh Saini
            'Get Max + 1 Number from Perticular Number Field from a table
Public Function DuplicateNumber(tbl As String, fld As String, ValNum As Double) As Boolean
    Dim rs As New Recordset
    rs.Open "Select * from " & tbl & " Where " & fld & " =" & ValNum, con
    If rs.RecordCount = 0 Then
         DuplicateNumber = False
    Else
         DuplicateNumber = True
    End If
    rs.Close
End Function
'-----------------------------------Coded by dinesh Saini---------------------------------------------
'-----------------------------To delete dependencies after deleting records--------------------------------
Sub chk_dep(dest_tab As String, field_dest As String, src_field As String)
If rel_tab.State = adStateOpen Then
    rel_tab.Close
End If
End Sub
Function DeleteRecord(V, fld, table) As Boolean
On Error GoTo err1
    con.Execute ("delete from " & table & " where cstr(" & fld & ")='" & V & "'")
    DeleteRecord = True

Exit Function
err1:
If Err.Number = -2147467259 Then
    MsgBox "Permission Denied!", vbCritical
    Exit Function
Else
    Resume Next
End If
End Function
Sub fillcombo(C As Control, Field, table, adod As ADODB.Connection)
On Error Resume Next
If CStr(Field) = "" Then Exit Sub
Set rs = adod.Execute("select " & Field & " from " & table & " group by " & Field & " order by " & Field)

If Not rs.EOF Then
    C.clear
    While Not rs.EOF
        C.AddItem rs(0)
        rs.MoveNext
    Wend
    C.ListIndex = 0
End If
End Sub
Sub popuplist11(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional COLLAGE As Boolean, Optional colmn As Integer)
    Dim fill As New ADODB.Recordset
    Set fill = New ADODB.Recordset
    popuplist1.vs.Cols = colmn
    fill.Open ST, cn1
    Set popuplist1.vs.DataSource = fill
End Sub
Sub setReportOption()
    frmMain.cr.WindowShowPrintBtn = True
    frmMain.cr.WindowState = crptMaximized
    frmMain.cr.WindowShowPrintBtn = True
    frmMain.cr.WindowShowProgressCtls = True
    frmMain.cr.WindowShowRefreshBtn = True
    frmMain.cr.WindowShowNavigationCtls = True
    frmMain.cr.Action = 1
    
End Sub
Sub popuplist2(ByVal ST As String, ByRef cn1 As ADODB.Connection, Optional ar As Integer, Optional font2 As String)

On Error Resume Next
Dim i As Integer
Dim m As Integer

Set rs1 = New ADODB.Recordset
rs1.Open ST, con

If rs1.RecordCount > 0 Then
         
        If ar = 0 Then
            ar = rs1.Fields.Count
        ElseIf ar > rs1.Fields.Count Then
            ar = rs1.Fields.Count
        End If
        ReDim Array1(ar) As Variant
        For q = 0 To ar - 1
            Array1(q) = 0
        Next q
        'Unload popuplist
        
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        
            For i = 1 To ar
            popuplist.ListView1.ColumnHeaders.Add i, , rs1.Fields(i - 1).Name
            If i = 1 Then
            If rs1.Fields(i - 1).Name = "name" Then
              popuplist.ListView1.ColumnHeaders(i).Width = 10000
            Else
            popuplist.ListView1.ColumnHeaders(i).Width = 2500
            End If
            ElseIf i = 3 Then
               popuplist.ListView1.ColumnHeaders(i).Width = 3000
            'Else
            '   popuplist.ListView1.ColumnHeaders(i).Width = 1400
            End If
            
            
        Next i
        
        
         If font2 = "e" Then
            popuplist.ListView1.Font = "english"
            popuplist.ListView1.Font.Size = 12
        End If

        
        
        
        popuplist.ListView1.View = lvwReport
        Dim litem As ListItem
        While Not rs1.EOF
        If Not IsNull(rs1.Fields(0)) Then
            Set litem = popuplist.ListView1.ListItems.Add(, , rs1.Fields(0).Value)
            If Len(rs1.Fields(0).Value) > Len(rs1.Fields(0).Name) Then
               Array1(0) = Len(rs1.Fields(0).Value)
            Else
               Array1(0) = Len(rs1.Fields(0).Name)
            End If
            If ar > 0 Then
                For m = 1 To ar - 1
                   If Not IsNull(rs1.Fields(m).Value) Then litem.SubItems(m) = rs1.Fields(m).Value
                    If Len(rs1.Fields(m).Value) > Len(rs1.Fields(m).Name) Then
                       Array1(m) = Len(rs1.Fields(m).Value)
                    Else
                       Array1(m) = Len(rs1.Fields(m).Name)
                    End If
                    
                  
                Next m
            End If
            End If
            
            
        
            
            
            
            rs1.MoveNext
        Wend
     
        popuplist.Show 1
End If
rs1.Close
End Sub
Sub companyname()
Dim MYSTRING1 As String
Open App.Path + "\company.mdb" For Input As #1
Line Input #1, MYSTRING1
Close #1
Select Case MYSTRING1
Case "1"
COMPNAME = "SUNIL BOOK BINDING HOUSE"
add1 = "MEERUT"
add2 = "-"
ADD3 = " "
' SAVE THE DATA IN SETUP

Case "2"
COMPNAME = "ILYAS BOOK BINDING WORKS"
add1 = "124,Purwa Afjalur Rahim, Peer Wali Gali,"
add2 = "Near Old Hapur Stand"
ADD3 = "Meerut-250002(U.P.)"

Case "3"
COMPNAME = "M.A. BOOK BINDING"
add1 = "Khasra No. 7, Village Hajipur, Hapur Road"
add2 = "Meerut"
ADD3 = "Ph. : 8439147635,9837887869"



End Select
Set rs = New ADODB.Recordset
   rs.Open "SETUP1", con, adOpenDynamic, adLockOptimistic, adCmdTable
   rs!CNAME = COMPNAME
   rs!add1 = add1
   rs!add2 = add2
   rs!CITY = ADD3

   If Not IsNull(rs!gst) Then
   gst = rs!gst & ""
   End If
   
   If (Len(rs!rem2) = 0 Or IsNull(rs!rem2)) Then
     rs!rem1 = 0
     rs!rem2 = "y"
   End If
   
   rs.Update
rs.Close

End Sub
Sub key_sec()

'''Dim fs As New FileSystemObject
'''Dim txt As TextStream
'''If softwarename = "BHAWNA" Then
'''    Set txt = fs.OpenTextFile(App.Path & "\din.neo", ForReading, False)
'''    ds = DateDiff("d", DateValue(txt.ReadLine), Date)
'''        If (ds > 360) Then

'''           Set txt = Nothing
'''           MsgBox "Data Currepted...", vbCritical
'''        End
'''    End If
'''ElseIf softwarename = "NHPrinting" Then
'''    Set txt = fs.OpenTextFile(App.Path & "\din.neo", ForReading, False)
'''    ds = DateDiff("d", DateValue(txt.ReadLine), Date)
'''    If (ds > 360) Then
'''        Set txt = Nothing
'''        MsgBox "Software Expired.Please contact your Vendor.", vbInformation
'''       End
'''    End If
'''End If



End Sub
Sub ii()
    
    Dim kk As Integer
    
    
    If rs.State = 1 Then rs.Close
    rs.Open "select rem1,court,cst from setup1", con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
    
    fyear = rs!cst & ""
    If rs!COURT = 1 Then
       End
    End If
    
    If rs.Fields(0).Value <> "" Then
       kk = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value)
    Else
       kk = 0
    End If
       
       
    
       If kk = 0 Then
          rs.Fields(0).Value = 1
          rs.Update
        Else
          rs.Fields(0).Value = (Val(rs.Fields(0).Value) + 1)
          rs.Update
       End If
       
       If kk >= 500 Then
          con.Execute "update setup1 set court=1"
          MsgBox "Unable to open data file", vbCritical
          End
       End If
    End If
    
End Sub
