Attribute VB_Name = "Module1"
Sub SaveAndClearInvoice()
    Application.ScreenUpdating = False
    Dim wsInv As Worksheet, wsSales As Worksheet, wsCust As Worksheet, wsMaster As Worksheet
    Dim tblSales As ListObject, tblCust As ListObject
    Dim newRow As ListRow, newCustRow As ListRow
    Dim i As Long, itemSaved As Boolean
    Dim prodID As String, custID As String, custName As String, custMob As String, custAdd As String, custGST As String
    Dim statVal As String, payMode As String
    Dim curInvNo As String, invNum As Long, curDate As String
    Dim qty As Double, rate As Double, gross As Double, discAmt As Double, gstAmt As Double, netAmt As Double
    Dim discPct As Variant, gstPct As Variant
    Dim fMaster As Range
    
    Set wsInv = ThisWorkbook.Sheets("Invoice")
    Set wsSales = ThisWorkbook.Sheets("Sales")
    Set wsCust = ThisWorkbook.Sheets("Customer")

    
    Set tblSales = wsSales.ListObjects(1)
    Set tblCust = wsCust.ListObjects(1)
    
    ' 1. Invoice No aur Date
    curInvNo = Trim(wsInv.Range("M1").MergeArea.Cells(1, 1).Value)
    curDate = Trim(wsInv.Range("M2").MergeArea.Cells(1, 1).Value)
    
    ' 2. Payment Mode aur Status
    payMode = Trim(wsInv.Range("M9").MergeArea.Cells(1, 1).Value)
    If payMode = "" Then payMode = "Cash"
    
    statVal = Trim(wsInv.Range("M11").MergeArea.Cells(1, 1).Value)
    If statVal = "" Then statVal = "Completed"
    
    ' 3. Customer Details
    custID = Trim(wsInv.Range("E9").MergeArea.Cells(1, 1).Value)
    custName = Trim(wsInv.Range("E10").MergeArea.Cells(1, 1).Value)
    custAdd = Trim(wsInv.Range("E11").MergeArea.Cells(1, 1).Value)
    custMob = Trim(wsInv.Range("E12").MergeArea.Cells(1, 1).Value)
    custGST = Trim(wsInv.Range("E13").MergeArea.Cells(1, 1).Value)
    
    ' =========================================================
    ' FIX: Auto-Add Customer (Ab CountIf se 100% save hoga)
    ' =========================================================
    If custID <> "" And custName <> "" Then
        If Application.WorksheetFunction.CountIf(wsCust.Range("A:A"), custID) = 0 Then
            Set newCustRow = tblCust.ListRows.Add
            newCustRow.Range(1, 1).Value = custID
            newCustRow.Range(1, 2).Value = custName
            newCustRow.Range(1, 3).Value = custMob
            newCustRow.Range(1, 4).Value = custAdd
            newCustRow.Range(1, 5).Value = custGST
        End If
    End If
    ' =========================================================
    
    itemSaved = False
    
    ' 4. Item Table Save
    For i = 16 To 31
        prodID = Trim(wsInv.Cells(i, "B").MergeArea.Cells(1, 1).Value)
        qty = Val(wsInv.Cells(i, "F").MergeArea.Cells(1, 1).Value)
        
        If prodID <> "" And qty > 0 Then
            itemSaved = True
            rate = Val(wsInv.Cells(i, "G").MergeArea.Cells(1, 1).Value)
            gross = Val(wsInv.Cells(i, "H").MergeArea.Cells(1, 1).Value)
            discAmt = Val(wsInv.Cells(i, "J").MergeArea.Cells(1, 1).Value)
            gstAmt = Val(wsInv.Cells(i, "M").MergeArea.Cells(1, 1).Value)
            netAmt = Val(wsInv.Cells(i, "N").MergeArea.Cells(1, 1).Value)
            
            discPct = wsInv.Cells(i, "I").MergeArea.Cells(1, 1).Value
            If Not IsNumeric(discPct) Or IsEmpty(discPct) Then discPct = 0
            
            gstPct = wsInv.Cells(i, "L").MergeArea.Cells(1, 1).Value
            If Not IsNumeric(gstPct) Or IsEmpty(gstPct) Then gstPct = 0
            
            Set newRow = tblSales.ListRows.Add
            newRow.Range(1, 1).Value = curInvNo
            newRow.Range(1, 2).Value = curDate
            newRow.Range(1, 3).Value = custID
            newRow.Range(1, 4).Value = custName
            newRow.Range(1, 5).Value = prodID
            
            
            If Not fMaster Is Nothing Then
                newRow.Range(1, 6).Value = fMaster.Offset(0, 1).Value
            Else
                newRow.Range(1, 6).Value = wsInv.Cells(i, "C").MergeArea.Cells(1, 1).Value
            End If
            
            newRow.Range(1, 8).Value = rate
            newRow.Range(1, 10).Value = discPct
            newRow.Range(1, 12).Value = gstPct
            newRow.Range(1, 15).Value = payMode
            newRow.Range(1, 16).Value = statVal
            
            If UCase(statVal) = "RETURN" Then
                newRow.Range(1, 7).Value = -qty
                newRow.Range(1, 9).Value = -gross
                newRow.Range(1, 11).Value = -discAmt
                newRow.Range(1, 13).Value = -gstAmt
                newRow.Range(1, 14).Value = -netAmt
            Else
                newRow.Range(1, 7).Value = qty
                newRow.Range(1, 9).Value = gross
                newRow.Range(1, 11).Value = discAmt
                newRow.Range(1, 13).Value = gstAmt
                newRow.Range(1, 14).Value = netAmt
            End If
        End If
    Next i
    
    If Not itemSaved Then
        MsgBox "Kripya Product ID aur QTY darj karein!", vbExclamation
        Exit Sub
    End If
    
    ' 5. Invoice No & Cust ID Auto-Increment
    invNum = Val(Replace(Replace(UCase(curInvNo), "INV", ""), "-", ""))
    If invNum > 0 Then wsInv.Range("M1").MergeArea.Cells(1, 1).Value = "INV-" & (invNum + 1)
    
    If UCase(custID) Like "CUST*" Then
        Dim custNum As Long
        custNum = Val(Replace(UCase(custID), "CUST", ""))
        If custNum > 0 Then wsInv.Range("E9").MergeArea.Cells(1, 1).Value = "CUST" & Format(custNum + 1, "0000")
    End If
    
    ' 6. Direct Cell Clear
    wsInv.Range("E10").MergeArea.Cells(1, 1).Value = ""
    wsInv.Range("E11").MergeArea.Cells(1, 1).Value = ""
    wsInv.Range("E12").MergeArea.Cells(1, 1).Value = ""
    wsInv.Range("E13").MergeArea.Cells(1, 1).Value = ""
    
    wsInv.Range("M9").MergeArea.Cells(1, 1).Value = ""
    wsInv.Range("M11").MergeArea.Cells(1, 1).Value = ""
    
    For i = 16 To 31
        If Not wsInv.Cells(i, "B").MergeArea.Cells(1, 1).HasFormula Then wsInv.Cells(i, "B").MergeArea.Cells(1, 1).Value = ""
        If Not wsInv.Cells(i, "F").MergeArea.Cells(1, 1).HasFormula Then wsInv.Cells(i, "F").MergeArea.Cells(1, 1).Value = ""
        If Not wsInv.Cells(i, "I").MergeArea.Cells(1, 1).HasFormula Then wsInv.Cells(i, "I").MergeArea.Cells(1, 1).Value = ""
        If Not wsInv.Cells(i, "L").MergeArea.Cells(1, 1).HasFormula Then wsInv.Cells(i, "L").MergeArea.Cells(1, 1).Value = ""
    Next i
    
    MsgBox "Tax Invoice Table aur Customer dono save ho gaye!", vbInformation, "Saved"
    Application.ScreenUpdating = True
End Sub
Function RupeesInWords(ByVal MyNumber As Variant) As String
    Dim Rupees As String, Paise As String
    Dim DecimalPlace As Long
    Dim NumStr As String
    
    If Not IsNumeric(MyNumber) Or Val(MyNumber) = 0 Then
        RupeesInWords = ""
        Exit Function
    End If
    
    NumStr = Trim(Format(MyNumber, "0.00"))
    DecimalPlace = InStr(NumStr, ".")
    
    If DecimalPlace > 0 Then
        Paise = ConvertThreeDigits(Val(Mid(NumStr, DecimalPlace + 1, 2)))
        NumStr = Left(NumStr, DecimalPlace - 1)
    End If
    
    Dim n As Double
    n = Val(NumStr)
    
    If n > 0 Then
        Rupees = ConvertIndianNumber(n) & " Rupees"
    End If
    
    If Paise <> "" Then
        If Rupees <> "" Then
            RupeesInWords = "Rupees " & ConvertIndianNumber(n) & " And " & Paise & " Paise Only"
        Else
            RupeesInWords = Paise & " Paise Only"
        End If
    Else
        If Rupees <> "" Then
            RupeesInWords = "Rupees " & ConvertIndianNumber(n) & " Only"
        End If
    End If
End Function

Private Function ConvertIndianNumber(ByVal n As Double) As String
    Dim Crores As Double, Lakhs As Double, Thousands As Double, Remainder As Double
    Dim Res As String
    
    Crores = Int(n / 10000000)
    n = n - Crores * 10000000
    
    Lakhs = Int(n / 100000)
    n = n - Lakhs * 100000
    
    Thousands = Int(n / 1000)
    Remainder = n - Thousands * 1000
    
    If Crores > 0 Then Res = Res & ConvertIndianNumber(Crores) & " Crore "
    If Lakhs > 0 Then Res = Res & ConvertTwoDigits(Lakhs) & " Lakh "
    If Thousands > 0 Then Res = Res & ConvertTwoDigits(Thousands) & " Thousand "
    If Remainder > 0 Then Res = Res & ConvertThreeDigits(Remainder)
    
    ConvertIndianNumber = Application.WorksheetFunction.Trim(Res)
End Function

Private Function ConvertThreeDigits(ByVal n As Double) As String
    Dim Hundreds As Long, Rest As Long
    Dim Res As String
    Hundreds = Int(n / 100)
    Rest = n Mod 100
    If Hundreds > 0 Then Res = ConvertTwoDigits(Hundreds) & " Hundred "
    If Rest > 0 Then Res = Res & ConvertTwoDigits(Rest)
    ConvertThreeDigits = Application.WorksheetFunction.Trim(Res)
End Function

Private Function ConvertTwoDigits(ByVal n As Long) As String
    Dim Units As Variant, Tens As Variant
    Units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", _
                  "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", _
                  "Seventeen", "Eighteen", "Nineteen")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    
    If n < 20 Then
        ConvertTwoDigits = Units(n)
    Else
        ConvertTwoDigits = Application.WorksheetFunction.Trim(Tens(Int(n / 10)) & " " & Units(n Mod 10))
    End If
End Function
Sub PrintInvoice()
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets("Invoice")
    
    ' 1. Sirf Invoice wale hisse (A1 se O44) ko Print Area set karein (Buttons print nahi honge)
    wsInv.PageSetup.PrintArea = "$A$1:$O$44"
    
    ' 2. Invoice ko 1 Page par bilkul Center aur Fit set karein
    With wsInv.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
    
    ' 3. Print Preview dikhayein (Taaki aap print ya PDF save karne se pehle dekh sakein)
    wsInv.PrintPreview
    
    ' Note: Agar aap bina Preview dekhe seedha Printer se print nikalna chahte hain,
    ' to upar wali 'wsInv.PrintPreview' line ke aage apostrophe (') laga dein
    ' aur niche wali line se apostrophe (') hata dein:
    
    ' wsInv.PrintOut Copies:=1
End Sub
Sub SaveAndClearPurchase()
    Dim wsPurBill As Worksheet, wsPur As Worksheet, wsSup As Worksheet, wsMaster As Worksheet
    Dim tblPur As ListObject, tblSup As ListObject
    Dim newRow As ListRow, newSupRow As ListRow
    Dim i As Long, itemSaved As Boolean
    Dim prodID As String, supID As String, supName As String, supMob As String, supAdd As String, supState As String, supGST As String
    Dim statVal As String, payMode As String
    Dim curPurNo As String, purNum As Long, curDate As String
    Dim qty As Double, rate As Double, gross As Double, discAmt As Double, gstAmt As Double, netAmt As Double
    Dim discPct As Variant, gstPct As Variant
    
    Set wsPurBill = ThisWorkbook.Sheets("Purchase Bill")
    Set wsPur = ThisWorkbook.Sheets("Purchase")
    Set wsSup = ThisWorkbook.Sheets("Supplier")
    
    
    Set tblPur = wsPur.ListObjects(1)
    Set tblSup = wsSup.ListObjects(1)
    
    ' 1. Purchase No aur Date (Aapke bataye M3 aur M4 se)
    curPurNo = Trim(wsPurBill.Range("M3").MergeArea.Cells(1, 1).Value)
    curDate = Trim(wsPurBill.Range("M4").MergeArea.Cells(1, 1).Value)
    
    ' 2. Payment Mode aur Status (M9 aur M11 se)
    payMode = Trim(wsPurBill.Range("M9").MergeArea.Cells(1, 1).Value)
    If payMode = "" Then payMode = "Cash"
    
    statVal = Trim(wsPurBill.Range("M11").MergeArea.Cells(1, 1).Value)
    If statVal = "" Then statVal = "Received"
    
    ' 3. Supplier Details (E9 se E14)
    supID = Trim(wsPurBill.Range("E9").MergeArea.Cells(1, 1).Value)
    supName = Trim(wsPurBill.Range("E10").MergeArea.Cells(1, 1).Value)
    supMob = Trim(wsPurBill.Range("E11").MergeArea.Cells(1, 1).Value)
    supAdd = Trim(wsPurBill.Range("E12").MergeArea.Cells(1, 1).Value)
    supState = Trim(wsPurBill.Range("E13").MergeArea.Cells(1, 1).Value)
    supGST = Trim(wsPurBill.Range("E14").MergeArea.Cells(1, 1).Value)
    
    ' =========================================================
    ' FIX: Auto-Add Supplier (100% Save hoga)
    ' =========================================================
    If supID <> "" And supName <> "" Then
        If Application.WorksheetFunction.CountIf(wsSup.Range("A:A"), supID) = 0 Then
            Set newSupRow = tblSup.ListRows.Add
            newSupRow.Range(1, 1).Value = supID
            newSupRow.Range(1, 2).Value = supName
            newSupRow.Range(1, 3).Value = supMob
            newSupRow.Range(1, 4).Value = supAdd
            newSupRow.Range(1, 5).Value = supState
            newSupRow.Range(1, 6).Value = supGST
        End If
    End If
    ' =========================================================
    
    itemSaved = False
    
    ' 4. Item Table Save (Row 16 se 31)
    For i = 16 To 31
        prodID = Trim(wsPurBill.Cells(i, "B").MergeArea.Cells(1, 1).Value)
        qty = Val(wsPurBill.Cells(i, "F").MergeArea.Cells(1, 1).Value) ' QTY -> F
        
        If prodID <> "" And qty > 0 Then
            itemSaved = True
            
            rate = Val(wsPurBill.Cells(i, "G").MergeArea.Cells(1, 1).Value)
            gross = Val(wsPurBill.Cells(i, "H").MergeArea.Cells(1, 1).Value)
            discAmt = Val(wsPurBill.Cells(i, "J").MergeArea.Cells(1, 1).Value)
            gstAmt = Val(wsPurBill.Cells(i, "M").MergeArea.Cells(1, 1).Value)
            netAmt = Val(wsPurBill.Cells(i, "N").MergeArea.Cells(1, 1).Value)
            
            discPct = wsPurBill.Cells(i, "I").MergeArea.Cells(1, 1).Value
            If Not IsNumeric(discPct) Or IsEmpty(discPct) Then discPct = 0
            
            gstPct = wsPurBill.Cells(i, "L").MergeArea.Cells(1, 1).Value
            If Not IsNumeric(gstPct) Or IsEmpty(gstPct) Then gstPct = 0
            
            Set newRow = tblPur.ListRows.Add
            newRow.Range(1, 1).Value = curPurNo
            newRow.Range(1, 2).Value = curDate
            newRow.Range(1, 3).Value = supID
            newRow.Range(1, 4).Value = supName
            newRow.Range(1, 5).Value = supMob
            newRow.Range(1, 6).Value = supAdd
            newRow.Range(1, 7).Value = supState
            newRow.Range(1, 8).Value = supGST
            newRow.Range(1, 9).Value = prodID
            
            newRow.Range(1, 11).Value = rate
            newRow.Range(1, 13).Value = discPct
            newRow.Range(1, 15).Value = gstPct
            newRow.Range(1, 18).Value = statVal
            
            If UCase(statVal) = "RETURN" Then
                newRow.Range(1, 10).Value = -qty
                newRow.Range(1, 12).Value = -gross
                newRow.Range(1, 14).Value = -discAmt
                newRow.Range(1, 16).Value = -gstAmt
                newRow.Range(1, 17).Value = -netAmt
            Else
                newRow.Range(1, 10).Value = qty
                newRow.Range(1, 12).Value = gross
                newRow.Range(1, 14).Value = discAmt
                newRow.Range(1, 16).Value = gstAmt
                newRow.Range(1, 17).Value = netAmt
            End If
        End If
    Next i
    
    If Not itemSaved Then
        MsgBox "Kripya Product ID aur QTY darj karein!", vbExclamation
        Exit Sub
    End If
    
    ' 5. Auto-Increment Purchase No (Sirf Purchase No badhega)
    purNum = Val(Replace(Replace(UCase(curPurNo), "PUR-", ""), "PUR", ""))
    If purNum > 0 Then wsPurBill.Range("M3").MergeArea.Cells(1, 1).Value = "PUR-" & (purNum + 1)
    
    ' 6. Direct Clear Form (Supplier ID E9 ab MIT JAYEGA)
     wsPurBill.Range("E9").MergeArea.Cells(1, 1).Value = ""  ' Supplier ID Clear
    'wsPurBill.Range("E10").MergeArea.Cells(1, 1).Value = "" ' Name Clear
    'wsPurBill.Range("E11").MergeArea.Cells(1, 1).Value = "" ' Contact Clear
    'wsPurBill.Range("E12").MergeArea.Cells(1, 1).Value = "" ' Address Clear
    'wsPurBill.Range("E13").MergeArea.Cells(1, 1).Value = "" ' State Clear
    'wsPurBill.Range("E14").MergeArea.Cells(1, 1).Value = "" ' GSTIN Clear
    
    wsPurBill.Range("M9").MergeArea.Cells(1, 1).Value = ""
    wsPurBill.Range("M11").MergeArea.Cells(1, 1).Value = ""
    
    For i = 16 To 31
        If Not wsPurBill.Cells(i, "B").MergeArea.Cells(1, 1).HasFormula Then wsPurBill.Cells(i, "B").MergeArea.Cells(1, 1).Value = ""
        If Not wsPurBill.Cells(i, "F").MergeArea.Cells(1, 1).HasFormula Then wsPurBill.Cells(i, "F").MergeArea.Cells(1, 1).Value = ""
        If Not wsPurBill.Cells(i, "I").MergeArea.Cells(1, 1).HasFormula Then wsPurBill.Cells(i, "I").MergeArea.Cells(1, 1).Value = ""
        If Not wsPurBill.Cells(i, "L").MergeArea.Cells(1, 1).HasFormula Then wsPurBill.Cells(i, "L").MergeArea.Cells(1, 1).Value = ""
    Next i
    
    MsgBox "Purchase Bill aur Supplier dono save ho gaye!", vbInformation, "Saved"
End Sub
' =========================================================
' 1. SEARCH INVOICE CODE (Bill dhundhne ke liye)
' =========================================================
Sub SearchInvoice()
    Dim wsRep As Worksheet, wsSales As Worksheet, wsCust As Worksheet
    Dim tblSales As ListObject, tblCust As ListObject
    Dim invNo As String, r As Long, rowRep As Long
    Dim found As Boolean
    Dim custID As String, fCust As Range
    
    ' Apni sheets ke naam set karein
    Set wsRep = ThisWorkbook.Sheets("Reprint INVS")
    Set wsSales = ThisWorkbook.Sheets("Sales")
    Set wsCust = ThisWorkbook.Sheets("Customer")
    
    Set tblSales = wsSales.ListObjects(1)
    Set tblCust = wsCust.ListObjects(1)
    
    ' Search box (M1) se Invoice Number lena
    invNo = Trim(wsRep.Range("M1").MergeArea.Cells(1, 1).Value)
    
    If invNo = "" Then
        MsgBox "Kripya search karne ke liye Invoice Number (M1 me) darj karein!", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    ' 1. Purana data clear karna (Taaki naya data saaf dikhe)
    wsRep.Range("M2").MergeArea.Cells(1, 1).Value = "" ' Date
    wsRep.Range("E9").MergeArea.Cells(1, 1).Value = "" ' Cust ID
    wsRep.Range("E10").MergeArea.Cells(1, 1).Value = "" ' Name
    wsRep.Range("E11").MergeArea.Cells(1, 1).Value = "" ' Address
    wsRep.Range("E12").MergeArea.Cells(1, 1).Value = "" ' Mobile
    wsRep.Range("E13").MergeArea.Cells(1, 1).Value = "" ' GSTIN
    wsRep.Range("M9").MergeArea.Cells(1, 1).Value = "" ' Pay Mode
    wsRep.Range("M11").MergeArea.Cells(1, 1).Value = "" ' Status
    
    For r = 16 To 31
        wsRep.Cells(r, "B").MergeArea.Cells(1, 1).Value = ""
        wsRep.Cells(r, "C").MergeArea.Cells(1, 1).Value = ""
        wsRep.Cells(r, "F").MergeArea.Cells(1, 1).Value = ""
        wsRep.Cells(r, "G").MergeArea.Cells(1, 1).Value = ""
        wsRep.Cells(r, "I").MergeArea.Cells(1, 1).Value = ""
        wsRep.Cells(r, "L").MergeArea.Cells(1, 1).Value = ""
    Next r
    
    found = False
    rowRep = 16
    
    ' 2. Sales Table me Invoice dhundhna
    For r = 1 To tblSales.ListRows.Count
        ' Agar Invoice No match karta hai
        If UCase(Trim(tblSales.DataBodyRange(r, 1).Value)) = UCase(invNo) Then
            
            ' Pehli bar match hone par Customer Details bharna
            If Not found Then
                wsRep.Range("M2").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 2).Value ' Date
                custID = Trim(tblSales.DataBodyRange(r, 3).Value)
                wsRep.Range("E9").MergeArea.Cells(1, 1).Value = custID
                wsRep.Range("E10").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 4).Value ' Name
                wsRep.Range("M9").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 15).Value ' Pay Mode
                wsRep.Range("M11").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 16).Value ' Status
                
                ' Customer Table se Mobile, Address, GSTIN nikalna
                If custID <> "" Then
                    Set fCust = wsCust.Columns(tblCust.ListColumns(1).Range.Column).Find(What:=custID, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not fCust Is Nothing Then
                        wsRep.Range("E12").MergeArea.Cells(1, 1).Value = fCust.Offset(0, 2).Value ' Mobile
                        wsRep.Range("E11").MergeArea.Cells(1, 1).Value = fCust.Offset(0, 3).Value ' Address
                        wsRep.Range("E13").MergeArea.Cells(1, 1).Value = fCust.Offset(0, 4).Value ' GSTIN
                    End If
                End If
                found = True
            End If
            
            ' Item Details bharna (Row 16 se aage)
            If rowRep <= 31 Then
                wsRep.Cells(rowRep, "B").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 5).Value ' Prod ID
                wsRep.Cells(rowRep, "C").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 6).Value ' Prod Name
                wsRep.Cells(rowRep, "F").MergeArea.Cells(1, 1).Value = Abs(tblSales.DataBodyRange(r, 7).Value) ' Qty (Abs = minus hatane ke liye)
                wsRep.Cells(rowRep, "G").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 8).Value ' Rate
                wsRep.Cells(rowRep, "I").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 10).Value ' Disc %
                wsRep.Cells(rowRep, "L").MergeArea.Cells(1, 1).Value = tblSales.DataBodyRange(r, 12).Value ' GST %
                
                rowRep = rowRep + 1
            End If
        End If
    Next r
    
    If Not found Then
        MsgBox "Ye Invoice Number (" & invNo & ") Sales record me nahi mila!", vbCritical, "Not Found"
    Else
        MsgBox "Invoice ki puri detail aa gayi hai. Ab aap ise Print kar sakte hain!", vbInformation, "Success"
    End If
End Sub

' =========================================================
' 2. PRINT INVOICE CODE (Print karne ke liye)
' =========================================================
Sub PrintReprintInvoice()
    ' Ye code pehle Print Preview dikhayega taaki page waste na ho.
    ' Print preview me upar 'Print' ka option aata hai.
    ThisWorkbook.Sheets("Reprint INVS").PrintPreview
    
    ' Note: Agar aap seedha printer se paper nikalna chahte hain bina preview dekhe,
    ' to upar wali line delete karke niche wali line likh lein:
    ' ThisWorkbook.Sheets("Reprint INVS").PrintOut Copies:=1
End Sub
' =========================================================
' 3. SEARCH PURCHASE BILL (Purchase Bill dhundhne ke liye)
' =========================================================
Sub SearchPurchaseBill()
    Dim wsRepPur As Worksheet, wsPur As Worksheet, wsSup As Worksheet, wsMaster As Worksheet
    Dim tblPur As ListObject, tblSup As ListObject
    Dim purNo As String, r As Long, rowRep As Long
    Dim found As Boolean
    Dim supID As String, prodID As String
    Dim fSup As Range, fMaster As Range
    
    Set wsRepPur = ThisWorkbook.Sheets("Reprint Purh")
    Set wsPur = ThisWorkbook.Sheets("Purchase")
    Set wsSup = ThisWorkbook.Sheets("Supplier")
    
    
    Set tblPur = wsPur.ListObjects(1)
    Set tblSup = wsSup.ListObjects(1)
    
    ' Search box (M3 se Purchase Number lena - M aur N merge hain)
    purNo = Trim(wsRepPur.Range("M3").MergeArea.Cells(1, 1).Value)
    
    If purNo = "" Then
        MsgBox "Kripya search karne ke liye Purchase Bill Number darj karein!", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    ' 1. Purana data clear karna
    wsRepPur.Range("M4").MergeArea.Cells(1, 1).Value = "" ' Date
    wsRepPur.Range("E9").MergeArea.Cells(1, 1).Value = "" ' Supplier ID
    wsRepPur.Range("E10").MergeArea.Cells(1, 1).Value = "" ' Name
    wsRepPur.Range("E11").MergeArea.Cells(1, 1).Value = "" ' Contact
    wsRepPur.Range("E12").MergeArea.Cells(1, 1).Value = "" ' Address
    wsRepPur.Range("E13").MergeArea.Cells(1, 1).Value = "" ' State
    wsRepPur.Range("E14").MergeArea.Cells(1, 1).Value = "" ' GSTIN
    
    ' Pay Mode aur Status (LMN merge hain, isliye L par target karenge)
    wsRepPur.Range("L9").MergeArea.Cells(1, 1).Value = ""
    wsRepPur.Range("L11").MergeArea.Cells(1, 1).Value = ""
    
    For r = 16 To 31
        wsRepPur.Cells(r, "B").Value = ""
        wsRepPur.Cells(r, "C").Value = ""
        wsRepPur.Cells(r, "F").Value = ""
        wsRepPur.Cells(r, "G").Value = ""
        wsRepPur.Cells(r, "H").Value = ""
        wsRepPur.Cells(r, "I").Value = ""
        wsRepPur.Cells(r, "J").Value = ""
        wsRepPur.Cells(r, "L").Value = ""
        wsRepPur.Cells(r, "M").Value = ""
        wsRepPur.Cells(r, "N").Value = ""
    Next r
    
    found = False
    rowRep = 16
    
    ' 2. Purchase Table me Bill dhundhna
    For r = 1 To tblPur.ListRows.Count
        ' Agar Purchase No match karta hai
        If UCase(Trim(tblPur.DataBodyRange(r, 1).Value)) = UCase(purNo) Then
            
            ' Pehli bar match hone par Supplier Details bharna
            If Not found Then
                wsRepPur.Range("M4").MergeArea.Cells(1, 1).Value = tblPur.DataBodyRange(r, 2).Value ' Date
                supID = Trim(tblPur.DataBodyRange(r, 3).Value)
                wsRepPur.Range("E9").MergeArea.Cells(1, 1).Value = supID
                wsRepPur.Range("E10").MergeArea.Cells(1, 1).Value = tblPur.DataBodyRange(r, 4).Value ' Name
                
                ' FIX: Status ab L11 me 100% aayega (Purchase Table ke 18th column se)
                wsRepPur.Range("L11").MergeArea.Cells(1, 1).Value = tblPur.DataBodyRange(r, 18).Value
                
                ' Payment Mode (Agar jarurat ho to)
                wsRepPur.Range("L9").MergeArea.Cells(1, 1).Value = "Cash"
                
                ' Supplier Table se baaki details nikalna
                If supID <> "" Then
                    Set fSup = wsSup.Columns(tblSup.ListColumns(1).Range.Column).Find(What:=supID, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not fSup Is Nothing Then
                        wsRepPur.Range("E11").MergeArea.Cells(1, 1).Value = fSup.Offset(0, 2).Value ' Contact
                        wsRepPur.Range("E12").MergeArea.Cells(1, 1).Value = fSup.Offset(0, 3).Value ' Address
                        wsRepPur.Range("E13").MergeArea.Cells(1, 1).Value = fSup.Offset(0, 4).Value ' State
                        wsRepPur.Range("E14").MergeArea.Cells(1, 1).Value = fSup.Offset(0, 5).Value ' GSTIN
                    End If
                End If
                found = True
            End If
            
            ' Item Details bharna (Row 16 se aage)
            If rowRep <= 31 Then
                prodID = Trim(tblPur.DataBodyRange(r, 9).Value)
                wsRepPur.Cells(rowRep, "B").Value = prodID ' Prod ID
                
                ' Master Sheet se Product Name nikalna
                
                If Not fMaster Is Nothing Then
                    wsRepPur.Cells(rowRep, "C").Value = fMaster.Offset(0, 1).Value
                Else
                    wsRepPur.Cells(rowRep, "C").Value = "Name Not Found"
                End If
                
                wsRepPur.Cells(rowRep, "F").Value = Abs(tblPur.DataBodyRange(r, 10).Value) ' Qty
                wsRepPur.Cells(rowRep, "G").Value = tblPur.DataBodyRange(r, 11).Value ' Rate
                wsRepPur.Cells(rowRep, "H").Value = Abs(tblPur.DataBodyRange(r, 12).Value) ' Gross
                wsRepPur.Cells(rowRep, "I").Value = tblPur.DataBodyRange(r, 13).Value ' Disc %
                wsRepPur.Cells(rowRep, "J").Value = Abs(tblPur.DataBodyRange(r, 14).Value) ' Disc Amt
                wsRepPur.Cells(rowRep, "L").Value = tblPur.DataBodyRange(r, 15).Value ' GST %
                wsRepPur.Cells(rowRep, "M").Value = Abs(tblPur.DataBodyRange(r, 16).Value) ' GST Amt
                wsRepPur.Cells(rowRep, "N").Value = Abs(tblPur.DataBodyRange(r, 17).Value) ' Net Amt
                
                rowRep = rowRep + 1
            End If
        End If
    Next r
    
    If Not found Then
        MsgBox "Ye Purchase Bill Number (" & purNo & ") Purchase record me nahi mila!", vbCritical, "Not Found"
    Else
        MsgBox "Purchase Bill ki puri detail aa gayi hai!", vbInformation, "Success"
    End If
End Sub

' =========================================================
' 4. PRINT PURCHASE BILL CODE (Yahan se print preview khulega)
' =========================================================
Sub PrintPurchaseBill()
    ThisWorkbook.Sheets("Reprint Purh").PrintPreview
End Sub
