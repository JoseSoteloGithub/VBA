Attribute VB_Name = "modDeviations"
Option Explicit

Sub subOpenCurrentFilePath()

    ActiveWorkbook.FollowHyperlink ActiveWorkbook.Path

End Sub

Sub subSaveFolderAndFileBasedOnCompliance()

Dim fileToSaveName As String
Dim masterCmpnyVndrName As String
Dim masterCmpnyVndrNameColumnLetter As String
Dim masterCmpnyVndrNamePosition As Range
Dim programMasterName As String
Dim programMasterNameColumnLetter As String
Dim programMasterNamePosition As Range
Dim wkb As Workbook

Dim testIfFileSaved As String
Dim testIfFileExists As String

fileToSaveName = ""
masterCmpnyVndrName = ""
masterCmpnyVndrNameColumnLetter = ""
Set masterCmpnyVndrNamePosition = Nothing
programMasterName = ""
programMasterNameColumnLetter = ""
Set programMasterNamePosition = Nothing
Set wkb = Nothing

    If Range("A1") = "Run Date" Then
    
    ElseIf Range("A1") = "Run DT" Then
        
    Else
    
        MsgBox ("""Run Date"" or ""Run DT"" were not found in range A1.  Exiting")
        
        Exit Sub
    
    End If

    Set wkb = ActiveWorkbook
    
    Set masterCmpnyVndrNamePosition = Range("A1:IV1").Find("Master Cmpny Vndr Name", lookat:=xlWhole)
    
    If masterCmpnyVndrNamePosition Is Nothing Then
    
        Set masterCmpnyVndrNamePosition = Range("A1:IV1").Find("Mstr Cmpny Vndr", lookat:=xlWhole)
    
    End If
    
    masterCmpnyVndrNameColumnLetter = Evaluate("MID(""" & masterCmpnyVndrNamePosition.Address & """,FIND(""$"",""" & masterCmpnyVndrNamePosition.Address & """)+1,FIND(""$"",""" & masterCmpnyVndrNamePosition.Address & """,2)-2)")
    
    masterCmpnyVndrName = Range(masterCmpnyVndrNameColumnLetter & 2)
    
    If InStr(masterCmpnyVndrName, "'") > 0 Then
    
        masterCmpnyVndrName = Replace(masterCmpnyVndrName, "'", " ")
    
    End If
    
    If InStr(masterCmpnyVndrName, ".") > 0 Then
    
        masterCmpnyVndrName = Replace(masterCmpnyVndrName, ".", "")
    
    End If
    
    If InStr(masterCmpnyVndrName, "_") > 0 Then
    
        masterCmpnyVndrName = Replace(masterCmpnyVndrName, "_", "")
    
    End If
    
    If InStr(masterCmpnyVndrName, "(") > 0 Then
    
        masterCmpnyVndrName = Replace(masterCmpnyVndrName, "(", "")
    
    End If
    
    If InStr(masterCmpnyVndrName, ")") > 0 Then
    
        masterCmpnyVndrName = Replace(masterCmpnyVndrName, ")", "")
    
    End If
    
    If InStr(masterCmpnyVndrName, "/") > 0 Then
    
        masterCmpnyVndrName = Replace(masterCmpnyVndrName, "/", " ")
    
    End If
    
    If InStr(masterCmpnyVndrName, "\") > 0 Then
    
        masterCmpnyVndrName = Replace(masterCmpnyVndrName, "\", " ")
    
    End If
    
    If InStr(masterCmpnyVndrName, ",") > 0 Then
    
        masterCmpnyVndrName = Replace(masterCmpnyVndrName, ",", "")
    
    End If
    
    Set programMasterNamePosition = Range("A1:IV1").Find("Program Master Name", lookat:=xlWhole)
    
    If programMasterNamePosition Is Nothing Then
    
        Set programMasterNamePosition = Range("A1:IV1").Find("Pgm Mstr Name", lookat:=xlWhole)
    
    End If
        
    programMasterNameColumnLetter = Evaluate("MID(""" & programMasterNamePosition.Address & """,FIND(""$"",""" & programMasterNamePosition.Address & """)+1,FIND(""$"",""" & programMasterNamePosition.Address & """,2)-2)")
    
    programMasterName = Range(programMasterNameColumnLetter & 2)
    
    If InStr(programMasterName, "'") > 0 Then
    
        programMasterName = Replace(programMasterName, "'", " ")
    
    End If
    
    If InStr(programMasterName, ".") > 0 Then
    
        programMasterName = Replace(programMasterName, ".", " ")
    
    End If
    
    If InStr(programMasterName, "_") > 0 Then
    
        programMasterName = Replace(programMasterName, "_", " ")
    
    End If
    
    If InStr(programMasterName, "(") > 0 Then
    
        programMasterName = Replace(programMasterName, "(", "")
    
    End If
    
    If InStr(programMasterName, ")") > 0 Then
    
        programMasterName = Replace(programMasterName, ")", "")
    
    End If
    
    If InStr(programMasterName, "/") > 0 Then
    
        programMasterName = Replace(programMasterName, "/", " ")
    
    End If
    
    If InStr(programMasterName, "\") > 0 Then
    
        programMasterName = Replace(programMasterName, "\", " ")
    
    End If
    
    If InStr(programMasterName, ",") > 0 Then
    
        programMasterName = Replace(programMasterName, ",", "")
    
    End If
    
    fileToSaveName = masterCmpnyVndrName & " and " & programMasterName
    
    If Len(fileToSaveName) >= 46 Then
    
        fileToSaveName = Left(fileToSaveName, 45)
    
    End If
    
    Do Until Right(fileToSaveName, 1) <> " "
    
        If Right(fileToSaveName, 1) = " " Then
    
            fileToSaveName = Left(fileToSaveName, Len(fileToSaveName) - 1)
    
        End If
    
    Loop
    
    'G:\National SIS\Shared\Remedy\Remedy Research

    On Error Resume Next

    wkb.SaveAs "G:\National SIS\Shared\Remedy\Remedy Research\" & fileToSaveName & "\" & fileToSaveName & ".xlsx"

    testIfFileExists = Dir("G:\National SIS\Shared\Remedy\Remedy Research\" & fileToSaveName & "\" & fileToSaveName & ".xlsx")
    
    If testIfFileExists = "" Then
    
    Else
    
        ActiveWorkbook.FollowHyperlink "G:\National SIS\Shared\Remedy\Remedy Research\" & fileToSaveName
    
    End If

    If Err.Number = 1004 Then
    
        MkDir "G:\National SIS\Shared\Remedy\Remedy Research\" & fileToSaveName
    
        wkb.SaveAs "G:\National SIS\Shared\Remedy\Remedy Research\" & fileToSaveName & "\" & fileToSaveName & ".xlsx"
    
        Err.Clear
    
    End If
    
    testIfFileSaved = Dir("G:\National SIS\Shared\Remedy\Remedy Research\" & fileToSaveName & "\" & fileToSaveName & ".xlsx")
    
    If testIfFileSaved = "" Then
    
        ActiveWorkbook.FollowHyperlink "G:\National SIS\Shared\Remedy\Remedy Research\" & fileToSaveName
    
'        sheetToMove = ActiveSheet
'
'        ActiveSheet.Copy
'
'        sheetToMove.Move After:=Workbooks("Book21").Sheets(1)
    
        MsgBox ("File did not save properly")
    
        Set wkb = ActiveWorkbook
    
        wkb.SaveAs "G:\National SIS\Shared\Remedy\Remedy Research\" & fileToSaveName & "\" & fileToSaveName & ".xlsx"
    
    End If
    
fileToSaveName = ""
masterCmpnyVndrName = ""
masterCmpnyVndrNameColumnLetter = ""
Set masterCmpnyVndrNamePosition = Nothing
programMasterName = ""
programMasterNameColumnLetter = ""
Set programMasterNamePosition = Nothing
Set wkb = Nothing

End Sub

Sub subMerlinGetProdDivFrmCmplnc()

Dim beginingSheet As Worksheet
Dim divisionNumberLastRow As Long
Dim divNumberColumnLetter As String
Dim divNumberPosition As Range
Dim iE As Object
Dim iEFound As Boolean
Dim lastRow As Long
Dim lookingForFolder As Object
Dim lookingForFolderFound As Boolean
Dim newSheet As Worksheet
Dim productNumberlastRow As Long
Dim responseFromMessageBox As Long
Dim uSFProdNumberColumnLetter As String
Dim uSFProdNumberPosition As Range

Dim discovererLink As Object

Set beginingSheet = Nothing
divisionNumberLastRow = Empty
divNumberColumnLetter = ""
Set divNumberPosition = Nothing
Set iE = Nothing
iEFound = Empty
lastRow = Empty
Set lookingForFolder = Nothing
lookingForFolderFound = Empty
Set newSheet = Nothing
productNumberlastRow = Empty
responseFromMessageBox = Empty
uSFProdNumberColumnLetter = ""
Set uSFProdNumberPosition = Nothing

    If Range("A1") = "Run Date" Then
    
    ElseIf Range("A1") = "Run DT" Then
    
    Else
        
        responseFromMessageBox = MsgBox("""Run Date"" was not found in cell A1.  Would you like to continue?", vbYesNo)
        
        If responseFromMessageBox = 7 Then
        
            Exit Sub
    
        End If
            
    End If

    Set beginingSheet = ActiveSheet

    Set uSFProdNumberPosition = Range("A1:IV1").Find("USF Prod #", lookat:=xlWhole)
    
    uSFProdNumberColumnLetter = Evaluate("MID(""" & uSFProdNumberPosition.Address & """,FIND(""$"",""" & uSFProdNumberPosition.Address & """)+1,FIND(""$"",""" & uSFProdNumberPosition.Address & """,2)-2)")

    Columns(uSFProdNumberColumnLetter).Select
    
    Selection.Copy
    
    On Error Resume Next
    
    Sheets.Add(After:=beginingSheet).Name = "MerlinReportParameter"
    
    Set newSheet = ActiveSheet
    
    newSheet.Paste
    Application.CutCopyMode = False
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    newSheet.Range("$A$1:$A$" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    
    beginingSheet.Select
    
    Set divNumberPosition = Range("A1:IV1").Find("Div #", lookat:=xlWhole)
    
    divNumberColumnLetter = Evaluate("MID(""" & divNumberPosition.Address & """,FIND(""$"",""" & divNumberPosition.Address & """)+1,FIND(""$"",""" & divNumberPosition.Address & """,2)-2)")
    
    Columns(divNumberColumnLetter).Select
    
    Selection.Copy
    newSheet.Select
    Range("D1").Select
    newSheet.Paste
    Application.CutCopyMode = False
    newSheet.Range("$D$1:$D$" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    
    productNumberlastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Select Case productNumberlastRow
        
        Case 2
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
        
        Case 3
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("B3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"",""&RC[-1]"
        
        Case Is > 3
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("B3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"",""&RC[-1]"
            Range("B3").Select
            Selection.AutoFill Destination:=Range("B3:B" & productNumberlastRow)
    
    End Select
    
    divisionNumberLastRow = Cells(Rows.Count, 4).End(xlUp).Row
    
    Select Case divisionNumberLastRow
        
        Case 2
            Range("E2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
        
        Case 3
            Range("E2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("E3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"",""&RC[-1]"
        
        Case Is > 3
            Range("E2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("E3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"",""&RC[-1]"
            Range("E3").Select
            Selection.AutoFill Destination:=Range("E3:E" & divisionNumberLastRow)
    
    End Select
    
    For Each lookingForFolder In CreateObject("Shell.Application").Windows
    
        If lookingForFolder = "Windows Explorer" Then
            
            If lookingForFolder.locationurl = "file:///C:/Users/" & Environ("username") Then
                
                lookingForFolderFound = True
                
                Exit For
            
            End If
        
        End If
    
    Next lookingForFolder
    
    If lookingForFolderFound = True Then
    
        lookingForFolder.Visible = True
    
    Else
    
        ThisWorkbook.FollowHyperlink "file:///C:/Users/" & Environ("username")
    
    End If
    
'    lookingForFolderFound = Empty
'
'    For Each lookingForFolder In CreateObject("Shell.Application").Windows
'
'        If lookingForFolder = "Windows Explorer" Then
'
'            If lookingForFolder.locationurl = "file://1pxprinffs01/Group/National%20SIS/Shared/Remedy/Macros/Archive" Then
'
'                lookingForFolderFound = True
'
'                Exit For
'
'            End If
'
'        End If
'
'    Next lookingForFolder
'
'    If lookingForFolderFound = True Then
'
'        lookingForFolder.Visible = True
'
'    Else
    
'        ThisWorkbook.FollowHyperlink "file://1pxprinffs01/Group/National%20SIS/Shared/Remedy/Macros/Archive"
    
'    End If
    
    ThisWorkbook.FollowHyperlink "\\1pxprinffs01\Group\National SIS\Shared\Remedy\Macros\Archive\Pull Purchase Data into Compliance Report With Ship From Info.xlsm"
    
    For Each iE In CreateObject("Shell.Application").Windows
    
        If iE = "Internet Explorer" Then
            
            If iE.locationname = "OracleBI Discoverer" Then
                
                iEFound = True
                
                Exit For
            
            End If
        
        End If
    
    Next iE
    
    If iEFound = True Then
    
        iE.Visible = True

    Else
    
        Set iE = CreateObject("InternetExplorer.Application")
        
        iE.Visible = True
        
        iE.navigate "http://merlin/analytics/saw.dll?Dashboard"
        
        While iE.busy
        DoEvents
        Wend
        
        Application.Wait Now + TimeSerial(0, 0, 1.5)
        
        For Each discovererLink In iE.document.body.getelementsbytagname("a")
            
            If discovererLink.Title = "Discoverer" Then
            
                Exit For
            
            End If
                   
        Next discovererLink
        
        discovererLink.Click
  
        While iE.busy
        DoEvents
        Wend
        
        Application.Wait Now + TimeSerial(0, 0, 1.5)
        
        iE.document.all("connect").Click
  
        Application.Wait Now + TimeSerial(0, 0, 5)
  
        iE.Quit
  
    End If
    
Set beginingSheet = Nothing
divisionNumberLastRow = Empty
divNumberColumnLetter = ""
Set divNumberPosition = Nothing
Set iE = Nothing
iEFound = Empty
lastRow = Empty
Set lookingForFolder = Nothing
lookingForFolderFound = Empty
Set newSheet = Nothing
productNumberlastRow = Empty
responseFromMessageBox = Empty
uSFProdNumberColumnLetter = ""
Set uSFProdNumberPosition = Nothing
    
End Sub

Sub subMerlinGetProdDivTsfFrmCmplnc()

Dim beginingSheet As Worksheet
Dim companyVendorNameColumnLetter As String
Dim companyVendorNamePosition As Range
Dim divisionNumberLastRow As Long
Dim lastRow As Long
Dim newSheet As Worksheet
Dim productNumberlastRow As Long

Dim companyVendorNumberPosition As Range
Dim companyVendorNumberColumnLetter As String
Dim responseFromMessageBox As Long

Set beginingSheet = Nothing
companyVendorNameColumnLetter = ""
Set companyVendorNamePosition = Nothing
divisionNumberLastRow = Empty
lastRow = Empty
Set newSheet = Nothing
productNumberlastRow = Empty

    If Range("A1") <> "Run Date" Then
    
        responseFromMessageBox = MsgBox("""Run Date"" was not found in cell A1.  Would you like to continue?", vbYesNo)
        
        If responseFromMessageBox = 7 Then
        
            Exit Sub
    
        End If
    
    End If

    Set beginingSheet = ActiveSheet

    Set companyVendorNamePosition = Range("A1:IV1").Find("Company Vendor Name", lookat:=xlWhole)
    
    companyVendorNameColumnLetter = Evaluate("MID(""" & companyVendorNamePosition.Address & """,FIND(""$"",""" & companyVendorNamePosition.Address & """)+1,FIND(""$"",""" & companyVendorNamePosition.Address & """,2)-2)")

    Columns(companyVendorNameColumnLetter & ":" & companyVendorNameColumnLetter).Copy
    
    Sheets.Add(After:=beginingSheet).Name = "DivTransferReportParameter"
    
    Set newSheet = ActiveSheet
    
    newSheet.Paste
    
    Application.CutCopyMode = False
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    newSheet.Range("$A$1:$A$" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("A2").Select
    
    Do Until ActiveCell.Row > lastRow
    
        If InStr(ActiveCell, "US FOODS") > 0 Then
        
            ActiveCell.Offset(1).Select
        
        Else
    
            Rows(ActiveCell.Row).EntireRow.Delete
    
            lastRow = lastRow - 1
        
        End If
    
    Loop
    
    Range("C2:C" & lastRow).Formula = "=A2"
    
    Range("C2:C" & lastRow).Value = Range("C2:C" & lastRow).Value
    
    Range("B2:B" & lastRow).Formula = "=MID(A2,10,2)"
    
    Range("B2:B" & lastRow).Value = Range("B2:B" & lastRow).Value
    
    Range("A2:A" & lastRow).Value = Range("B2:B" & lastRow).Value
    
    Range("B2:B" & lastRow).Formula = "=INDEX('G:\National SIS\Shared\[Division Servers.xlsx]Sheet1'!$A$1:$A$61,MATCH(A2,'G:\National SIS\Shared\[Division Servers.xlsx]Sheet1'!$D$1:$D$61,0))"
    
    Range("B2:B" & lastRow).Value = Range("B2:B" & lastRow).Value
    
    Range("A2:A" & lastRow).Value = Range("B2:B" & lastRow).Value
    
    beginingSheet.Select
    
    Columns("T:T").Copy
    
    newSheet.Select
    
    Range("D1").Select
    
    newSheet.Paste
    
    Application.CutCopyMode = False
    
    lastRow = Cells(Rows.Count, 4).End(xlUp).Row
    
    newSheet.Range("$D$1:$D$" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    
    productNumberlastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Select Case productNumberlastRow
        
        Case 2
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
        
        Case 3
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("B3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"",""&RC[-1]"
        
        Case Is > 3
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("B3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"",""&RC[-1]"
            Range("B3").Select
            Selection.AutoFill Destination:=Range("B3:B" & productNumberlastRow)
    
    End Select
    
    divisionNumberLastRow = Cells(Rows.Count, 4).End(xlUp).Row
    
    Select Case divisionNumberLastRow
        
        Case 2
            Range("E2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
        
        Case 3
            Range("E2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("E3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"",""&RC[-1]"
        
        Case Is > 3
            Range("E2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("E3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"",""&RC[-1]"
            Range("E3").Select
            Selection.AutoFill Destination:=Range("E3:E" & divisionNumberLastRow)
    
    End Select
    
    beginingSheet.Activate
    
    Set companyVendorNumberPosition = Range("A1:IV1").Find("Company Vendor Number", lookat:=xlWhole)
    
    companyVendorNumberColumnLetter = Evaluate("MID(""" & companyVendorNumberPosition.Address & """,FIND(""$"",""" & companyVendorNumberPosition.Address & """)+1,FIND(""$"",""" & companyVendorNumberPosition.Address & """,2)-2)")
    
    Range(companyVendorNumberColumnLetter & 2).Select
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Do Until ActiveCell.Row > lastRow
    
        If InStr(Range(companyVendorNameColumnLetter & ActiveCell.Row), "US FOODS") > 0 Then
        
            ActiveCell = Evaluate("INDEX(" & newSheet.Name & "!A:A,MATCH(" & companyVendorNameColumnLetter & ActiveCell.Row & "," & newSheet.Name & "!C:C,0))")
        
        End If
        
        ActiveCell.Offset(1).Select
    
    Loop
    
Set beginingSheet = Nothing
companyVendorNameColumnLetter = ""
Set companyVendorNamePosition = Nothing
divisionNumberLastRow = Empty
lastRow = Empty
Set newSheet = Nothing
productNumberlastRow = Empty

    
End Sub

Sub subIEPRIMESearchFromCompliance()

    If Range("A1") = "Run Date" Then
    
    ElseIf Range("A1") = "Run DT" Then
    
    Else
    
        MsgBox ("Compliance report was not found.  This macro looks for ""Run Date"" or ""Run DT"" in range(""A1"").  Exiting.")
        
        Exit Sub
    
    End If
    
Dim contractNumber As Long
Dim contractNumberColumnNumber As Long
Dim iE As Object
Dim loginbutton As Variant
Dim passWord As String
Dim productNumber As Long
Dim productNumberColumnNumber As Long
Dim shipDate As Date
Dim shipDateColumnNumber As Long
Dim timeOut As Variant
Dim userName As String
Dim contractNumberPosition As Range
Dim productNumberPosition As Range
Dim shipDatePosition As Range

contractNumber = Empty
contractNumberColumnNumber = Empty
Set iE = Nothing
Set loginbutton = Nothing
passWord = ""
productNumber = Empty
productNumberColumnNumber = Empty
shipDate = Empty
shipDateColumnNumber = Empty
Set timeOut = Nothing
userName = ""
    
    Set contractNumberPosition = Range("A1:IV1").Find("Cntrct #", lookat:=xlWhole)
    
    If contractNumberPosition Is Nothing Then
    
        Set contractNumberPosition = Range("A1:IV1").Find("Contract", lookat:=xlWhole)
    
    End If
    
    contractNumberColumnNumber = contractNumberPosition.Column
    
    Set productNumberPosition = Range("A1:IV1").Find("USF Prod #", lookat:=xlWhole)
    
    If productNumberPosition Is Nothing Then
    
        Set productNumberPosition = Range("A1:IV1").Find("USF Prod Nbr", lookat:=xlWhole)
    
    End If
    
    productNumberColumnNumber = productNumberPosition.Column
    
    Set shipDatePosition = Range("A1:IV1").Find("Ship Date", lookat:=xlWhole)
    
    If shipDatePosition Is Nothing Then
    
        Set shipDatePosition = Range("A1:IV1").Find("Ship DT", lookat:=xlWhole)
    
    End If
    
    shipDateColumnNumber = shipDatePosition.Column
    
    If ActiveCell.Row = 1 Then
    
        ActiveCell.Offset(1).Select
    
    End If
    
    contractNumber = Cells(ActiveCell.Row, contractNumberColumnNumber)
    
    productNumber = Cells(ActiveCell.Row, productNumberColumnNumber)
    
    shipDate = Cells(ActiveCell.Row, shipDateColumnNumber)
    
    For Each iE In CreateObject("Shell.Application").Windows
    
        If iE = "Internet Explorer" Then
            
            Debug.Print iE.locationname
        
            If InStr(iE.locationname, "PRIME: Pricing Rules") > 0 Then
                iE.Visible = True
                Exit For
            End If
        End If
    
    Next iE
    
    If iE Is Nothing Then
    
        Set iE = CreateObject("InternetExplorer.Application")
        
        iE.navigate "http://prime.usfood.com/Prime/splash.do"
        
        iE.Visible = True
    
        While iE.busy
        DoEvents
        Wend
    
        Application.Wait Now + TimeSerial(0, 0, 1.5)
    
        If InStr(iE.locationurl, "http://prime.usfood.com/Prime/login.jsp") > 0 Then
        
            Do Until iE.locationurl = "http://prime.usfood.com/Prime/splash.do"
            
                MsgBox ("Log in then click Ok")
            
                While iE.busy
                DoEvents
                Wend
                
                Application.Wait Now + TimeSerial(0, 0, 3)
            
                If InStr(iE.locationurl, "http://prime.usfood.com/Prime/login.jsp") > 0 Then
                
                    Do Until iE.locationurl = "http://prime.usfood.com/Prime/splash.do"
                    
                        MsgBox ("Log in then click Ok")
                    
                        While iE.busy
                        DoEvents
                        Wend
                        
                        Application.Wait Now + TimeSerial(0, 0, 3)
                    
                    Loop
                
                End If
            
            Loop
        
        End If
        
    End If
    
thereWasAProblemReDo:
    
    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeSerial(0, 0, 1.5)
    
    'Navigates to the CONTRACT section in PRIME
    iE.navigate "http://prime.usfood.com/Prime/contract/find.do"
    
    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeSerial(0, 0, 1.5)
    
    iE.document.all("inquiryDateField").Value = shipDate
    
    'Search for the "contractNumber" field in PRIME and enter in the value from the spreadsheet, it's dimensioned above
    iE.document.all("contractNumber").Value = contractNumber
    
    Application.Wait Now + TimeSerial(0, 0, 1.5)
    
    'Search for the "QuickSubmit" button and Click to search for the contract number
    iE.document.all("QuickSubmit").Click
    
    While iE.busy
    DoEvents
    Wend
        
    'Navigate to Contract > General > Header
    iE.navigate "http://prime.usfood.com/Prime/contract/searchresults.do?number=" & contractNumber & "&contractStatus=INFC&isworkarea=false&waNbr=0&waEffDate=&action=view"
    
    While iE.busy
    DoEvents
    Wend
        
    Application.Wait Now + TimeSerial(0, 0, 1.5)
        
    iE.navigate "http://prime.usfood.com/Prime/contract/listSharedGroups.do?groupsTab=1"
        
    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeSerial(0, 0, 1.5)
    
    'There was a problem with PRIME error handling
    For Each timeOut In iE.document.getelementsbytagname("div")
    
    '    Debug.Print timeOut.innertext
    
        If InStr(timeOut.innertext, "There was a problem") Then
        
    '        MsgBox ("Found there was a problem")
            GoTo thereWasAProblemReDo
        End If
        
    Next
    
    On Error Resume Next
    
    'Find the product number field in PRIME and enter in the product number from the spreadsheet
    iE.document.all("goToProductNbr").Value = productNumber
    
    If Err.Number = 91 Then
    
        iE.navigate "http://prime.usfood.com/Prime/contract/listSpecificGroups.do"
        
        While iE.busy
        DoEvents
        Wend
    
        Application.Wait Now + TimeSerial(0, 0, 1.5)
    
        iE.document.all("goToProductNbr").Value = productNumber
    
    End If
    
    While iE.busy
    DoEvents
    Wend
        
    Application.Wait Now + TimeSerial(0, 0, 1.5)
    
    'Find the "Search" button and click that mug
    iE.document.all("FindProduct").Click
    
    While iE.busy
    DoEvents
    Wend
  
'    MsgBox ("Complete")
  
contractNumber = Empty
contractNumberColumnNumber = Empty
Set iE = Nothing
Set loginbutton = Nothing
passWord = ""
productNumber = Empty
productNumberColumnNumber = Empty
shipDate = Empty
shipDateColumnNumber = Empty
Set timeOut = Nothing
userName = ""

End Sub

Sub iECASISContractProductHisotrySrchBasedOnActiveRow()


'''''' If this is updated, also update the macro in the folder "G:\National SIS\Shared\Remedy\Macros\iECASISContractProductHisotrySrchBasedOnActiveRow.xlsm"

'Set IE = CreateObject("InternetExplorer.Application")
'
'IE.navigate "http://cas.usfood.com/CAS/contractProductHistory.jsp?productNumber=" & Worksheets("Sheet0").Range("F" & (ActiveCell.Row)) & "&contractNumber=" & Worksheets("Sheet0").Range("D" & (ActiveCell.Row))
'
'IE.Visible = True

Dim iE As Object

Set iE = Nothing
    
    For Each iE In CreateObject("Shell.Application").Windows
    
        If iE = "Windows Internet Explorer" Then
            If InStr(iE.locationurl, "http://cas.usfood.com/CAS/contractProductHistory.jsp?productNumber=" & Worksheets("Sheet0").Range("F" & (ActiveCell.Row)) & "&contractNumber=" & Worksheets("Sheet0").Range("D" & (ActiveCell.Row))) <> 0 Then
    
                Exit For
    
            End If
    
        'Use code below to locate hidden IE windows
        'IE.Visible = True
    
        End If
    
    Next iE
    
    If iE Is Nothing Then
    
    '''    Set IE = CreateObject("InternetExplorer.Application")
    '''    IE.navigate "http://cas.usfood.com/CAS/specialBillOpportunity.jsp?action=doSearch"
    
        Set iE = CreateObject("InternetExplorer.Application")
        iE.navigate "http://cas.usfood.com/CAS/login.jsp"
        iE.Visible = True
    
        Application.Wait Now + 1 / (24 * 60 * 60# * 2) 'TimeValue("00:00:01")
        
        iE.document.all("j_username").Value = "D3Q1700"
        iE.document.all("j_password").Value = "today123456789"
        
        Application.Wait Now + 1 / (24 * 60 * 60# * 2) 'TimeValue("00:00:01")
        
        iE.document.all("login").Click
        
        Application.Wait Now + 1 / (24 * 60 * 60# * 2) 'TimeValue("00:00:01")
    
    End If
    
    'Display Internet Explorer since the default is not to display it
    iE.Visible = True
    
    While iE.busy
    'wait until IE is done loading the page.
    DoEvents
    Wend
    
    iE.navigate "http://cas.usfood.com/CAS/contractProductHistory.jsp?productNumber=" & ActiveSheet.Range("F" & (ActiveCell.Row)) & "&contractNumber=" & Worksheets("Sheet0").Range("D" & (ActiveCell.Row))

Set iE = Nothing

End Sub



Sub subIESpecialBillFromComplianceMultiples()

    Application.Run "'G:\National SIS\Shared\Remedy\Macros\Special Bills By Compliance Report Macro.xlsm'!subIESpecialBillFromComplianceMultiples"
    
End Sub

Sub subIEContractProductHistoryFromCompliance()

    If Cells(1, 1) <> "Run Date" Then
    
        MsgBox ("""Run Date"" was not found in range (A1).  This does not seem to be a compliance report.  Exiting.")
    
        Exit Sub
    
    ElseIf Cells(ActiveCell.Row, 1) = "" Then
    
        MsgBox ("Current row and column A detected as blank.  Select a row with data.  Exiting")
        
        Exit Sub
        
    ElseIf ActiveCell.Row = 1 Then
    
        MsgBox ("Header row 1 detected.  Select a row with data.  Exitiing")
        
        Exit Sub
        
    End If

Dim contractColumnLetter As String
Dim contractNumber As Long
Dim contractPosition As Range
Dim iE As Object
Dim productColumnLetter As String
Dim productNumber As Long
Dim productPosition As Range

contractColumnLetter = ""
contractNumber = Empty
Set contractPosition = Nothing
Set iE = Nothing
productColumnLetter = ""
productNumber = Empty
Set productPosition = Nothing
    
    Set contractPosition = Range("A1:IV1").Find("Cntrct #", lookat:=xlWhole)
    
    contractColumnLetter = Evaluate("MID(""" & contractPosition.Address & """,FIND(""$"",""" & contractPosition.Address & """)+1,FIND(""$"",""" & contractPosition.Address & """,2)-2)")
    
    contractNumber = Range(contractColumnLetter & ActiveCell.Row)
    
    Set productPosition = Range("A1:IV1").Find("USF Prod #", lookat:=xlWhole)
    
    productColumnLetter = Evaluate("MID(""" & productPosition.Address & """,FIND(""$"",""" & productPosition.Address & """)+1,FIND(""$"",""" & productPosition.Address & """,2)-2)")
    
    productNumber = Range(productColumnLetter & ActiveCell.Row)
    
    Set iE = CreateObject("InternetExplorer.Application")
    
    iE.navigate "http://cas.usfood.com/CAS/contractProductHistory.jsp?productNumber=" & productNumber & "&contractNumber=" & contractNumber
    
    iE.Visible = True
    
    Application.Wait Now + TimeSerial(0, 0, 1.5)
    
    Do Until iE.locationurl <> "http://cas.usfood.com/CAS/login.jsp"
    
        MsgBox ("Log in to CASIS then click Ok")
    
        While iE.busy
        DoEvents
        Wend
        
        Application.Wait Now + TimeSerial(0, 0, 1)
    
    Loop
    
contractColumnLetter = ""
contractNumber = Empty
Set contractPosition = Nothing
Set iE = Nothing
productColumnLetter = ""
productNumber = Empty
Set productPosition = Nothing

End Sub

Sub subIEInvoiceReversalInvoiceSelect()

Application.Run "'G:\National SIS\Shared\Remedy\Macros\Invoice Reversals Macro.xlsm'!subIEInvoiceReversalInvoiceSelect"
'
'Dim currentInvoiceAmountFromCompliance As Double
'Dim currentInvoiceNumberFromIE As Long
'Dim getPosition As Range
'Dim i As Long
'Dim iE As Object
'Dim lastColumnColumnLetter As String
'Dim lastColumnPosition As Range
'Dim lastRow As Long
'Dim oldestInvoiceDate As Date
'Dim programMaster As String
'Dim programMasterColumnLetter As String
'Dim RecordCount As Long
'Dim resolutionColumnLetter As String
'Dim reversalCheckBox As Object
'Dim runDateColumnLetter As String
'Dim runDateColumnNumber As Long
'Dim sISInvColumnLetter As String
'Dim today As Date
'Dim variableCheckBox As String
'
'currentInvoiceAmountFromCompliance = Empty
'currentInvoiceNumberFromIE = Empty
'Set getPosition = Nothing
'i = Empty
'Set iE = Nothing
'lastColumnColumnLetter = ""
'Set lastColumnPosition = Nothing
'lastRow = Empty
'oldestInvoiceDate = Empty
'programMaster = ""
'programMasterColumnLetter = ""
'RecordCount = Empty
'resolutionColumnLetter = ""
'Set reversalCheckBox = Nothing
'runDateColumnLetter = ""
'runDateColumnNumber = Empty
'sISInvColumnLetter = ""
'today = Empty
'variableCheckBox = ""
'
'    Set getPosition = Range("A1:IV1").Find("Run Date", lookat:=xlWhole)
'
'    If getPosition Is Nothing Then
'
'        MsgBox ("Run Date field was not found.  Exiting")
'
'        Exit Sub
'
'    End If
'
'    runDateColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    runDateColumnNumber = getPosition.Column
'
'    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
'
'    Set lastColumnPosition = Cells(1, Columns.count).End(xlToLeft)
'
'    lastColumnColumnLetter = Evaluate("MID(""" & lastColumnPosition.Address & """,FIND(""$"",""" & lastColumnPosition.Address & """)+1,FIND(""$"",""" & lastColumnPosition.Address & """,2)-2)")
'
'    lastRow = Cells(Rows.count, 1).End(xlUp).Row
'
'    Range("A1:" & lastColumnColumnLetter & lastRow).AutoFilter
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(runDateColumnLetter & "1:" & runDateColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'
'    With ActiveSheet.AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    oldestInvoiceDate = Evaluate("Min(" & runDateColumnLetter & ":" & runDateColumnLetter & ")")
'
'    today = Date
'
'    Set getPosition = Range("A1:IV1").Find("Program Master", lookat:=xlWhole)
'
'    programMasterColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    programMaster = Range(programMasterColumnLetter & 2)
'
'
'    Set iE = CreateObject("internetexplorer.application")
'
'    iE.Visible = True
'
'    iE.navigate "http://cas.usfood.com/CAS/supplierInvoiceSearch.jsp"
'
'    While iE.busy
'    DoEvents
'    Wend
'
'    Application.Wait Now + TimeSerial(0, 0, 1)
'
'    Do Until iE.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'        MsgBox ("Log in to CASIS then click Ok")
'
'        While iE.busy
'        DoEvents
'        Wend
'
'        Application.Wait Now + TimeSerial(0, 0, 1)
'
'    Loop
'
'    iE.document.all("from").Value = oldestInvoiceDate
'
'    iE.document.all("to").Value = today
'
'    iE.document.all("masterNbr").Value = programMaster
'
'    iE.document.all("doSearch").Click
'
'    While iE.busy
'    DoEvents
'    Wend
'
'    Application.Wait Now + TimeSerial(0, 0, 1)
'
'    Do Until iE.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'        MsgBox ("Log in to CASIS then click Ok")
'
'        While iE.busy
'        DoEvents
'        Wend
'
'        Application.Wait Now + TimeSerial(0, 0, 1)
'
'    Loop
'
'    Set getPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole)
'
'    resolutionColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("SIS Inv #", lookat:=xlWhole)
'
'    sISInvColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    RecordCount = iE.document.all("recordCount").Value
'
''   i is a number and variableCheckBox is a string.  The reversalCheckBox.Value returns a string and variableCheckBox is needed to compare it.
'
'    variableCheckBox = 0
'
'    i = 0
'
'    Do Until i >= RecordCount
'
'        currentInvoiceNumberFromIE = iE.document.all("invoiceList[" & variableCheckBox & "].invoice").Value
'
'        currentInvoiceAmountFromCompliance = Evaluate("SUMIFS($" & resolutionColumnLetter & "$1:$" & resolutionColumnLetter & "$" & lastRow & ",$" & sISInvColumnLetter & "$1:$" & sISInvColumnLetter & "$" & lastRow & "," & currentInvoiceNumberFromIE & ")")
'
'        If currentInvoiceAmountFromCompliance < 0 Then
'
'            For Each reversalCheckBox In iE.document.body.GetElementsByTagName("input")
'
''                Debug.Print reversalCheckBox.Type
''
''                Debug.Print reversalCheckBox.ID
''
''                Debug.Print reversalCheckBox.Value
'
'                If reversalCheckBox.ID = "check_inv" Then
'
'                    If reversalCheckBox.Value = variableCheckBox Then
'
'                        reversalCheckBox.Click
'
''                        MsgBox ("Matching reversal check box found.")
'
'                        Exit For
'
'                    End If
'
'                End If
'
'            Next reversalCheckBox
'
'        End If
'
'        variableCheckBox = variableCheckBox + 1
'
'        i = i + 1
'
'    Loop
'
'    MsgBox ("Complete!")
'
'currentInvoiceAmountFromCompliance = Empty
'currentInvoiceNumberFromIE = Empty
'Set getPosition = Nothing
'i = Empty
'Set iE = Nothing
'lastColumnColumnLetter = ""
'Set lastColumnPosition = Nothing
'lastRow = Empty
'oldestInvoiceDate = Empty
'programMaster = ""
'programMasterColumnLetter = ""
'RecordCount = Empty
'resolutionColumnLetter = ""
'Set reversalCheckBox = Nothing
'runDateColumnLetter = ""
'runDateColumnNumber = Empty
'sISInvColumnLetter = ""
'today = Empty
'variableCheckBox = ""
    
End Sub

Sub subIEInvoiceReversalTrackingProductSelect()

Application.Run "'G:\National SIS\Shared\Remedy\Macros\Invoice Reversals Macro.xlsm'!subIEInvoiceReversalTrackingProductSelect"
'
'Dim currentInvoiceAmountFromCompliance As Double
'Dim currentInvoiceAmountFromIETag As Double
'Dim currentInvoiceBeginningRow As Long
'Dim currentInvoiceDateFromIEString As String
'Dim currentInvoiceLastRow As Long
'Dim currentInvoiceNumberFromIELong As Long
'Dim currentInvoiceNumberFromIEString As String
'Dim currentInvoiceNumberFromIETag As String
'Dim currentInvoiceRunDate As String
'Dim currentMasterProgramNumber As String
'Dim currentTrackingAmountFromCompliance As Double
'Dim currentTrackingProgramNumber As String
'Dim currentTrackingProgramReversalAmountDouble As Double
'Dim currentTrackingProgramReversalAmountString As String
'Dim getPosition As Range
'Dim i As Long
'Dim iE As Object
'Dim iESpecificInvoice As Object
'Dim lastColumnColumnLetter As String
'Dim lastColumnPosition As Range
'Dim lastRow As Long
'Dim lookForVendorNameForError As String
'Dim oldestInvoiceDate As Date
'Dim programMaster As String
'Dim programMasterColumnLetter As String
'Dim RecordCount As Long
'Dim recordsProcessed As Long
'Dim resolutionColumnLetter As String
'Dim runDateColumnLetter As String
'Dim runDateColumnNumber As Long
'Dim sISInvColumnLetter As String
'Dim tdTag As Object
'Dim tdTagsToReachInvoice As Long
'Dim tdTagsToReachReversalAmount As Long
'Dim today As Date
'Dim trackingProgramCount As Variant
'Dim trackingProgramNumberColumnLetter As String
'Dim uSFProdNumberColumnLetter As String
'Dim variableCheckBox As String
'Dim variableNumber As Variant
'Dim x As Long
'
'currentInvoiceAmountFromCompliance = Empty
'currentInvoiceAmountFromIETag = Empty
'currentInvoiceBeginningRow = Empty
'currentInvoiceDateFromIEString = ""
'currentInvoiceLastRow = Empty
'currentInvoiceNumberFromIELong = Empty
'currentInvoiceNumberFromIEString = ""
'currentInvoiceNumberFromIETag = ""
'currentInvoiceRunDate = ""
'currentMasterProgramNumber = ""
'currentTrackingAmountFromCompliance = Empty
'currentTrackingProgramNumber = ""
'currentTrackingProgramReversalAmountDouble = Empty
'currentTrackingProgramReversalAmountString = ""
'Set getPosition = Nothing
'i = Empty
'Set iE = Nothing
'Set iESpecificInvoice = Nothing
'lastColumnColumnLetter = ""
'Set lastColumnPosition = Nothing
'lastRow = Empty
'lookForVendorNameForError = ""
'oldestInvoiceDate = Empty
'programMaster = ""
'programMasterColumnLetter = ""
'RecordCount = Empty
'recordsProcessed = Empty
'resolutionColumnLetter = ""
'runDateColumnLetter = ""
'runDateColumnNumber = Empty
'sISInvColumnLetter = ""
'Set tdTag = Nothing
'tdTagsToReachInvoice = Empty
'tdTagsToReachReversalAmount = Empty
'today = Empty
'Set trackingProgramCount = Nothing
'trackingProgramNumberColumnLetter = ""
'uSFProdNumberColumnLetter = ""
'variableCheckBox = ""
'Set variableNumber = Nothing
'x = Empty
'
'    Set getPosition = Range("A1:IV1").Find("Run Date", lookat:=xlWhole)
'
'    If getPosition Is Nothing Then
'
'        MsgBox ("Run Date field was not found.  Exiting")
'
'        Exit Sub
'
'    End If
'
'    runDateColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    runDateColumnNumber = getPosition.Column
'
'    Set getPosition = Range("A1:IV1").Find("Program Master", lookat:=xlWhole)
'
'    programMasterColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole)
'
'    resolutionColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("SIS Inv #", lookat:=xlWhole)
'
'    sISInvColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Tracking Program #", lookat:=xlWhole)
'
'    trackingProgramNumberColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("USF Prod #", lookat:=xlWhole)
'
'    uSFProdNumberColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
'
'    Set lastColumnPosition = Cells(1, Columns.count).End(xlToLeft)
'
'    lastColumnColumnLetter = Evaluate("MID(""" & lastColumnPosition.Address & """,FIND(""$"",""" & lastColumnPosition.Address & """)+1,FIND(""$"",""" & lastColumnPosition.Address & """,2)-2)")
'
'    lastRow = Cells(Rows.count, 1).End(xlUp).Row
'
'    Range("A1:" & lastColumnColumnLetter & lastRow).AutoFilter
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(programMasterColumnLetter & "2:" & programMasterColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(runDateColumnLetter & "2:" & runDateColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(sISInvColumnLetter & "2:" & sISInvColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(trackingProgramNumberColumnLetter & "2:" & trackingProgramNumberColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(uSFProdNumberColumnLetter & "2:" & uSFProdNumberColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'
'    With ActiveSheet.AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    oldestInvoiceDate = Evaluate("Min(" & runDateColumnLetter & ":" & runDateColumnLetter & ")")
'
'    today = Date
'
'    programMaster = Range(programMasterColumnLetter & 2)
'
'    Set iE = CreateObject("internetexplorer.application")
'
'    iE.Visible = True
'
'    iE.navigate "http://cas.usfood.com/CAS/invoiceReversalSearch.jsp"
'
'    While iE.busy
'    DoEvents
'    Wend
'
'    Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'    Do Until iE.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'        MsgBox ("Log in to CASIS then click Ok")
'
'        While iE.busy
'        DoEvents
'        Wend
'
'        Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'    Loop
'
'    iE.document.all("mstrPgmNbr").Value = programMaster
'
'    iE.document.all("statuses").Value = "P"
'
'    iE.navigate "javascript:submitAction('search')"
'
'    While iE.busy
'    DoEvents
'    Wend
'
'    Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'    Do Until iE.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'        MsgBox ("Log in to CASIS then click Ok")
'
'        While iE.busy
'        DoEvents
'        Wend
'
'        Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'    Loop
'
'    RecordCount = iE.document.all("recordCount").Value
'
'    recordsProcessed = 0
'
'    variableNumber = "" & recordsProcessed & ""
'
'    tdTagsToReachInvoice = 86
'
'    tdTagsToReachReversalAmount = 106
'
'    Do Until recordsProcessed >= RecordCount
'
'        currentInvoiceDateFromIEString = iE.document.all(variableNumber).Value
'
'        currentInvoiceDateFromIEString = Left(currentInvoiceDateFromIEString, 8)
'
'        currentInvoiceNumberFromIEString = iE.document.all(variableNumber).Value
'
'        currentInvoiceNumberFromIEString = Left(currentInvoiceNumberFromIEString, 16)
'
'        currentInvoiceNumberFromIEString = Right(currentInvoiceNumberFromIEString, 7)
'
'        currentInvoiceNumberFromIELong = currentInvoiceNumberFromIEString
'
'        currentInvoiceAmountFromCompliance = Evaluate("SUMIFS($" & resolutionColumnLetter & "$1:$" & resolutionColumnLetter & "$" & lastRow & ",$" & sISInvColumnLetter & "$1:$" & sISInvColumnLetter & "$" & lastRow & "," & currentInvoiceNumberFromIELong & ")")
'
'        currentInvoiceAmountFromCompliance = Round(currentInvoiceAmountFromCompliance, 2)
'
'        If currentInvoiceAmountFromCompliance < 0 Then
'
'            i = 0
'
'            ' table is separated by 26 td tags
'
'            For Each tdTag In iE.document.body.GetElementsByTagName("td")
'
'                If i = tdTagsToReachInvoice Then
'
'                    currentInvoiceNumberFromIETag = tdTag.innertext
'
'                    currentInvoiceNumberFromIETag = Evaluate("SUBSTITUTE(" & currentInvoiceNumberFromIETag & ","" "","""")")
'
'                    If currentInvoiceNumberFromIEString = currentInvoiceNumberFromIETag Then
'
'                    Else
'
'                        MsgBox ("Invoice from Sequence hyperlink is not matching invoice from displayed table")
'
'                    End If
'
'                End If
'
'                If i >= tdTagsToReachReversalAmount Then
'
'                    currentInvoiceAmountFromIETag = tdTag.innertext
'
'                    Exit For
'
'                End If
'
'                i = i + 1
'
'            Next tdTag
'
'            If currentInvoiceAmountFromIETag = currentInvoiceAmountFromCompliance Then
'
'            Else
'
'                Set iESpecificInvoice = CreateObject("internetexplorer.application")
'
'                iESpecificInvoice.Visible = True
'
'                iESpecificInvoice.navigate "http://cas.usfood.com/CAS/invoiceReversalRequest.jsp?action=doSearch&originalInv=" & currentInvoiceNumberFromIEString & "&date=" & currentInvoiceDateFromIEString & "&seq=1"
'
'                While iESpecificInvoice.busy
'                DoEvents
'                Wend
'
'                Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                Do Until iESpecificInvoice.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'                    MsgBox ("Log in to CASIS then click Ok")
'
'                    While iESpecificInvoice.busy
'                    DoEvents
'                    Wend
'
'                    Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                Loop
'
'                trackingProgramCount = 0
'
'                Do Until Err.Number >= 91
'
'                    currentMasterProgramNumber = iESpecificInvoice.document.all("detailsList[" & trackingProgramCount & "].pgmMstrNbr").Value
'
'                    currentInvoiceRunDate = iESpecificInvoice.document.all("detailsList[" & trackingProgramCount & "].origRunDate").Value
'
'                    currentTrackingProgramNumber = iESpecificInvoice.document.all("detailsList[" & trackingProgramCount & "].programTrkNbr").Value
'
'                    currentTrackingProgramReversalAmountString = iESpecificInvoice.document.all("detailsList[" & trackingProgramCount & "].amtDue").Value
'
'                    currentTrackingProgramReversalAmountString = Evaluate("SUBSTITUTE(" & currentTrackingProgramReversalAmountString & ","" "","""")")
'
'                    currentTrackingProgramReversalAmountDouble = currentTrackingProgramReversalAmountString * -1
'
'                    currentTrackingAmountFromCompliance = Evaluate("SUMIFS($" & resolutionColumnLetter & "$1:$" & resolutionColumnLetter & "$" & lastRow & ",$" & sISInvColumnLetter & "$1:$" & sISInvColumnLetter & "$" & lastRow & "," & currentInvoiceNumberFromIELong & ",$" & trackingProgramNumberColumnLetter & "$1:$" & trackingProgramNumberColumnLetter & "$" & lastRow & ",""" & currentTrackingProgramNumber & """)")
'
'                    currentTrackingAmountFromCompliance = Round(currentTrackingAmountFromCompliance, 2)
'
'                    If currentTrackingProgramReversalAmountDouble = currentTrackingAmountFromCompliance Then
'
'                    ElseIf currentTrackingAmountFromCompliance >= 0 Then
'
'                        iESpecificInvoice.document.all("detailsList[" & trackingProgramCount & "].prodInd").Value = "IN"
'
'                        iESpecificInvoice.document.all("saveButton").Click
'
'                        While iESpecificInvoice.busy
'                        DoEvents
'                        Wend
'
'                        Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                        Do Until iESpecificInvoice.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'                            MsgBox ("Log in to CASIS then click Ok")
'
'                            While iESpecificInvoice.busy
'                            DoEvents
'                            Wend
'
'                            Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                        Loop
'
'                    Else
'
'                        iESpecificInvoice.document.all("detailsList[" & trackingProgramCount & "].prodInd").Value = "IN"
'
'                        iESpecificInvoice.document.all("saveButton").Click
'
'                        While iESpecificInvoice.busy
'                        DoEvents
'                        Wend
'
'                        Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                        Do Until iESpecificInvoice.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'                            MsgBox ("Log in to CASIS then click Ok")
'
'                            While iESpecificInvoice.busy
'                            DoEvents
'                            Wend
'
'                            Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                        Loop
'
'                        Set iESpecificInvoiceTracking = CreateObject("internetexplorer.application")
'
'                        iESpecificInvoiceTracking.Visible = True
'
'                        iESpecificInvoiceTracking.navigate "http://cas.usfood.com/CAS/invoiceReversalRequestSales.jsp?origSisInvNbr=" & currentInvoiceNumberFromIELong & "&pgmMstrNbr=" & currentMasterProgramNumber & "&origRunDate=" & currentInvoiceRunDate & "&reverslSeqNbr=1&sisInvDtlId=" & 1 + trackingProgramCount
'
'                        While iESpecificInvoiceTracking.busy
'                        DoEvents
'                        Wend
'
'                        Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                        Do Until iESpecificInvoiceTracking.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'                            MsgBox ("Log in to CASIS then click Ok")
'
'                            While iESpecificInvoiceTracking.busy
'                            DoEvents
'                            Wend
'
'                            Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                        Loop
'
'                        ' Start of loop to select products within tracking
'
'                        currentInvoiceBeginningRow = Range(sISInvColumnLetter & "1:" & sISInvColumnLetter & lastRow).Find(currentInvoiceNumberFromIELong, lookat:=xlWhole, searchdirection:=xlNext).Row
'
'                        currentInvoiceLastRow = Range(sISInvColumnLetter & "1:" & sISInvColumnLetter & lastRow).Find(currentInvoiceNumberFromIELong, lookat:=xlWhole, searchdirection:=xlPrevious).Row
'
'                        If Range(trackingProgramNumberColumnLetter & currentInvoiceBeginningRow) = currentTrackingProgramNumber Then
'
'                            currentInvoiceTrackingBeginningRow = currentInvoiceBeginningRow
'
'                        Else
'
'                            currentInvoiceTrackingBeginningRow = Range(trackingProgramNumberColumnLetter & currentInvoiceBeginningRow & ":" & trackingProgramNumberColumnLetter & currentInvoiceLastRow).Find(currentTrackingProgramNumber, lookat:=xlWhole, searchdirection:=xlNext).Row
'
'                        End If
'
'                        currentInvoiceTrackingLastRow = Range(trackingProgramNumberColumnLetter & currentInvoiceBeginningRow & ":" & trackingProgramNumberColumnLetter & currentInvoiceLastRow).Find(currentTrackingProgramNumber, lookat:=xlWhole, searchdirection:=xlPrevious).Row
'
'                        countOfInvoiceTrackingProducts = Evaluate("SUM(IF(FREQUENCY(MATCH(" & uSFProdNumberColumnLetter & currentInvoiceTrackingBeginningRow & ":" & uSFProdNumberColumnLetter & currentInvoiceTrackingLastRow & "," & uSFProdNumberColumnLetter & currentInvoiceTrackingBeginningRow & ":" & uSFProdNumberColumnLetter & currentInvoiceTrackingLastRow & ",0),MATCH(" & uSFProdNumberColumnLetter & currentInvoiceTrackingBeginningRow & ":" & uSFProdNumberColumnLetter & currentInvoiceTrackingLastRow & "," & uSFProdNumberColumnLetter & currentInvoiceTrackingBeginningRow & ":" & uSFProdNumberColumnLetter & currentInvoiceTrackingLastRow & ",0))>0,1))")
'
'                        x = 0
'
'                        Do Until x >= countOfInvoiceTrackingProducts
'
'                            currentInvoiceTrackingProductNumber = Range(uSFProdNumberColumnLetter & currentInvoiceTrackingBeginningRow)
'
'                            currentInvoiceTrackingProductAmount = Evaluate("SUMIFS($" & resolutionColumnLetter & "$" & currentInvoiceTrackingBeginningRow & ":$" & resolutionColumnLetter & "$" & currentInvoiceTrackingLastRow & ",$" & uSFProdNumberColumnLetter & "$" & currentInvoiceTrackingBeginningRow & ":$" & uSFProdNumberColumnLetter & "$" & currentInvoiceTrackingLastRow & "," & currentInvoiceTrackingProductNumber & ")")
'
'                            If currentInvoiceTrackingProductAmount < 0 Then
'
'                                iESpecificInvoiceTracking.document.all("productNbr").Value = currentInvoiceTrackingProductNumber
'
'                                iESpecificInvoiceTracking.document.all("Search").Click
'
'                                While iESpecificInvoiceTracking.busy
'                                DoEvents
'                                Wend
'
'                                Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                                Do Until iESpecificInvoiceTracking.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'                                    MsgBox ("Log in to CASIS then click Ok")
'
'                                    While iESpecificInvoiceTracking.busy
'                                    DoEvents
'                                    Wend
'
'                                    Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                                Loop
'
'                                iESpecificInvoiceTracking.document.all("prodSelectionCheckBox").Click
'
'                                iESpecificInvoiceTracking.document.all("Save").Click
'
'                                While iESpecificInvoiceTracking.busy
'                                DoEvents
'                                Wend
'
'                                Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                                Do Until iESpecificInvoiceTracking.locationURL <> "http://cas.usfood.com/CAS/login.jsp"
'
'                                    MsgBox ("Log in to CASIS then click Ok")
'
'                                    While iESpecificInvoiceTracking.busy
'                                    DoEvents
'                                    Wend
'
'                                    Application.Wait Now + TimeSerial(0, 0, 1.15)
'
'                                Loop
'
'                            End If
'
'                            'On Error Resume Next
'
'                            currentInvoiceTrackingBeginningRow = Evaluate("MATCH(FALSE," & uSFProdNumberColumnLetter & currentInvoiceTrackingBeginningRow & ":" & uSFProdNumberColumnLetter & currentInvoiceTrackingLastRow & "=" & currentInvoiceTrackingProductNumber & ",0)") + currentInvoiceTrackingBeginningRow - 1
'
'                            x = x + 1
'
'                        Loop
'
'                        iESpecificInvoiceTracking.Quit
'
'                    End If
'
'                    trackingProgramCount = trackingProgramCount + 1
'
'                    On Error Resume Next
'
'                    lookForVendorNameForError = iESpecificInvoice.document.all("detailsList[" & trackingProgramCount & "].vndrName").Value
'
'                Loop
'
'                Err.Number = 0
'
'                iESpecificInvoice.Quit
'
'            End If
'
'        End If
'
'        tdTagsToReachInvoice = tdTagsToReachInvoice + 26
'
'        tdTagsToReachReversalAmount = tdTagsToReachReversalAmount + 26
'
'        recordsProcessed = recordsProcessed + 1
'
'        variableNumber = "" & recordsProcessed & ""
'
'    Loop
'
'    MsgBox ("Complete!")
'
'currentInvoiceAmountFromCompliance = Empty
'currentInvoiceAmountFromIETag = Empty
'currentInvoiceBeginningRow = Empty
'currentInvoiceDateFromIEString = ""
'currentInvoiceLastRow = Empty
'currentInvoiceNumberFromIELong = Empty
'currentInvoiceNumberFromIEString = ""
'currentInvoiceNumberFromIETag = ""
'currentInvoiceRunDate = ""
'currentMasterProgramNumber = ""
'currentTrackingAmountFromCompliance = Empty
'currentTrackingProgramNumber = ""
'currentTrackingProgramReversalAmountDouble = Empty
'currentTrackingProgramReversalAmountString = ""
'Set getPosition = Nothing
'i = Empty
'Set iE = Nothing
'Set iESpecificInvoice = Nothing
'lastColumnColumnLetter = ""
'Set lastColumnPosition = Nothing
'lastRow = Empty
'lookForVendorNameForError = ""
'oldestInvoiceDate = Empty
'programMaster = ""
'programMasterColumnLetter = ""
'RecordCount = Empty
'recordsProcessed = Empty
'resolutionColumnLetter = ""
'runDateColumnLetter = ""
'runDateColumnNumber = Empty
'sISInvColumnLetter = ""
'Set tdTag = Nothing
'tdTagsToReachInvoice = Empty
'tdTagsToReachReversalAmount = Empty
'today = Empty
'Set trackingProgramCount = Nothing
'trackingProgramNumberColumnLetter = ""
'uSFProdNumberColumnLetter = ""
'variableCheckBox = ""
'Set variableNumber = Nothing
'x = Empty
    
End Sub





Sub subCallPullComplianceReportMacro()

Dim CreateReport As Variant
Dim invoiceDate As String
Dim invoiceDateColumnLetter As String
Dim invoiceDatePosition As Range
Dim masterProgramNumber As String
Dim masterProgramNumberColumnLetter As String
Dim masterProgramNumberPosition As Range

Dim oldestInvoiceDate As Date
Dim iE As Object

Set CreateReport = Nothing
invoiceDate = ""
invoiceDateColumnLetter = ""
Set invoiceDatePosition = Nothing
masterProgramNumber = ""
masterProgramNumberColumnLetter = ""
Set masterProgramNumberPosition = Nothing

    Set masterProgramNumberPosition = Range("A1:IV1").Find("Master Program Number", lookat:=xlWhole)
    
    If masterProgramNumberPosition Is Nothing Then
    
        Set masterProgramNumberPosition = Range("A1:IV1").Find("Program Number", lookat:=xlWhole)
    
        If masterProgramNumberPosition Is Nothing Then
    
            Set masterProgramNumberPosition = Range("A1:IV1").Find("Program Master", lookat:=xlWhole)
    
            If masterProgramNumberPosition Is Nothing Then
            
                Set masterProgramNumberPosition = Range("A1:IV1").Find("Prog/ID", lookat:=xlWhole)
                
                If masterProgramNumberPosition Is Nothing Then
                
                    MsgBox ("Master Program Number column was not found.  Exiting.")
                    
                    Exit Sub
                
                End If
            
            End If
        
        End If
    
    End If
    
    masterProgramNumberColumnLetter = Evaluate("MID(""" & masterProgramNumberPosition.Address & """,FIND(""$"",""" & masterProgramNumberPosition.Address & """)+1,FIND(""$"",""" & masterProgramNumberPosition.Address & """,2)-2)")

    masterProgramNumber = Range(masterProgramNumberColumnLetter & ActiveCell.Row)
    
    Set invoiceDatePosition = Range("A1:IV1").Find("Invoice Date", lookat:=xlWhole)
    
    If invoiceDatePosition Is Nothing Then
    
        Set invoiceDatePosition = Range("A1:IV1").Find("Run Date", lookat:=xlWhole)
    
        If invoiceDatePosition Is Nothing Then
        
            Set invoiceDatePosition = Range("A1:IV1").Find("Inv Date", lookat:=xlWhole)
        
            If invoiceDatePosition Is Nothing Then
            
                Set invoiceDatePosition = Range("A1:IV1").Find("U S F Invoice Date", lookat:=xlWhole)
            
            End If
                    
        End If
    
    End If
    
    invoiceDateColumnLetter = Evaluate("MID(""" & invoiceDatePosition.Address & """,FIND(""$"",""" & invoiceDatePosition.Address & """)+1,FIND(""$"",""" & invoiceDatePosition.Address & """,2)-2)")

    invoiceDate = Range(invoiceDateColumnLetter & ActiveCell.Row)
    
    oldestInvoiceDate = Evaluate("=MIN(IF($" & masterProgramNumberColumnLetter & ":$" & masterProgramNumberColumnLetter & "=" & masterProgramNumberColumnLetter & ActiveCell.Row & ",$" & invoiceDateColumnLetter & ":$" & invoiceDateColumnLetter & "))")
    
    Workbooks.Open ("G:\National SIS\Shared\Remedy\Macros\Pull Compliance Reports by Program.xlsm")
    
    Range("A4") = masterProgramNumber
    
    Range("D3") = oldestInvoiceDate
    
    Range("F3") = Evaluate("=TODAY()")
    
    Set iE = CreateObject("InternetExplorer.Application")
    
    iE.Visible = True
    
    iE.navigate "http://cas.usfood.com/CAS/supplierInvoiceSearch.jsp"
    
    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeSerial(0, 0, 1.5)

    Do Until iE.locationurl <> "http://cas.usfood.com/CAS/login.jsp"
    
        MsgBox ("Log in to CASIS then click Ok")
    
        While iE.busy
        DoEvents
        Wend
        
        Application.Wait Now + TimeSerial(0, 0, 1)
    
    Loop
    
    iE.document.all("masterNbr").Value = masterProgramNumber
    
    iE.document.all("from").Value = oldestInvoiceDate
    
    iE.document.all("doSearch").Click
    
    Application.Run "'G:\National SIS\Shared\Remedy\Macros\Pull Compliance Reports by Program.xlsm'!PullProgramData"
  
    
  
Set CreateReport = Nothing
invoiceDate = ""
invoiceDateColumnLetter = ""
Set invoiceDatePosition = Nothing
masterProgramNumber = ""
masterProgramNumberColumnLetter = ""
Set masterProgramNumberPosition = Nothing
    
End Sub

Sub subPivotForResolution()

Dim mainSheet As String
Dim mainSheetLastColumnNumber As Long
Dim mainSheetLastRow As Long
Dim pvtFld As PivotField
Dim pvtTbl As PivotTable
Dim resolutionPosition As Range
Dim runDatePosition As Range
Dim sISInvoiceNumberPosition As Range

Dim lastRow As Long

mainSheet = ""
mainSheetLastColumnNumber = Empty
mainSheetLastRow = Empty
Set pvtFld = Nothing
Set pvtTbl = Nothing
Set resolutionPosition = Nothing
Set runDatePosition = Nothing
Set sISInvoiceNumberPosition = Nothing


    Set resolutionPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole)
    
    If resolutionPosition Is Nothing Then
    
        MsgBox ("Resolution column was not found.  Exiting.")
        
        Exit Sub
    
    End If

    Set sISInvoiceNumberPosition = Range("A1:IV1").Find("SIS Inv #", lookat:=xlWhole)
    
    If sISInvoiceNumberPosition Is Nothing Then
    
        Set sISInvoiceNumberPosition = Range("A1:IV1").Find("SIS Invoice", lookat:=xlWhole)
    
    End If
    
    Set runDatePosition = Range("A1:IV1").Find("Run Date", lookat:=xlWhole)
    
    If runDatePosition Is Nothing Then
    
        Set runDatePosition = Range("A1:IV1").Find("Run DT", lookat:=xlWhole)
    
    End If

    mainSheet = ActiveSheet.Name
    mainSheetLastColumnNumber = Cells(1, Columns.Count).End(xlToLeft).Column
    mainSheetLastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Sheets.Add.Name = "ResolutionPivot"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        mainSheet & "!R1C1:R" & mainSheetLastRow & "C" & mainSheetLastColumnNumber, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=ActiveSheet.Name & "!R3C1", TableName:="ResolutionPivotTable", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets(ActiveSheet.Name).Select
    Cells(3, 1).Select
    
    With ActiveSheet.PivotTables("ResolutionPivotTable").PivotFields(sISInvoiceNumberPosition.Value)
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("ResolutionPivotTable").PivotFields(runDatePosition.Value)
        .Orientation = xlRowField
        .Position = 2
    End With
    Range("A4").Select
    With ActiveSheet.PivotTables("ResolutionPivotTable").PivotFields(sISInvoiceNumberPosition.Value)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
    Set pvtTbl = ActiveSheet.PivotTables("ResolutionPivotTable")
    
    With pvtTbl
        For Each pvtFld In .PivotFields
            pvtFld.Subtotals(1) = True
            pvtFld.Subtotals(1) = False
        Next pvtFld
    End With
    
    ActiveSheet.PivotTables("ResolutionPivotTable").AddDataField ActiveSheet.PivotTables("ResolutionPivotTable").PivotFields("Reb Amt"), "Sum of Reb Amt", xlSum
    
    ActiveSheet.PivotTables("ResolutionPivotTable").AddDataField ActiveSheet.PivotTables("ResolutionPivotTable").PivotFields("Resolution"), "Sum of Resolution", xlSum

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Sheets("ResolutionPivot").Select

    Range("E3").Value = "Interaction ID"

    Range("E4:E" & lastRow - 1).Formula = "=INDEX('C:\Users\d3q1700\Desktop\[Open.xlsm]Open'!$A:$A,MATCH(A4,'C:\Users\d3q1700\Desktop\[Open.xlsm]Open'!$E:$E,0))"

    Range("F3").Value = "Incident ID"

    Range("F4:F" & lastRow - 1).Formula = "=INDEX('C:\Users\d3q1700\Desktop\[Open.xlsm]Open'!$B:$B,MATCH(A4,'C:\Users\d3q1700\Desktop\[Open.xlsm]Open'!$E:$E,0))"

    Range("G3").Value = "Dispute Amount"
    
    Range("G4:G" & lastRow - 1).Formula = "=INDEX('C:\Users\d3q1700\Desktop\[Open.xlsm]Open'!$G:$G,MATCH(A4,'C:\Users\d3q1700\Desktop\[Open.xlsm]Open'!$E:$E,0))"
    
    Range("H3").Value = "Special Bill ID"

    Columns("D:D").NumberFormat = "0.00"

mainSheet = ""
mainSheetLastColumnNumber = Empty
mainSheetLastRow = Empty
Set pvtFld = Nothing
Set pvtTbl = Nothing
Set resolutionPosition = Nothing
Set runDatePosition = Nothing
Set sISInvoiceNumberPosition = Nothing

End Sub

Sub subEMailReversalToLSISSalesAudits()

Dim OutApp As Object
Dim OutMail As Object
Dim masterCmpnyVndrNamePosition As Range
Dim masterCmpnyVndrNameColumnLetter As String
Dim masterCmpnyVndrName As String
Dim masterVndrNumberPosition As Range
Dim masterVndrNumberColumnLetter As String
Dim masterVndrNumber As String
Dim programMasterPosition As Range
Dim programMasterColumnLetter As String
Dim programMaster As String

    Set masterCmpnyVndrNamePosition = Range("A1:IV1").Find("Master Cmpny Vndr Name", lookat:=xlWhole)
    
    If masterCmpnyVndrNamePosition Is Nothing Then
    
        MsgBox ("""Master Cmpny Vndr Name"" was not found.  Activesheet does not seem to be a compliance report.  Exiting.")
        
        Exit Sub
    
    End If
    
    masterCmpnyVndrNameColumnLetter = Evaluate("MID(""" & masterCmpnyVndrNamePosition.Address & """,FIND(""$"",""" & masterCmpnyVndrNamePosition.Address & """)+1,FIND(""$"",""" & masterCmpnyVndrNamePosition.Address & """,2)-2)")
    
    masterCmpnyVndrName = Range(masterCmpnyVndrNameColumnLetter & 2)

    Set masterVndrNumberPosition = Range("A1:IV1").Find("Master Vndr #", lookat:=xlWhole)
    
    masterVndrNumberColumnLetter = Evaluate("MID(""" & masterVndrNumberPosition.Address & """,FIND(""$"",""" & masterVndrNumberPosition.Address & """)+1,FIND(""$"",""" & masterVndrNumberPosition.Address & """,2)-2)")
    
    masterVndrNumber = Range(masterVndrNumberColumnLetter & 2)

    Set programMasterPosition = Range("A1:IV1").Find("Program Master", lookat:=xlWhole)
    
    programMasterColumnLetter = Evaluate("MID(""" & programMasterPosition.Address & """,FIND(""$"",""" & programMasterPosition.Address & """)+1,FIND(""$"",""" & programMasterPosition.Address & """,2)-2)")
    
    programMaster = Range(programMasterColumnLetter & 2)

    Set OutApp = CreateObject("Outlook.Application")
    
    Set OutMail = OutApp.CreateItemFromTemplate("C:\Users\" & Environ("username") & "\Desktop\Invoice Reversal Template.oft")
    
    OutMail.display

    OutMail.To = "LSIS_Sales_Audits@usfoods.com"

    OutMail.Subject = "Vendor " & masterCmpnyVndrName & ", Vendor # " & masterVndrNumber & ", Master # " & programMaster & ", Invoice # (If multiple, ""multiple invoices"")"

    

    OutMail.htmlbody = Replace(OutMail.htmlbody, "CAS Master Program", "CAS Master Program: " & programMaster)

    'OutMail.Subject = Replace(OutMail.Subject, "Vendor", "VENDOR")

End Sub

Sub subServicePortalLocalContractsRequest()

Dim branchCodePosition As Range
Dim branchCodeColumnLetter As String
Dim branchCode As String
Dim programMasterNamePosition As Range
Dim programMasterNameColumnLetter As String
Dim programMasterName As String
Dim masterCmpnyVndrNamePosition As Range
Dim masterCmpnyVndrNameColumnLetter As String
Dim masterCmpnyVndrName As String
Dim mainSheet As Worksheet
Dim newSheet As Worksheet
Dim productNumberPosition As Range
Dim productNumberColumnLetter As String
Dim lastRow As Long
Dim productNumberlastRow As Long
Dim productsList As String
Dim shipDatePosition As Range
Dim shipDateColumnLetter As String
Dim maximumDate As Date
Dim minimumDate As Date
Dim contractNumberPosition As Range
Dim contractNumberColumnLetter As String
Dim contractNumber As String

Dim iE As Object

    Set branchCodePosition = Range("A1:IV1").Find("Branch Cd", lookat:=xlWhole)
    
    branchCodeColumnLetter = Evaluate("MID(""" & branchCodePosition.Address & """,FIND(""$"",""" & branchCodePosition.Address & """)+1,FIND(""$"",""" & branchCodePosition.Address & """,2)-2)")
    
    branchCode = Range(branchCodeColumnLetter & 2).Value

    Set programMasterNamePosition = Range("A1:IV1").Find("Program Master Name", lookat:=xlWhole)
    
    programMasterNameColumnLetter = Evaluate("MID(""" & programMasterNamePosition.Address & """,FIND(""$"",""" & programMasterNamePosition.Address & """)+1,FIND(""$"",""" & programMasterNamePosition.Address & """,2)-2)")
    
    programMasterName = Range(programMasterNameColumnLetter & 2)

    Set masterCmpnyVndrNamePosition = Range("A1:IV1").Find("Master Cmpny Vndr Name", lookat:=xlWhole)
    
    masterCmpnyVndrNameColumnLetter = Evaluate("MID(""" & masterCmpnyVndrNamePosition.Address & """,FIND(""$"",""" & masterCmpnyVndrNamePosition.Address & """)+1,FIND(""$"",""" & masterCmpnyVndrNamePosition.Address & """,2)-2)")
    
    masterCmpnyVndrName = Range(masterCmpnyVndrNameColumnLetter & 2)

    Set productNumberPosition = Range("A1:IV1").Find("USF Prod #", lookat:=xlWhole)
    
    productNumberColumnLetter = Evaluate("MID(""" & productNumberPosition.Address & """,FIND(""$"",""" & productNumberPosition.Address & """)+1,FIND(""$"",""" & productNumberPosition.Address & """,2)-2)")
    
    Columns(productNumberColumnLetter & ":" & productNumberColumnLetter).Select
    
    Selection.Copy
    
    Set mainSheet = ActiveSheet

    Set shipDatePosition = Range("A1:IV1").Find("Ship Date", lookat:=xlWhole)
    
    shipDateColumnLetter = Evaluate("MID(""" & shipDatePosition.Address & """,FIND(""$"",""" & shipDatePosition.Address & """)+1,FIND(""$"",""" & shipDatePosition.Address & """,2)-2)")

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    minimumDate = Evaluate("MIN(" & shipDateColumnLetter & "1:" & shipDateColumnLetter & lastRow & ")")
    
    maximumDate = Evaluate("MAX(" & shipDateColumnLetter & "1:" & shipDateColumnLetter & lastRow & ")")

    Set contractNumberPosition = Range("A1:IV1").Find("Cntrct #", lookat:=xlWhole)
    
    contractNumberColumnLetter = Evaluate("MID(""" & contractNumberPosition.Address & """,FIND(""$"",""" & contractNumberPosition.Address & """)+1,FIND(""$"",""" & contractNumberPosition.Address & """,2)-2)")
    
    contractNumber = Range(contractNumberColumnLetter & 2)

    On Error Resume Next

    Sheets.Add.Name = "Products"
    
    Set newSheet = ActiveSheet
    
    newSheet.Paste
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Application.CutCopyMode = False
    
    newSheet.Range("$A$1:$A$" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("$A$1:$A$" & lastRow).Sort key1:=Range("A2"), order1:=xlAscending, Header:=xlYes
    
    productNumberlastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Select Case productNumberlastRow
        
        Case 2
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
        
        Case 3
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("B3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"", ""&RC[-1]"
        
        Case Is > 3
            Range("B2").Select
            ActiveCell.FormulaR1C1 = "=RC[-1]"
            Range("B3").Select
            ActiveCell.FormulaR1C1 = "=R[-1]C&"", ""&RC[-1]"
            Range("B3").Select
            Selection.AutoFill Destination:=Range("B3:B" & productNumberlastRow)
    
    End Select
    
    productsList = Cells(Rows.Count, 2).End(xlUp).Value

    Range("A2:A" & lastRow).Delete

    Range("B1") = productsList

    Range("A2") = "Branch Code"
    
    Range("B2") = branchCode
    
    Range("A3") = "Customer"
    
    Range("B3") = programMasterName
    
    Range("A4") = "Vendor"
    
    Range("B4") = masterCmpnyVndrName
    
    Range("A5") = "Start Date"
    
    Range("B5") = minimumDate
    
    Range("A6") = "End Date"
    
    Range("B6") = maximumDate
    
    Range("A7") = "Contract Number"
    
    Range("B7") = contractNumber
    
    Range("A8") = "Title"
    
    Range("B8") = masterCmpnyVndrName & " And " & programMasterName & " contract needed, please."

    Range("A9") = "Description"
    
    Range("B9") = "Hello, can you please send the contract for PRIME contract " & contractNumber & " and product/s " & productsList & " for period " & minimumDate & "-" & maximumDate & "?   If this contract was auto-renewed, can you please provide the auto-renewal notification?  Thank you,"

    Set iE = CreateObject("InternetExplorer.Application")
    
    iE.navigate "http://servicecatalog.main.usfood.com/SRC9.41/secure/main.jsp"
    
    iE.Visible = True

End Sub

Sub subStatusReport()

    Application.Run "'G:\National SIS\Shared\Remedy\Macros\Status Report Macro.xlsm'!subStatusReport"
   
'Dim casesCount As Integer
'Dim closedByColumnLetter As String
'Dim closedByPosition As Range
'Dim closeTimeColumnLetter As String
'Dim closeTimePosition As Range
'Dim companyFullNameColumnLetter As String
'Dim companyFullNamePosition As Range
'Dim currentVendorDisputeCount As Long
'Dim currentVendorDisputedSum As Double
'Dim currentVendorName As String
'Dim currentVendorRepaidSum As Double
'Dim currentWeek As Long
'Dim dateForCurrentWeek As Date
'Dim deniedCount As Variant
'Dim disputeAmountColumnLetter As String
'Dim disputeAmountPosition As Range
'Dim firstDayOfCurrentWeek As Date
'Dim firstDayOfCurrentWeekString As String
'Dim fW As Integer
'Dim lastColumnLetter As String
'Dim lastColumnPosition As Range
'Dim lastDayOfCurrentWeek As Date
'Dim lastDayOfCurrentWeekString As String
'Dim lastRow As Variant
'Dim OutApp As Object
'Dim OutMail As Object
'Dim repaymentsCount As Long
'Dim something As Long
'Dim sumOfDenied As Double
'Dim sumOfDisputedCases As Double
'Dim sumOfRepayments As Double
'Dim userName As String
'Dim uSFResolutionAmountColumnLetter As String
'Dim uSFResolutionAmountPosition As Range
'Dim weekColumnLetter As String
'Dim weekColumnPosition As Range
'Dim weekPosition As Range
'
'Dim statusReport As Worksheet
'Dim currentTime As String
'Dim currentTimeLength As Long
'Dim currentTimeFirstSpace As Long
'Dim fso As Object
'Dim textFileOut As Object
'
'casesCount = Empty
'closedByColumnLetter = ""
'Set closedByPosition = Nothing
'closeTimeColumnLetter = ""
'Set closeTimePosition = Nothing
'companyFullNameColumnLetter = ""
'Set companyFullNamePosition = Nothing
'currentVendorDisputeCount = Empty
'currentVendorDisputedSum = Empty
'currentVendorName = ""
'currentVendorRepaidSum = Empty
'currentWeek = Empty
'dateForCurrentWeek = Empty
'Set deniedCount = Nothing
'disputeAmountColumnLetter = ""
'Set disputeAmountPosition = Nothing
'firstDayOfCurrentWeek = Empty
'firstDayOfCurrentWeekString = ""
'fW = Empty
'lastColumnLetter = ""
'Set lastColumnPosition = Nothing
'lastDayOfCurrentWeek = Empty
'lastDayOfCurrentWeekString = ""
'Set lastRow = Nothing
'Set OutApp = Nothing
'Set OutMail = Nothing
'repaymentsCount = Empty
'something = Empty
'sumOfDenied = Empty
'sumOfDisputedCases = Empty
'sumOfRepayments = Empty
'userName = ""
'uSFResolutionAmountColumnLetter = ""
'Set uSFResolutionAmountPosition = Nothing
'weekColumnLetter = ""
'Set weekColumnPosition = Nothing
'Set weekPosition = Nothing
'
'    Set disputeAmountPosition = Range("A1:IV1").Find("U S F Dispute Amount", lookat:=xlWhole)
'
'    If disputeAmountPosition Is Nothing Then
'
'        Set disputeAmountPosition = Range("A1:IV1").Find("Dispute Amount", lookat:=xlWhole)
'
'        If disputeAmountPosition Is Nothing Then
'
'            MsgBox ("U S F Dispute Amount was not found.  Exiting")
'
'            Exit Sub
'
'        End If
'
'    End If
'
'    Set statusReport = ActiveSheet
'
'    disputeAmountColumnLetter = Evaluate("MID(""" & disputeAmountPosition.Address & """,FIND(""$"",""" & disputeAmountPosition.Address & """)+1,FIND(""$"",""" & disputeAmountPosition.Address & """,2)-2)")
'
'    Set lastColumnPosition = Cells(1, Columns.Count).End(xlToLeft)
'
'    lastColumnLetter = Evaluate("MID(""" & lastColumnPosition.Address & """,FIND(""$"",""" & lastColumnPosition.Address & """)+1,FIND(""$"",""" & lastColumnPosition.Address & """,2)-2)")
'
'    Set weekColumnPosition = Range("A1:IV1").Find("Week", lookat:=xlWhole)
'
'    If weekColumnPosition Is Nothing Then
'
'        Set weekColumnPosition = lastColumnPosition.Offset(, 1)
'
'        weekColumnPosition.Value = "Week"
'
'        Set lastColumnPosition = Cells(1, Columns.Count).End(xlToLeft)
'
'        lastColumnLetter = Evaluate("MID(""" & lastColumnPosition.Address & """,FIND(""$"",""" & lastColumnPosition.Address & """)+1,FIND(""$"",""" & lastColumnPosition.Address & """,2)-2)")
'
'    End If
'
'    weekColumnLetter = Evaluate("MID(""" & weekColumnPosition.Address & """,FIND(""$"",""" & weekColumnPosition.Address & """)+1,FIND(""$"",""" & weekColumnPosition.Address & """,2)-2)")
'
'    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
'
'    Set companyFullNamePosition = Range("A1:IV1").Find("U S F Company Full Name", lookat:=xlWhole)
'
'    If companyFullNamePosition Is Nothing Then
'
'        Set companyFullNamePosition = Range("A1:IV1").Find("Company Full Name", lookat:=xlWhole)
'
'    End If
'
'    companyFullNameColumnLetter = Evaluate("MID(""" & companyFullNamePosition.Address & """,FIND(""$"",""" & companyFullNamePosition.Address & """)+1,FIND(""$"",""" & companyFullNamePosition.Address & """,2)-2)")
'
'    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
'
'    ActiveSheet.Range("A1:" & lastColumnLetter & lastRow).AutoFilter
'
'    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
'        (companyFullNameColumnLetter & "1:" & companyFullNameColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'        xlSortTextAsNumbers
'    With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    Set closeTimePosition = Range("A1:IV1").Find("Close Time", lookat:=xlWhole)
'
'    closeTimeColumnLetter = Evaluate("MID(""" & closeTimePosition.Address & """,FIND(""$"",""" & closeTimePosition.Address & """)+1,FIND(""$"",""" & closeTimePosition.Address & """,2)-2)")
'
'    Range(weekColumnLetter & "2:" & weekColumnLetter & lastRow).Formula = "=WEEKNUM(" & closeTimeColumnLetter & "2)"
'
'    Range(weekColumnLetter & "2:" & weekColumnLetter & lastRow).Value = Range(weekColumnLetter & "2:" & weekColumnLetter & lastRow).Value
'
'    Set closedByPosition = Range("A1:IV1").Find("Closed By", lookat:=xlWhole)
'
'    closedByColumnLetter = Evaluate("MID(""" & closedByPosition.Address & """,FIND(""$"",""" & closedByPosition.Address & """)+1,FIND(""$"",""" & closedByPosition.Address & """,2)-2)")
'
'    Set weekPosition = Range("A1:IV1").Find("Week", lookat:=xlWhole)
'
'    weekColumnLetter = Evaluate("MID(""" & weekPosition.Address & """,FIND(""$"",""" & weekPosition.Address & """)+1,FIND(""$"",""" & weekPosition.Address & """,2)-2)")
'
'    Set uSFResolutionAmountPosition = Range("A1:IV1").Find("U S F Resolution Amount", lookat:=xlWhole)
'
'    uSFResolutionAmountColumnLetter = Evaluate("MID(""" & uSFResolutionAmountPosition.Address & """,FIND(""$"",""" & uSFResolutionAmountPosition.Address & """)+1,FIND(""$"",""" & uSFResolutionAmountPosition.Address & """,2)-2)")
'
'    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
'
'    userName = UCase(Environ("username"))
'
'    userName = Application.InputBox(userName & " will be used to pull the data.  If you need the data for a different user, enter it below.", Default:=userName, Type:=2)
'
'    userName = UCase(userName)
'
'    dateForCurrentWeek = Application.InputBox(Date & " will be used to pull the data.  If you need the data for a different week, enter a date that falls in the week below.", Default:=Date, Type:=2)
'
'    currentWeek = CInt(Format(dateForCurrentWeek, "ww", fW))
'
'    casesCount = Evaluate("COUNTIFS(" & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & ")")
'
'    sumOfDisputedCases = Evaluate("SUMIFS(" & disputeAmountColumnLetter & "1:" & disputeAmountColumnLetter & lastRow & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & ")")
'
'    repaymentsCount = Evaluate("COUNTIFS(" & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ","">0""," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """)")
'
'    repaymentsCount = Evaluate("COUNTIFS(" & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ",""<0""," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """)") + repaymentsCount
'
'    sumOfRepayments = Evaluate("SUMIFS(" & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & "," & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ",""<0""," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """)")
'
'    sumOfRepayments = sumOfRepayments * -1
'
'    sumOfRepayments = Evaluate("SUMIFS(" & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & "," & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ","">0""," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """)") + sumOfRepayments
'
'    deniedCount = Evaluate("COUNTIFS(" & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ",""""," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """)+COUNTIFS(" & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ",0," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """)")
'
'    sumOfDenied = Evaluate("SUMIFS(" & disputeAmountColumnLetter & "1:" & disputeAmountColumnLetter & lastRow & "," & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ",""0""," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """)+SUMIFS(" & disputeAmountColumnLetter & "1:" & disputeAmountColumnLetter & lastRow & "," & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ",""""," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & "," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """)")
'
'    Set OutApp = CreateObject("Outlook.Application")
'
'    Set OutMail = OutApp.CreateItem(0)
'
'    OutMail.display
'
'    Cells(2, ActiveCell.Column).Select
'
'    Do Until ActiveCell.Row > lastRow
'
'        currentVendorName = Range(companyFullNameColumnLetter & ActiveCell.Row)
'
'        currentVendorDisputeCount = Evaluate("COUNTIFS(" & companyFullNameColumnLetter & "1:" & companyFullNameColumnLetter & lastRow & ",""" & currentVendorName & """," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & ")")
'
'        If currentVendorDisputeCount > 0 Then
'
'            currentVendorDisputedSum = Evaluate("SUMIFS(" & disputeAmountColumnLetter & "1:" & disputeAmountColumnLetter & lastRow & "," & companyFullNameColumnLetter & "1:" & companyFullNameColumnLetter & lastRow & ",""" & currentVendorName & """," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & ")")
'
'            currentVendorRepaidSum = Evaluate("SUMIFS(" & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & "," & uSFResolutionAmountColumnLetter & "1:" & uSFResolutionAmountColumnLetter & lastRow & ","">0""," & companyFullNameColumnLetter & "1:" & companyFullNameColumnLetter & lastRow & ",""" & currentVendorName & """," & closedByColumnLetter & "1:" & closedByColumnLetter & lastRow & ",""" & userName & """," & weekColumnLetter & "1:" & weekColumnLetter & lastRow & "," & currentWeek & ")")
'
'            OutMail.htmlbody = currentVendorName & ": $" & currentVendorDisputedSum & " Disputed, $" & currentVendorRepaidSum & " repaid" & "<br>" & OutMail.htmlbody
'        '"<font face=""Calibri""><span style=""font-size: .90em"">" &
'        End If
'
'        Do Until currentVendorName <> Range(companyFullNameColumnLetter & ActiveCell.Row)
'
'            ActiveCell.Offset(1).Select
'
'        Loop
'
'    Loop
'
'    Range("A1").Select
'
'    OutMail.htmlbody = "<font face=""Calibri""><span style=""font-size: .90em"">" & "Cases Closed: " & casesCount & "/$" & sumOfDisputedCases & "<br>" _
'    & "Repaid: " & repaymentsCount & "/$" & sumOfRepayments & "<br>" _
'    & "Denied: " & deniedCount & "/$" & sumOfDenied & "<br><br>" _
'    & OutMail.htmlbody
'
'    OutMail.to = "Stephanie.Go@usfoods.com"
'
'    firstDayOfCurrentWeek = dateForCurrentWeek - (Weekday(dateForCurrentWeek) - 1)
'
'    firstDayOfCurrentWeekString = Format(firstDayOfCurrentWeek, "MM/DD/YY")
'
'    lastDayOfCurrentWeek = dateForCurrentWeek + ((Weekday(dateForCurrentWeek) - 7) * -1)
'
'    lastDayOfCurrentWeekString = Format(lastDayOfCurrentWeek, "MM/DD/YY")
'
'    OutMail.Subject = "Weekly HPSM Status Report Update " & firstDayOfCurrentWeekString & " - " & lastDayOfCurrentWeekString
'
'' Usage section
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    currentTime = Now
'
'    currentTimeLength = Len(currentTime)
'
'    currentTimeFirstSpace = Evaluate("=FIND("" "",""" & currentTime & """)")
'
'    currentTime = Right(currentTime, currentTimeLength - currentTimeFirstSpace)
'
'    currentTime = Evaluate("SUBSTITUTE(""" & currentTime & ""","":"","""")")
'
'    Set textFileOut = fso.CreateTextFile("G:\National SIS\Shared\Remedy\Macros\Usage Saves\" & Left(ThisWorkbook.Name, InStr(ThisWorkbook.Name, ".xl") - 1) & " UsageSave " & Format(Now, "MMDDYY") & " " & currentTime & ".txt", True, True)
'
'    textFileOut.Write Now & "|" & UCase(Environ("username")) & "|" & 0.25
'
'    textFileOut.Close
'
'casesCount = Empty
'closedByColumnLetter = ""
'Set closedByPosition = Nothing
'closeTimeColumnLetter = ""
'Set closeTimePosition = Nothing
'companyFullNameColumnLetter = ""
'Set companyFullNamePosition = Nothing
'currentVendorDisputeCount = Empty
'currentVendorDisputedSum = Empty
'currentVendorName = ""
'currentVendorRepaidSum = Empty
'currentWeek = Empty
'dateForCurrentWeek = Empty
'Set deniedCount = Nothing
'disputeAmountColumnLetter = ""
'Set disputeAmountPosition = Nothing
'firstDayOfCurrentWeek = Empty
'firstDayOfCurrentWeekString = ""
'fW = Empty
'lastColumnLetter = ""
'Set lastColumnPosition = Nothing
'lastDayOfCurrentWeek = Empty
'lastDayOfCurrentWeekString = ""
'Set lastRow = Nothing
'Set OutApp = Nothing
'Set OutMail = Nothing
'repaymentsCount = Empty
'something = Empty
'sumOfDenied = Empty
'sumOfDisputedCases = Empty
'sumOfRepayments = Empty
'userName = ""
'uSFResolutionAmountColumnLetter = ""
'Set uSFResolutionAmountPosition = Nothing
'weekColumnLetter = ""
'Set weekColumnPosition = Nothing
'Set weekPosition = Nothing
   
End Sub

Sub subConvertToText()

Dim lastRow As Long

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Do Until ActiveCell.Row > lastRow

    ActiveCell.Value = "'" & ActiveCell.Value

    ActiveCell.Offset(1).Select

Loop

End Sub

Sub subRecalculateCustomerPickUpAllowanceCatchWtY()

    Application.Run "'G:\National SIS\Shared\Remedy\Macros\Freight Billed Above Customer Pickup Allowance Macro.xlsm'!subRecalculateCustomerPickUpAllowanceCatchWtY"

'Dim getPosition As Range
'Dim customerPickupAllowancePerUnitColumnLetter As String
'Dim cWColumnLetter As String
'Dim lastRow As Long
'Dim lastColumnColumnLetter As String
'Dim currentCatchWeightYBeginningRow As Long
'Dim currentCatchWeightYAndNotZeroBeginningRow As Long
'Dim vendorProductPricePerUnitColumnLetter As String
'Dim freightRatePerUnitColumnLetter As String
'Dim divisionDeliveredPricePerUnitColumnLetter As String
'Dim prePayAndAddFreightPerUnitColumnLetter As String
'Dim runDateColumnLetter As String
'Dim useRecalculationColumnYesNo As Integer
'Dim recalculationColumnLetter As String
'Dim contBasisColumnLetter As String
'Dim poundsColumnLetter As String
'Dim rebAmtColumnLetter As String
'Dim useResolutionColumnYesNo As Integer
'Dim resolutionColumnLetter As String
'Dim reasonColumnLetter As String
'
'    Set getPosition = Range("A1:IV1").Find("CUSTOMER PICKUP ALLOWANCE PER UNIT", lookat:=xlWhole)
'
'    If getPosition Is Nothing Then
'
'        MsgBox ("CUSTOMER PICKUP ALLOWANCE PER UNIT field was not found.  Exiting")
'
'        Exit Sub
'
'    End If
'
'    customerPickupAllowancePerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("CW", lookat:=xlWhole)
'
'    cWColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("VENDOR PRODUCT PRICE PER UNIT", lookat:=xlWhole)
'
'    vendorProductPricePerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("FREIGHT RATE PER UNIT", lookat:=xlWhole)
'
'    freightRatePerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("DIVISION DELIVERED PRICE PER UNIT", lookat:=xlWhole)
'
'    divisionDeliveredPricePerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("PREPAY AND ADD FREIGHT PER UNIT", lookat:=xlWhole)
'
'    prePayAndAddFreightPerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Run Date", lookat:=xlWhole)
'
'    runDateColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Cont Basis", lookat:=xlWhole)
'
'    contBasisColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Pounds", lookat:=xlWhole)
'
'    poundsColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Reb Amt", lookat:=xlWhole)
'
'    rebAmtColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Recalculation", lookat:=xlWhole)
'
'    If Not getPosition Is Nothing Then
'
'        useRecalculationColumnYesNo = MsgBox("Recalculation column found per cell " & getPosition.Address & ".  Would you like to use the existing one?  A new one will be created if no.", vbYesNo)
'
'        If useRecalculationColumnYesNo = 6 Then
'
'            recalculationColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        Else
'
'            Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Recalculation"
'
'            Set getPosition = Range("A1:IV1").Find("Recalculation", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'            recalculationColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        End If
'
'    Else
'
'        Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Recalculation"
'
'        Set getPosition = Range("A1:IV1").Find("Recalculation", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'        recalculationColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    End If
'
'    Set getPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole)
'
'    If Not getPosition Is Nothing Then
'
'        useResolutionColumnYesNo = MsgBox("Resolution column found per cell " & getPosition.Address & ".  Would you like to use the existing one?  A new one will be created if no.", vbYesNo)
'
'        If useResolutionColumnYesNo = 6 Then
'
'            resolutionColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        Else
'
'            Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Resolution"
'
'            Set getPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'            resolutionColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        End If
'
'    Else
'
'        Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Resolution"
'
'        Set getPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'        resolutionColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    End If
'
'    Set getPosition = Range("A1:IV1").Find("Reason", lookat:=xlWhole)
'
'    If Not getPosition Is Nothing Then
'
'        useReasonColumnYesNo = MsgBox("Reason column found per cell " & getPosition.Address & ".  Would you like to use the existing one?  A new one will be created if no.", vbYesNo)
'
'        If useReasonColumnYesNo = 6 Then
'
'            reasonColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        Else
'
'            Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Reason"
'
'            Set getPosition = Range("A1:IV1").Find("Reason", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'            reasonColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        End If
'
'    Else
'
'        Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Reason"
'
'        Set getPosition = Range("A1:IV1").Find("Reason", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'        reasonColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    End If
'
'    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
'
'    Set getPosition = Cells(1, Columns.count).End(xlToLeft)
'
'    lastColumnColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    lastRow = Cells(Rows.count, 1).End(xlUp).Row
'
'    Range("A1:" & lastColumnColumnLetter & lastRow).AutoFilter
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
'        (cWColumnLetter & "2:" & cWColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
'        (customerPickupAllowancePerUnitColumnLetter & "2:" & customerPickupAllowancePerUnitColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    With ActiveSheet.AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    currentCatchWeightYBeginningRow = Range(cWColumnLetter & "1:" & cWColumnLetter & lastRow).Find("Y", lookat:=xlWhole, searchdirection:=xlNext).Row
'
'    currentCatchWeightYAndNotZeroBeginningRow = Evaluate("=MATCH(FALSE," & customerPickupAllowancePerUnitColumnLetter & currentCatchWeightYBeginningRow & ":" & customerPickupAllowancePerUnitColumnLetter & lastRow & "=0,0)")
'
'    currentCatchWeightYAndNotZeroBeginningRow = currentCatchWeightYAndNotZeroBeginningRow + currentCatchWeightYBeginningRow - 1
'
'    Range(customerPickupAllowancePerUnitColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & ":" & customerPickupAllowancePerUnitColumnLetter & lastRow).Formula = "=" & vendorProductPricePerUnitColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & "+" & prePayAndAddFreightPerUnitColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & "+" & freightRatePerUnitColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & "-" & divisionDeliveredPricePerUnitColumnLetter & currentCatchWeightYAndNotZeroBeginningRow
'
'    Range(customerPickupAllowancePerUnitColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & ":" & customerPickupAllowancePerUnitColumnLetter & lastRow).Value = Range(customerPickupAllowancePerUnitColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & ":" & customerPickupAllowancePerUnitColumnLetter & lastRow).Value
'
'    Range(recalculationColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & ":" & recalculationColumnLetter & lastRow).Formula = "=(" & vendorProductPricePerUnitColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & "-" & contBasisColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & ")*" & poundsColumnLetter & currentCatchWeightYAndNotZeroBeginningRow
'
'    Range(resolutionColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & ":" & resolutionColumnLetter & lastRow).Formula = "=" & recalculationColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & "-" & rebAmtColumnLetter & currentCatchWeightYAndNotZeroBeginningRow
'
'    Range(reasonColumnLetter & currentCatchWeightYAndNotZeroBeginningRow & ":" & reasonColumnLetter & lastRow).Value = "Freight billed above customer pickup allowance is to be repaid per FreightAllowanceRule"
'
'    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
'
'    Set getPosition = Cells(1, Columns.count).End(xlToLeft)
'
'    lastColumnColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    lastRow = Cells(Rows.count, 1).End(xlUp).Row
'
'    Range("A1:" & lastColumnColumnLetter & lastRow).AutoFilter
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
'        (runDateColumnLetter & "2:" & runDateColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    With ActiveSheet.AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    On Error Resume Next
'
'    Set freightAllowanceRuleSheet = ActiveWorkbook.Sheets("FreightAllowanceRule")
'
'    If Err.Number > 0 Then
'
'        Set mainSheet = ActiveSheet
'
'        Err.Clear
'
'        Sheets.Add After:=ActiveSheet
'
'        Set freightAllowanceRuleSheet = ActiveSheet
'
'        freightAllowanceRuleSheet.Name = "FreightAllowanceRule"
'        ActiveSheet.OLEObjects.Add(Filename:= _
'            "http://sharepoint.usfood.com/sites/Finance/CASS/HPSM/HPSM%20Team/Shared%20Documents/Freight%20allowance%20rule.msg" _
'            , link:=False, DisplayAsIcon:=False).Select
'
'        mainSheet.Select
'
'    End If
'
'    MsgBox ("This macro does not consider if the delivered price is less than the contract price.")

End Sub

Sub subRecalculateCustomerPickUpAllowanceCatchWtN()

    Application.Run "'G:\National SIS\Shared\Remedy\Macros\Freight Billed Above Customer Pickup Allowance Macro.xlsm'!subRecalculateCustomerPickUpAllowanceCatchWtN"
'
'Dim getPosition As Range
'Dim customerPickupAllowancePerUnitColumnLetter As String
'Dim cWColumnLetter As String
'Dim lastRow As Long
'Dim lastColumnColumnLetter As String
'Dim currentCatchWeightNBeginningRow As Long
'Dim currentCatchWeightNAndNotZeroBeginningRow As Long
'Dim vendorProductPricePerUnitColumnLetter As String
'Dim freightRatePerUnitColumnLetter As String
'Dim divisionDeliveredPricePerUnitColumnLetter As String
'Dim prePayAndAddFreightPerUnitColumnLetter As String
'Dim runDateColumnLetter As String
'Dim useRecalculationColumnYesNo As Integer
'Dim recalculationColumnLetter As String
'Dim contBasisColumnLetter As String
'Dim poundsColumnLetter As String
'Dim rebAmtColumnLetter As String
'Dim useResolutionColumnYesNo As Integer
'Dim resolutionColumnLetter As String
'Dim casesColumnLetter As String
'Dim freightAllowanceRuleSheet As Worksheet
'Dim mainSheet As Worksheet
'
'    Set getPosition = Range("A1:IV1").Find("CUSTOMER PICKUP ALLOWANCE PER UNIT", lookat:=xlWhole)
'
'    If getPosition Is Nothing Then
'
'        MsgBox ("CUSTOMER PICKUP ALLOWANCE PER UNIT field was not found.  Exiting")
'
'        Exit Sub
'
'    End If
'
'    customerPickupAllowancePerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("CW", lookat:=xlWhole)
'
'    cWColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("VENDOR PRODUCT PRICE PER UNIT", lookat:=xlWhole)
'
'    vendorProductPricePerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("FREIGHT RATE PER UNIT", lookat:=xlWhole)
'
'    freightRatePerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("DIVISION DELIVERED PRICE PER UNIT", lookat:=xlWhole)
'
'    divisionDeliveredPricePerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("PREPAY AND ADD FREIGHT PER UNIT", lookat:=xlWhole)
'
'    prePayAndAddFreightPerUnitColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Run Date", lookat:=xlWhole)
'
'    runDateColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Cont Basis", lookat:=xlWhole)
'
'    contBasisColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Pounds", lookat:=xlWhole)
'
'    poundsColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Reb Amt", lookat:=xlWhole)
'
'    rebAmtColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    Set getPosition = Range("A1:IV1").Find("Cases", lookat:=xlWhole)
'
'    casesColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'' Optional Fields Below
'
'    Set getPosition = Range("A1:IV1").Find("Recalculation", lookat:=xlWhole)
'
'    If Not getPosition Is Nothing Then
'
'        useRecalculationColumnYesNo = MsgBox("Recalculation column found per cell " & getPosition.Address & ".  Would you like to use the existing one?  A new one will be created if no.", vbYesNo)
'
'        If useRecalculationColumnYesNo = 6 Then
'
'            recalculationColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        Else
'
'            Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Recalculation"
'
'            Set getPosition = Range("A1:IV1").Find("Recalculation", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'            recalculationColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        End If
'
'    Else
'
'        Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Recalculation"
'
'        Set getPosition = Range("A1:IV1").Find("Recalculation", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'        recalculationColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    End If
'
'    Set getPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole)
'
'    If Not getPosition Is Nothing Then
'
'        useResolutionColumnYesNo = MsgBox("Resolution column found per cell " & getPosition.Address & ".  Would you like to use the existing one?  A new one will be created if no.", vbYesNo)
'
'        If useResolutionColumnYesNo = 6 Then
'
'            resolutionColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        Else
'
'            Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Resolution"
'
'            Set getPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'            resolutionColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        End If
'
'    Else
'
'        Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Resolution"
'
'        Set getPosition = Range("A1:IV1").Find("Resolution", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'        resolutionColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    End If
'
'    Set getPosition = Range("A1:IV1").Find("Reason", lookat:=xlWhole)
'
'    If Not getPosition Is Nothing Then
'
'        useReasonColumnYesNo = MsgBox("Reason column found per cell " & getPosition.Address & ".  Would you like to use the existing one?  A new one will be created if no.", vbYesNo)
'
'        If useReasonColumnYesNo = 6 Then
'
'            reasonColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        Else
'
'            Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Reason"
'
'            Set getPosition = Range("A1:IV1").Find("Reason", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'            reasonColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'        End If
'
'    Else
'
'        Cells(1, Columns.count).End(xlToLeft).Offset(, 1) = "Reason"
'
'        Set getPosition = Range("A1:IV1").Find("Reason", lookat:=xlWhole, searchdirection:=xlPrevious)
'
'        reasonColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    End If
'
'    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
'
'    Set getPosition = Cells(1, Columns.count).End(xlToLeft)
'
'    lastColumnColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    lastRow = Cells(Rows.count, 1).End(xlUp).Row
'
'    Range("A1:" & lastColumnColumnLetter & lastRow).AutoFilter
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
'        (cWColumnLetter & "2:" & cWColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'        xlSortNormal
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
'        (customerPickupAllowancePerUnitColumnLetter & "2:" & customerPickupAllowancePerUnitColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    With ActiveSheet.AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    currentCatchWeightNBeginningRow = Range(cWColumnLetter & "1:" & cWColumnLetter & lastRow).Find("N", lookat:=xlWhole, searchdirection:=xlNext).Row
'
'    currentCatchWeightNAndNotZeroBeginningRow = Evaluate("=MATCH(FALSE," & customerPickupAllowancePerUnitColumnLetter & currentCatchWeightNBeginningRow & ":" & customerPickupAllowancePerUnitColumnLetter & lastRow & "=0,0)")
'
'    currentCatchWeightNAndNotZeroBeginningRow = currentCatchWeightNAndNotZeroBeginningRow + currentCatchWeightNBeginningRow - 1
'
''    Range(customerPickupAllowancePerUnitColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & ":" & customerPickupAllowancePerUnitColumnLetter & lastRow).Formula = "=" & vendorProductPricePerUnitColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & "+" & prePayAndAddFreightPerUnitColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & "+" & freightRatePerUnitColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & "-" & divisionDeliveredPricePerUnitColumnLetter & currentCatchWeightNAndNotZeroBeginningRow
'
''    Range(customerPickupAllowancePerUnitColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & ":" & customerPickupAllowancePerUnitColumnLetter & lastRow).Value = Range(customerPickupAllowancePerUnitColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & ":" & customerPickupAllowancePerUnitColumnLetter & lastRow).Value
'
'    Range(recalculationColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & ":" & recalculationColumnLetter & lastRow).Formula = "=(" & vendorProductPricePerUnitColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & "-" & contBasisColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & ")*" & casesColumnLetter & currentCatchWeightNAndNotZeroBeginningRow
'
'    Range(resolutionColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & ":" & resolutionColumnLetter & lastRow).Formula = "=" & recalculationColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & "-" & rebAmtColumnLetter & currentCatchWeightNAndNotZeroBeginningRow
'
'    Range(reasonColumnLetter & currentCatchWeightNAndNotZeroBeginningRow & ":" & reasonColumnLetter & lastRow).Value = "Freight billed above customer pickup allowance is to be repaid per FreightAllowanceRule"
'
'    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
'
'    Set getPosition = Cells(1, Columns.count).End(xlToLeft)
'
'    lastColumnColumnLetter = Evaluate("MID(""" & getPosition.Address & """,FIND(""$"",""" & getPosition.Address & """)+1,FIND(""$"",""" & getPosition.Address & """,2)-2)")
'
'    lastRow = Cells(Rows.count, 1).End(xlUp).Row
'
'    Range("A1:" & lastColumnColumnLetter & lastRow).AutoFilter
'
'    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
'        (runDateColumnLetter & "2:" & runDateColumnLetter & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    With ActiveSheet.AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    On Error Resume Next
'
'    Set freightAllowanceRuleSheet = ActiveWorkbook.Sheets("FreightAllowanceRule")
'
'    If Err.Number > 0 Then
'
'        Set mainSheet = ActiveSheet
'
'        Err.Clear
'
'        Sheets.Add After:=ActiveSheet
'
'        Set freightAllowanceRuleSheet = ActiveSheet
'
'        freightAllowanceRuleSheet.Name = "FreightAllowanceRule"
'        ActiveSheet.OLEObjects.Add(Filename:= _
'            "http://sharepoint.usfood.com/sites/Finance/CASS/HPSM/HPSM%20Team/Shared%20Documents/Freight%20allowance%20rule.msg" _
'            , link:=False, DisplayAsIcon:=False).Select
'
'        mainSheet.Select
'
'    End If
'
'    MsgBox ("This macro does not consider if the delivered price is less than the contract price.")

End Sub

Sub subIEPRISMLaunch()

    Dim iE As Object
    Dim passWord As String
    Dim userName As String
    
    Set iE = CreateObject("InternetExplorer.Application")
    
    iE.navigate "http://appcenter/Citrix/XenApp/auth/login.aspx"
    
    iE.Visible = True
    
    userName = "D3Q1700"
    passWord = "Today123456789"
    
    'This keeps the rest of the macro from executing until Internet Explorer has loaded
    While iE.busy
    DoEvents
    Wend
    
    
    'Wait a second before the next step
    Application.Wait Now + TimeValue("00:00:03")
    
    
    iE.document.all("user").Value = userName
    iE.document.all("password").Value = passWord
    iE.document.all("domain").Value = "USF"
    
    iE.document.all("loginButtonWrapper").Click
    
    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeValue("00:00:01")
    
    iE.document.all("folderLink_2").Click
    
    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeValue("00:00:01")
    
    iE.document.all("appLink_0").Click
    
    Application.Wait Now + TimeValue("00:00:10")
    
    iE.Quit

End Sub

Sub subIEPullInfoFromCasisBasedOnCustNbr()

Dim iE As Object
Dim lastRow As Long

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Set iE = CreateObject("InternetExplorer.Application")

iE.Visible = True

iE.navigate "http://cas.usfood.com/CAS/customerInformation.jsp"

While iE.busy
DoEvents
Wend

Application.Wait Now + TimeSerial(0, 0, 1.5)

Do Until iE.locationurl <> "http://cas.usfood.com/CAS/login.jsp"

    MsgBox ("Log in to CASIS then click Ok")

    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeSerial(0, 0, 1)

Loop

Do Until ActiveCell.Row > lastRow

    iE.document.all("srchCntlLocn").Value = 6026
    
    iE.document.all("srchCustNbr").Value = ActiveCell.Offset(, -1)
    'http://cas.usfood.com/CAS/customerInformation.jsp
    
    iE.navigate "JavaScript:submitAction('search')"
    
    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeSerial(0, 0, 1.5)
    
    Do Until iE.locationurl <> "http://cas.usfood.com/CAS/login.jsp"
    
        MsgBox ("Log in to CASIS then click Ok")
    
        While iE.busy
        DoEvents
        Wend
        
        Application.Wait Now + TimeSerial(0, 0, 1)
    
    Loop
    
    If iE.document.all("custNbr").Value = ActiveCell.Offset(, -1).Text Then
    
        ActiveCell = iE.document.all("division").Value
    
    Else
    
        ActiveCell = "Customer Number Not Found"
    
    End If
    
    ActiveCell.Offset(1, 0).Select

Loop

End Sub

'http://cas.usfood.com/CAS/trackingProgramSearch.jsp?action=doExport&vendorNbr=4016&pgmBasis=SL

Function RangetoHTML(rng As Range)
' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, and Outlook 2010.
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
 
    ' Copy the range and create a workbook to receive the data.
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
 
    Rows(1).Font.Bold = True
 
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
 
    Selection.WrapText = True
 
    Selection.Columns.AutoFit
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
 
    ' Publish the sheet to an .htm file.
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
 
    ' Read all data from the .htm file into the RangetoHTML subroutine.
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
 
    ' Close TempWB.
    TempWB.Close savechanges:=False
 
    ' Delete the htm file.
    Kill TempFile
 
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


