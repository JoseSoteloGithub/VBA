Attribute VB_Name = "modProgramDevelopment"

'Takes apart the word after "Dim"
'Mid(Cells(ActiveCell.Row, 1),firstSpace,variableCharacterLength)

Sub subDimensionClearingHelp()
    
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
        
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:A" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A1").Select

    Do Until Cells(ActiveCell.Row, 1) = ""

        rightSpace = InStrRev(Cells(ActiveCell.Row, 1), " ")
    
        firstSpace = Evaluate("FIND("" "",""" & Cells(ActiveCell.Row, 1) & """) + 1")
        
        variableCharacterLength = Evaluate("FIND("" "",""" & Cells(ActiveCell.Row, 1) & """,FIND("" "",""" & Cells(ActiveCell.Row, 1) & """) + 1) - FIND("" "",""" & Cells(ActiveCell.Row, 1) & """) - 1")
    
        Cells(ActiveCell.Row, 2) = Evaluate("RIGHT(""" & Cells(ActiveCell.Row, 1) & """,LEN(""" & Cells(ActiveCell.Row, 1) & """)-" & rightSpace & ")")
    
        Select Case Cells(ActiveCell.Row, 2)
        
            Case "Boolean"
                Cells(ActiveCell.Row, 2) = Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = empty"
        
            Case "Date"
                Cells(ActiveCell.Row, 2) = Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = empty"
        
            Case "Double"
                Cells(ActiveCell.Row, 2) = Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = empty"
        
            Case "Integer"
                Cells(ActiveCell.Row, 2) = Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = empty"
        
            Case "Object"
                Cells(ActiveCell.Row, 2) = "Set " & Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = nothing"
        
            Case "Range"
                Cells(ActiveCell.Row, 2) = "Set " & Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = nothing"
                    
            Case "String"
                Cells(ActiveCell.Row, 2) = Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = """""
        
            Case "Long"
                Cells(ActiveCell.Row, 2) = Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = empty"
            
            Case "Variant"
                Cells(ActiveCell.Row, 2) = "Set " & Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = nothing"
            
            Case "Workbook"
                Cells(ActiveCell.Row, 2) = "Set " & Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = nothing"
                
            Case "Worksheet"
                Cells(ActiveCell.Row, 2) = "Set " & Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = nothing"
                
            Case "PivotField"
                Cells(ActiveCell.Row, 2) = "Set " & Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = nothing"
            
            Case "PivotTable"
                Cells(ActiveCell.Row, 2) = "Set " & Mid(Cells(ActiveCell.Row, 1), firstSpace, variableCharacterLength) & " = nothing"
                        
        End Select
    
        ActiveCell.Offset(1).Select
    
    Loop
    
End Sub

Sub iEFindHiddenIEWindows()

Dim iE As Object


For Each iE In CreateObject("Shell.Application").Windows

'Debug.Print iE

    If iE = "Windows Internet Explorer" Then
        
            iE.Visible = True
            
'            iE.Quit
        
    End If

Next iE

End Sub

Sub iECloseAlliE()

    Dim iE As Object
    
    For Each iE In CreateObject("Shell.Application").Windows
    
        If iE = "Windows Internet Explorer" Then
        
            iE.Visible = False
        
            iE.Quit
        
        End If
        
    Next iE
    
    Set iE = Nothing
    
End Sub



Sub subCommonlyUsedVariables()

    Set lastColumnPosition = Cells(1, Columns.Count).End(xlToLeft)

    lastColumnColumnLetter = Evaluate("MID(""" & lastColumnPosition.Address & """,FIND(""$"",""" & lastColumnPosition.Address & """)+1,FIND(""$"",""" & lastColumnPosition.Address & """,2)-2)")
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Set contractPosition = Range("A1:IV1").Find("Cntrct #", lookat:=xlWhole)
    
    contractColumnLetter = Evaluate("MID(""" & contractPosition.Address & """,FIND(""$"",""" & contractPosition.Address & """)+1,FIND(""$"",""" & contractPosition.Address & """,2)-2)")



    Set masterProgramNumberPosition = Range("A1:IV1").Find("Master Program Number", lookat:=xlWhole)
    
    If masterProgramNumberPosition Is Nothing Then
    
        lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    
        Cells(1, lastColumn).Offset(, 1) = "Master Program Number"
    
        Set masterProgramNumberPosition = Range("A1:IV1").Find("Master Program Number")
    
    End If
    
    masterProgramNumberColumnLetter = Evaluate("MID(""" & masterProgramNumberPosition.Address & """,FIND(""$"",""" & masterProgramNumberPosition.Address & """)+1,FIND(""$"",""" & masterProgramNumberPosition.Address & """,2)-2)")

End Sub

Sub subDevelopmentFindByTag()

    For Each discovererLink In iE.document.body.getelementsbytagname("a")
        
        Debug.Print discovererLink.innertext
        
        Debug.Print discovererLink.Title
        
        If discovererLink.Title = "Discoverer" Then
        
            Exit For
        
        End If
               
    Next discovererLink
        
    For Each tdTag In iE.document.body.getelementsbytagname("td")
        
        'Debug.Print tdTag.Type

        'ActiveCell = "td" & i

        'Debug.Print "td" & i

        'ActiveCell.Offset(, 1) = tdTag.ID

        'Debug.Print tdTag.ID

        'Debug.Print tdTag.Value
        
        'Debug.Print tdTag.class
        
        'ActiveCell.Offset(, 2) = tdTag.innertext
        
        'Debug.Print tdTag.innertext
        
        If i >= tdTagsToReachInvoice Then
                
            MsgBox ("Found first invoice td tag " & tdTag.innertext)
            
            tdTagsToReachInvoice = tdTagsToReachInvoice + 26
            
        End If
        
        If i >= tdTagsToReachReversalAmount Then
                
            MsgBox ("Found first invoice td tag " & tdTag.innertext)
            
            tdTagsToReachReversalAmount = tdTagsToReachReversalAmount + 26
            
        End If
        
        i = i + 1
        
        'ActiveCell.Offset(1).Select
               
    Next tdTag
        
End Sub

Sub subDevelopmentEnviron()
    Dim i As Integer
    Dim stEnviron As String
    For i = 1 To 50
        ' get the environment variable
        stEnviron = Environ(i)
        ' see if there is a variable set
        If Len(stEnviron) > 0 Then
            Debug.Print i, Environ(i)
        Else
            Exit For
        End If
    Next

End Sub

Sub subCreateColumnHeadingVariable()

columnHeader = ActiveCell.Text

Set beginingSheet = ActiveSheet
    
Sheets.Add
    
Set newSheet = ActiveSheet
    
newSheet.Cells(Rows.Count, 1).End(xlUp) = columnHeader

newSheet.Cells(Rows.Count, 1).End(xlUp).Offset(2).Formula = "=SUBSTITUTE(A1,"" "","""")"

newSheet.Cells(Rows.Count, 1).End(xlUp).Offset(2).Formula = "=LOWER(LEFT(A3,1))&RIGHT(A3,LEN(A3)-1)"

newSheet.Cells(Rows.Count, 1).End(xlUp).Offset(2).Formula = "=""Set getPosition = Range(""""A1:IV1"""").Find(""""""&A1&"""""", lookat:=xlWhole)"""

newSheet.Cells(Rows.Count, 1).End(xlUp).Offset(2).Formula = "=A5&""ColumnLetter = Evaluate(""""MID("""""""""""" & ""&""getPosition.Address & """""""""""",FIND(""""""""$"""""""","""""""""""" & ""&""getPosition.Address & """""""""""")+1,FIND(""""""""$"""""""","""""""""""" & ""&""getPosition.Address & """""""""""",2)-2)"""")"""

End Sub

Sub subTestingWindowsAutomation()

'    i = 1
'
'    For Each thisshell In CreateObject("Shell.Application").Windows
'
'        Debug.Print i & "=" & thisshell & " thisShell location url is " & thisshell.locationurl
'
'        If thisshell = "Internet Explorer" Then
'
'            If thisshell.locationurl = "http://psfsapp9:8000/psp/FSPRD90/EMPLOYEE/ERP/h/?tab=DEFAULT" Then
'
'                Exit For
'
'            End If
'
'        End If
'
'        i = i + 1
'
'    Next thisshell
    
    Set iE = CreateObject("InternetExplorer.Application")
    
    iE.Visible = True
    
    iE.navigate "http://psfsapp9:8000/psp/FSPRD90/EMPLOYEE/ERP/h/?tab=DEFAULT"
    
    
    
    Set objWSS = CreateObject("WScript.shell")
    
    'thisshell.Focus
    
    i = 0
    
    iE.document.Focus
    
    Do Until i > 1
    
        objWSS.SendKeys "{TAB}"
        
        DoEvents
        
        Application.Wait Now + TimeSerial(0, 0, 1)
        
        i = i + 1
    
    Loop
    
    objWSS.SendKeys "{ENTER}"
    
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    iE.document.Focus
    
    objWSS.SendKeys "{TAB}"
    
    DoEvents
    
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    objWSS.SendKeys "{TAB}"
    
    DoEvents
    
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    objWSS.SendKeys "{ENTER}"
    
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    'http://psfsapp9:8000/psp/FSPRD90/EMPLOYEE/ERP/h/?tab=DEFAULT
    '/copMaintenance.jsp?action=doSaveAs
    i = 0
    
    Do Until i > 100
    
        thisshell.document.all("doInsert").Click
    
        i = i + 1
    
    Loop
    
End Sub



Sub subDeleteCookiesAndCloseIE()
    
    Dim iE As Object
    
    For Each iE In CreateObject("Shell.Application").Windows

        If iE = "Internet Explorer" Then

            iE.Quit

        End If

    Next iE
    
    Sleep (500)
    
    i = 0
    
    Do Until i = 2
    
        Shell "RunDll32.exe InetCpl.Cpl, ClearMyTracksByProcess 2"
        
        Sleep (500)
    
        i = i + 1
    
    Loop
    
    
''RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
'
'echo Clear Temporary Internet Files:
'RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
'
'echo Clear Cookies:
'RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
'
'echo Clear History:
'RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1
'
'echo Clear Form Data:
'RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16
'
'echo Clear Saved Passwords:
'RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32
'
'echo Delete All:
'RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255
'
'echo Delete All w/Clear Add-ons Settings:
'RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351
End Sub

Sub pasteClipboard()

    Set iE = CreateObject("InternetExplorer.Application")
    
    iE.Visible = True
    
    'iE.navigate "http://psfsapp9:8000/psp/FSPRD90/EMPLOYEE/ERP/h/?tab=DEFAULT"

    iE.navigate "http://psfsapp9:8000/psp/FSPRD90/EMPLOYEE/ERP/c/ENTER_VOUCHER_INFORMATION.VCHR_PO_STD.GBL?PORTALPARAM_PTCNAV=PO_VOUCHER&EOPP.SCNode=ERP&EOPP.SCPortal=EMPLOYEE&EOPP.SCName=EPCO_ACCOUNTS_PAYABLE&EOPP.SCLabel=Vouchers&EOPP.SCFName=EPAP_VOUCHERS&EOPP.SCSecondary=true&EOPP.SCPTfname=EPAP_VOUCHERS&FolderPath=PORTAL_ROOT_OBJECT.EPCO_ACCOUNTS_PAYABLE.EPAP_VOUCHERS.EPAP_VCHR_ENTRY.PO_VOUCHER&IsFolder=false"

    i = 1

    For Each discovererLink In iE.document.getelementbyid("1P")
        
        ActiveCell = i & "Name=" & discovererLink.Name & " Value=" & discovererLink.Value
        
        'Debug.Print discovererLink.Title
        
        If discovererLink.Title = "Discoverer" Then
        
            Exit For
        
        End If
               
        ActiveCell.Offset(1).Select
               
        i = i + 1
        
    Next discovererLink
    
    iE.navigate "http://psfsapp9:8000/psp/FSPRD90/EMPLOYEE/ERP/s/WEBLIB_PTPP_SC.HOMEPAGE.FieldFormula.IScript_AppHP?pt_fname=EPCO_ACCOUNTS_PAYABLE&amp;FolderPath=PORTAL_ROOT_OBJECT.EPCO_ACCOUNTS_PAYABLE&amp;IsFolder=true"
    iE.navigate "http://psfsapp9:8000/psc/FSPRD90/EMPLOYEE/ERP/s/WEBLIB_PTPP_SC.HOMEPAGE.FieldFormula.IScript_AppHP?scname=EPCO_ACCOUNTS_PAYABLE&secondary=true&fname=EPAP_VOUCHERS&pt_fname=EPAP_VOUCHERS&PortalCacheContent=true&PSCache-Control=role%2cmax-age%3d60&FolderPath=PORTAL_ROOT_OBJECT.EPCO_ACCOUNTS_PAYABLE.EPAP_VOUCHERS&IsFolder=true"
    
    iE.navigate "http://psfsapp9:8000/psp/FSPRD90/EMPLOYEE/ERP/c/ENTER_VOUCHER_INFORMATION.VCHR_PO_STD.GBL?PORTALPARAM_PTCNAV=PO_VOUCHER&EOPP.SCNode=ERP&EOPP.SCPortal=EMPLOYEE&EOPP.SCName=EPCO_ACCOUNTS_PAYABLE&EOPP.SCLabel=Vouchers&EOPP.SCFName=EPAP_VOUCHERS&EOPP.SCSecondary=true&EOPP.SCPTfname=EPAP_VOUCHERS&FolderPath=PORTAL_ROOT_OBJECT.EPCO_ACCOUNTS_PAYABLE.EPAP_VOUCHERS.EPAP_VCHR_ENTRY.PO_VOUCHER&IsFolder=false"
End Sub
   
   
   'http://psfsapp9:8000/psp/FSPRD90/EMPLOYEE/ERP/h/?tab=DEFAULT

    

