Attribute VB_Name = "modPersonal"
Option Explicit

Sub subPersonalCraiglistPoster()

Dim iE As Object
Dim loginButtonClick As Variant

Set iE = CreateObject("InternetExplorer.Application")

iE.Visible = True

iE.navigate "https://accounts.craigslist.org/login"

While iE.busy
DoEvents
Wend

Application.Wait Now + TimeSerial(0, 0, 1.5)

If InStr(iE.locationurl, "https://accounts.craigslist.org/login") > 0 Then

    Do Until iE.locationurl <> "https://accounts.craigslist.org/login"
    
        MsgBox ("Log in then click okay")
        
        While iE.busy
            
            DoEvents
        
        Wend

        Application.Wait Now + TimeSerial(0, 0, 1.5)
    
    Loop

Else

    iE.document.all("inputEmailHandle").Value = "XXXXXXXXX"
    
    iE.document.all("inputPassword").Value = "XXXXXXXXX"
    
    'Log in button
    
    For Each buttonTag In iE.document.body.getelementsbytagname("button")
    
        If buttonTag.innertext = "Log in" Then
    
            buttonTag.Click
    
            Exit For
    
        End If
    
    Next
    
    While iE.busy
    DoEvents
    Wend
    
    Application.Wait Now + TimeSerial(0, 0, 1.5)

End If

iE.navigate "https://post.craigslist.org/c/phx"

While iE.busy
DoEvents
Wend

Application.Wait Now + TimeSerial(0, 0, 1.5)

'For sale by owner radio

For Each inputTag In iE.document.body.getelementsbytagname("input")

    If inputTag.Value = "fso" Then

        inputTag.Click

        Exit For

    End If

Next

'Continue button

For Each buttonTag In iE.document.body.getelementsbytagname("button")

    If buttonTag.innertext = "continue" Then

        buttonTag.Click

        Exit For

    End If

Next

While iE.busy
DoEvents
Wend

Application.Wait Now + TimeSerial(0, 0, 1.5)

'Furniture by owner radio button

For Each inputTag In iE.document.body.getelementsbytagname("input")

    If inputTag.Value = "141" Then

        inputTag.Click

        Exit For

    End If

Next

While iE.busy
DoEvents
Wend

Application.Wait Now + TimeSerial(0, 0, 1.5)

'Continue Button

For Each inputTag In iE.document.body.getelementsbytagname("input")

    If inputTag.Value = "2" Then

        inputTag.Click

        Exit For

    End If

Next

'continueButton

For Each buttonTag In iE.document.body.getelementsbytagname("button")

    If buttonTag.innertext = "continue" Then

        buttonTag.Click

        Exit For

    End If

Next

While iE.busy
DoEvents
Wend

Application.Wait Now + TimeSerial(0, 0, 1.5)

iE.document.getelementbyid("PostingTitle").Value = "Brand New! BLACK Modern Bar Stools with adjustable seat height."

'Price

For Each inputTag In iE.document.body.getelementsbytagname("input")

    If inputTag.ID = "Ask" Then

        inputTag.Value = "65"

        Exit For

    End If

Next

'Specific Location

For Each inputTag In iE.document.body.getelementsbytagname("input")

    If inputTag.ID = "GeographicArea" Then

        inputTag.Value = "Scottsdale Rd and McKellips"

        Exit For

    End If

Next

'Postal code

For Each inputTag In iE.document.body.getelementsbytagname("input")

    If inputTag.ID = "postal_code" Then

        inputTag.Value = "85281"

        Exit For

    End If

Next

'Posting body

iE.document.all("PostingBody").Value = "Brand new modern barstools. Perfect for bars and islands. All black bonded leather.<br></br>" & _
"Color: Black<br></br>" & _
"Adjustable seat height: 24"" to 32""(From the ground to seat cushion)<br></br>" & _
"Seat Width: 21 3/4""<br></br>" & _
"Seat Depth : 16""<br></br>" & _
"10 3/4"" high back support.<br></br>" & _
"I have 6 Barstools<br></br>" & _
"$70 per stool<br></br>" & _
"$130 for 2<br></br>" & _
"Lower price when you buy 3 or more<br></br>" & _
"More pictures at <a href=""http://www.thebarstoolguys.com"">TheBarStoolGuys . com</a><br></br><br></br><br></br>" & _
"Bar stool, high chair, table, apartment, condo, studio, pub, office, hydraulic, contemporary, air lift, home, house, design, barstool, designer, barstools, kitchen, modern, chic, barstools,"

'Condition

iE.document.all("condition").Value = "10"

'Text preferred

iE.document.all("contact_text_ok").Click

'Phone number

iE.document.getelementbyid("contact_phone").Value = "4804207690"

'continueButton

For Each buttonTag In iE.document.body.getelementsbytagname("button")

    If buttonTag.innertext = "continue" Then

        buttonTag.Click

        Exit For

    End If

Next

While iE.busy
DoEvents
Wend

Application.Wait Now + TimeSerial(0, 0, 1.5)

'In maps screen

iE.document.all("xstreet0").Value = "Scottsdale Rd"

iE.document.all("xstreet1").Value = "McKellips"

'continueButton

For Each buttonTag In iE.document.body.getelementsbytagname("button")

    If buttonTag.innertext = "continue" Then

        buttonTag.Click

        Exit For

    End If

Next

While iE.busy
DoEvents
Wend

Application.Wait Now + TimeSerial(0, 0, 1.5)

End Sub
