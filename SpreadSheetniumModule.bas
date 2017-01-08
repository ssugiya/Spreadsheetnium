Attribute VB_Name = "SpreadSheetniumModule"
Option Explicit

#Const DBG = 0

Private Declare Function MessageBoxTimeoutA Lib "user32" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long, ByVal dwMilliseconds As Long) As Long
Public driver As New WebDriver
Public Verify As New Selenium.Verify
Public Const findElementTimeOut As Long = 0 '3000
Public Const passedColorCode As Long = 11854022 'RGB(198, 224, 180)
Public Const failedColorCode As Long = 11389944 'RGB(248, 203, 173)
Public Rtn


Private Sub Auto_Open()

    Dim testTargetSheet  As Worksheet
    
    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    Set testTargetSheet = Worksheets("BATCH")
    testTargetSheet.Activate
    
    If Range("AutoRun").Text = "Yes" Then
        If MessageBoxTimeoutA(0&, "Batch script will be started automatically in 10 seconds." & vbCrLf & "Please CANCEL if you stop batch script.", "Answer within 10 seconds!", vbMsgBoxSetForeground + vbQuestion + vbOKCancel + vbDefaultButton2, 0, 10000) = vbCancel Then
            Exit Sub
        Else
            Call batchRunScript
            ActiveWorkbook.Save
            Application.Quit
        End If
    End If
    
    Exit Sub
        
Err: '----------------------------
    Rtn = errHandler("Auto_Open", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If


End Sub
Private Function errHandler(procName As String, ErrNumber As Long)

    Dim errMsg As String
    
    Select Case ErrNumber
        Case 26 'unexpected alert open(Resume Next)
            errHandler = 0
        Case -2146233078 'ソースが見つかりませんでしたが､いくつかまたはすべてのログを検索できませんでした｡アクセス不可能なログ:   Security
            errHandler = 0
        Case Else
            errHandler = -1
            errMsg = Now() & vbCrLf & _
                    "Procedure: " & procName & vbCrLf & _
                    "Err number: " & Err.Number & vbCrLf & _
                    Err.Description
        
        #If DBG <> 0 Then
            Debug.Print errMsg & vbCrLf & vbCrLf
        #End If
        Cells(9, 12).Value = Cells(9, 12).Value & errMsg & vbCrLf & vbCrLf

    End Select
   

End Function
Public Sub runTestScriptConfirm()

    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    If MsgBox("Do you want to run test script?", vbOKCancel + vbExclamation + vbDefaultButton2, "Run test script") = vbCancel Then
        Exit Sub
    Else
        Call runScript
    End If

    If Range("ReportResults").Text = "Yes" Then
        'report for each script
        Call reportResults
    End If

    Exit Sub

Err: '----------------------------
    Rtn = errHandler("runTestScriptConfirm", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub

Private Sub runScript()

    Dim command As String, findMethod As String, actionTarget As String, actionValue As String
    Dim targetBrowser As String, baseURL As String, windowSizeW, windowSizeH As Integer, screenshotPath As String, screenshotFile As String
    Dim verificationCommand As String, verificationMethod As String, verificationTarget As String
    Dim Rtn
    Dim LS As ListObject
    Dim R As ListRow
    Dim by As New by
    
    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    Application.StatusBar = "Initializing."
    Cells(9, 12).Value = ""
    Call clearTestResults
    
    '==========================================
    'Initial settings
    '==========================================
    targetBrowser = Range("targetBrowser").Text
    baseURL = Range("baseURL").Text
    windowSizeW = Range("windowSizeW").Text
    windowSizeH = Range("windowSizeH").Text
    screenshotPath = Range("ScreenshotPath").Text
    
    '==========================================
    'Start test
    '==========================================
    driver.Start targetBrowser, baseURL
    driver.Window.SetSize windowSizeW, windowSizeH
    
    If LCase(Range("DeleteCookie").Text) = "yes" Then
        driver.Manage.DeleteAllCookies
    End If
    
    'Loop test cases
    Set LS = ActiveSheet.ListObjects(1)
    For Each R In LS.ListRows
        
        Application.StatusBar = "Test script is running...  " & R.Index & "/" & LS.ListRows.Count
        
        If LCase(R.Range(LS.ListColumns("runTarget").Index)) <> "yes" Then
            Call skipTest(R, LS, "Skipped (run-target does not Yes)")
            GoTo nextRowNum
        End If
        
        'get palameters from excel sheet
        command = R.Range(LS.ListColumns("command").Index)
        findMethod = R.Range(LS.ListColumns("FindMethod").Index)
        actionTarget = R.Range(LS.ListColumns("ActionTarget").Index)
        actionValue = R.Range(LS.ListColumns("ActionValue").Index)
        
    '==========================================
    'Run selenium action by command
    '==========================================
        Select Case command
            Case "Get"
                driver.Get actionTarget
            Case "Click"
                Rtn = commandClick(findMethod, actionTarget, R, LS)
                If Rtn = -1 Then: GoTo nextRowNum
            Case "SendKeys"
                Rtn = commandSendKeys(findMethod, actionTarget, actionValue, R, LS)
                If Rtn = -1 Then: GoTo nextRowNum
            Case "TakeScreenshot"
                driver.TakeScreenshot.SaveAs actionTarget & "\" & actionValue
            Case "Wait"
                driver.Wait actionValue
            Case "GoBack"
                driver.GoBack
                driver.Wait findElementTimeOut
            Case "Select"
                Rtn = commandSelect(findMethod, actionTarget, actionValue, R, LS)
                If Rtn = -1 Then: GoTo nextRowNum
            Case "Radio"
                Rtn = commandRadio(findMethod, actionTarget, actionValue, R, LS)
                If Rtn = -1 Then: GoTo nextRowNum
            Case "MouseMoveTo"
                Rtn = commandMouseMoveTo(findMethod, actionTarget, R, LS)
                If Rtn = -1 Then: GoTo nextRowNum
            Case "Submit"
                Rtn = commandSubmit(findMethod, actionTarget, R, LS)
                If Rtn = -1 Then: GoTo nextRowNum
            Case "Alert"
                Rtn = commandAlert(findMethod, actionTarget, actionValue, R, LS)
                If Rtn = -1 Then: GoTo nextRowNum
            Case "SwitchToWindow"
                driver.SwitchToWindowByTitle(actionTarget).Activate
                driver.Wait findElementTimeOut
            Case Else
        
        End Select
        
        
    '==========================================
    'Start verification
    '==========================================
        verificationCommand = R.Range(LS.ListColumns("VerificationCommand").Index)
        verificationMethod = R.Range(LS.ListColumns("VerificationMethod").Index)
        verificationTarget = R.Range(LS.ListColumns("VerificationTarget").Index)
        Rtn = ""
        
        If verificationCommand = "" Then
            Call skipTest(R, LS, "Skipped (No verification command)")
            GoTo nextRowNum
        End If
        
        'get actual results
        Select Case verificationCommand
            Case "Title"
                R.Range(LS.ListColumns("ActualResult").Index) = driver.Title
            Case "Url"
                R.Range(LS.ListColumns("ActualResult").Index) = driver.Url
            Case "Contains", "Equals", "Matches"
                Select Case verificationMethod
                    Case "Id"
                        If driver.IsElementPresent(by.ID(verificationTarget)) Then
                            R.Range(LS.ListColumns("ActualResult").Index) = driver.FindElementById(verificationTarget).Text
                        End If
                    Case "Css"
                        If driver.IsElementPresent(by.Css(verificationTarget)) Then
                            R.Range(LS.ListColumns("ActualResult").Index) = driver.FindElementByCss(verificationTarget).Text
                        Else
                            R.Range(LS.ListColumns("ErrorMessage").Index) = "Verification skipped(No element)"
                        End If
                    Case "Name"
                        R.Range(LS.ListColumns("ActualResult").Index) = driver.FindElementByName(verificationTarget).Text
                    Case "XPath"
                        R.Range(LS.ListColumns("ActualResult").Index) = driver.FindElementByXPath(verificationTarget).Text
                    Case Else
                        Call skipTest(R, LS, "Skipped (No verification method)")
                        GoTo nextRowNum
                    End Select
            Case Else
                Call skipTest(R, LS, "Skipped (No verification command)")
                GoTo nextRowNum
        End Select
        
        'verify results
        Select Case verificationCommand
            Case "Contains"
                Rtn = Verify.Contains(R.Range(LS.ListColumns("ExpectedResult").Index).Text, R.Range(LS.ListColumns("ActualResult").Index).Text)
            Case "Equals"
                Rtn = Verify.Equals(R.Range(LS.ListColumns("ExpectedResult").Index).Text, R.Range(LS.ListColumns("ActualResult").Index).Text)
            Case "Matches" 'regular expression
                Rtn = Verify.Matches(R.Range(LS.ListColumns("ExpectedResult").Index).Text, R.Range(LS.ListColumns("ActualResult").Index).Text)
        End Select
    
        'decide test results
        If Rtn = "OK" Then
            R.Range(LS.ListColumns("Result").Index) = "Passed"
            R.Range(LS.ListColumns("Result").Index).Interior.Color = passedColorCode
        ElseIf Rtn Like "NOK*" Then
            R.Range(LS.ListColumns("Result").Index) = "Failed"
            R.Range(LS.ListColumns("Result").Index).Interior.Color = failedColorCode
        ElseIf R.Range(LS.ListColumns("ActualResult").Index).Text = R.Range(LS.ListColumns("ExpectedResult").Index).Text Then
            R.Range(LS.ListColumns("Result").Index) = "Passed"
            R.Range(LS.ListColumns("Result").Index).Interior.Color = passedColorCode
        Else
            R.Range(LS.ListColumns("Result").Index) = "Failed"
            R.Range(LS.ListColumns("Result").Index).Interior.Color = failedColorCode
        End If
        DoEvents
        DoEvents
    
nextRowNum:
        
        'record datetime of test run
        R.Range(LS.ListColumns("LastUpdate").Index) = Now()
        
        'Save screenshot if you need
        If LCase(R.Range(LS.ListColumns("runTarget").Index)) = "yes" Then
            If screenshotPath <> "" Then
                screenshotFile = R.Range(LS.ListColumns("scriptID").Index) & "_" & driver.Title & "_" & R.Range(LS.ListColumns("Description").Index) & "_" & R.Range(LS.ListColumns("Result").Index) & ".png"
                screenshotFile = Replace(screenshotFile, "\", "")
                screenshotFile = Replace(screenshotFile, "/", "")
                screenshotFile = Replace(screenshotFile, ":", "")
                screenshotFile = Replace(screenshotFile, "*", "")
                screenshotFile = Replace(screenshotFile, "?", "")
                screenshotFile = Replace(screenshotFile, """", "")
                screenshotFile = Replace(screenshotFile, "<", "")
                screenshotFile = Replace(screenshotFile, ">", "")
                screenshotFile = Replace(screenshotFile, "|", "")
                driver.TakeScreenshot.SaveAs screenshotPath & "\" & screenshotFile
            End If
        End If
    
    Next R
    
        Call exitProgram
        Application.StatusBar = "Test script finished."
    
    Exit Sub

Err: '----------------------------

Rtn = errHandler("runScript", Err.Number)
If Rtn = 0 Then
    Resume Next
Else
    Call exitProgram
End If
    
End Sub

Private Sub skipTest(R As ListRow, LS As ListObject, msg As String)

    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    R.Range(LS.ListColumns("ActualResult").Index) = ""
    R.Range(LS.ListColumns("Result").Index) = msg
    R.Range(LS.ListColumns("LastUpdate").Index) = ""
    R.Range(LS.ListColumns("Result").Index).ClearFormats
    
    Exit Sub

Err: '----------------------------
    Rtn = errHandler("skipTest", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub

Private Sub exitProgram()

    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    driver.Quit
    ActiveWorkbook.Save
    
    Exit Sub

Err: '----------------------------
    Rtn = errHandler("exitProgram", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub

Private Function commandClick(findMethod, actionTarget As String, R As ListRow, LS As ListObject)

    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    Select Case findMethod
        Case "Id"
            #If DBG = 1 Then
                Dim by As New by
                Debug.Print driver.IsElementPresent(by.ID(actionTarget))
            #End If
            driver.FindElementById(actionTarget, findElementTimeOut).Click
        Case "LinkText"
            driver.FindElementByLinkText(actionTarget, findElementTimeOut).Click
        Case "Name"
            driver.FindElementByName(actionTarget, findElementTimeOut).Click
        Case "XPath"
            driver.FindElementByXPath(actionTarget, findElementTimeOut).Click
        Case "Css"
            driver.FindElementByCss(actionTarget, findElementTimeOut).Click
        Case Else
            Call skipTest(R, LS, "Skipped (No find method)")
            commandClick = -1
      End Select
      driver.Wait findElementTimeOut
    
    Exit Function

Err: '----------------------------
    Rtn = errHandler("commandClick", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Function

Private Function commandSubmit(findMethod, actionTarget As String, R As ListRow, LS As ListObject)

    #If DBG = 0 Then
        On Error GoTo Err
    #End If
        
    Select Case findMethod
        Case "Id"
            driver.FindElementById(actionTarget).Submit
        Case "Link"
            driver.FindElementByLinkText(actionTarget).Submit
        Case "Name"
            driver.FindElementByName(actionTarget).Submit
        Case "XPath"
            driver.FindElementByXPath(actionTarget).Submit
        Case "Css"
            driver.FindElementByCss(actionTarget).Submit
        Case Else
            Call skipTest(R, LS, "Skipped (No find method)")
            commandSubmit = -1
    End Select
    driver.Wait findElementTimeOut
    
    Exit Function

Err: '----------------------------
    Rtn = errHandler("commandSubmit", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Function

Private Function commandSendKeys(findMethod As String, actionTarget As String, actionValue As String, R As ListRow, LS As ListObject)

    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    Select Case findMethod
      Case "Id"
          With driver.FindElementById(actionTarget)
              .Clear
              .SendKeys actionValue
          End With
      Case "Name"
          With driver.FindElementByName(actionTarget)
              .Clear
              .SendKeys actionValue
          End With
      Case "XPath"
          With driver.FindElementByXPath(actionTarget)
              .Clear
              .SendKeys actionValue
          End With
      Case "Css"
          With driver.FindElementByCss(actionTarget)
              .Clear
              .SendKeys actionValue
          End With
      Case Else
          Call skipTest(R, LS, "Skipped (No find method)")
          commandSendKeys = -1
    End Select
          ' driver.FindElementById(actionTarget).SendKeys.Enter
          ' driver.findElement(By.id("id")).sendKeys(Keys.TAB);
          ' driver.findElement(By.id("id")).sendKeys(Keys.SHIFT, Keys.TAB);
          ' driver.findElement(By.id("id")).sendKeys(Keys.CONTROL + "c");

    Exit Function

Err: '----------------------------
    Rtn = errHandler("commandSendKeys", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Function

Private Function commandSelect(findMethod As String, actionTarget As String, actionValue As String, R As ListRow, LS As ListObject)

    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    Select Case findMethod
        Case "Id"
            driver.FindElementById(actionTarget).AsSelect.SelectByText actionValue
        Case "Name"
            driver.FindElementByName(actionTarget).AsSelect.SelectByText actionValue
        Case "XPath"
            driver.FindElementByXPath(actionTarget).AsSelect.SelectByText actionValue
        Case "Css"
            driver.FindElementByCss(actionTarget).AsSelect.SelectByText actionValue
        Case Else
            Call skipTest(R, LS, "Skipped (No find method)")
            commandSelect = -1
    End Select

    Exit Function

Err: '----------------------------
    Rtn = errHandler("commandSelect", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Function

Private Function commandRadio(findMethod As String, actionTarget As String, actionValue As String, R As ListRow, LS As ListObject)

    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    Select Case findMethod
        Case "Id"
            driver.FindElementById(actionTarget).Click
        Case "Name"
            driver.FindElementsByName(actionTarget).Item(actionValue).Click
        Case "XPath"
            driver.FindElementByXPath(actionTarget).Click
        Case "Css"
            driver.FindElementByCss(actionTarget).Click
        Case Else
            Call skipTest(R, LS, "Skipped (No find method)")
            commandRadio = -1
    End Select

    Exit Function

Err: '----------------------------
    Rtn = errHandler("commandRadio", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Function

Private Function commandAlert(findMethod As String, actionTarget As String, actionValue As String, R As ListRow, LS As ListObject)

    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    Select Case actionTarget
        Case "Accept"
            driver.SwitchToAlert.Accept
        Case "Dismiss"
            driver.SwitchToAlert.Dismiss
        Case "SendKeys"
            driver.SwitchToAlert.SendKeys actionValue
        Case Else
            Call skipTest(R, LS, "Skipped (No find method)")
            commandAlert = -1
    End Select

    Exit Function

Err: '----------------------------
    Rtn = errHandler("commandAlert", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Function

Private Function commandMouseMoveTo(findMethod As String, actionTarget As String, R As ListRow, LS As ListObject)

    Dim elm As Object
    
    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    Select Case findMethod
        Case "Id"
           Set elm = driver.FindElementById(actionTarget)
        Case "Name"
           Set elm = driver.FindElementByName(actionTarget)
        Case "XPath"
           Set elm = driver.FindElementByXPath(actionTarget)
        Case "Css"
           Set elm = driver.FindElementByCss(actionTarget)
        Case Else
            Call skipTest(R, LS, "Skipped (No find method)")
            commandMouseMoveTo = -1
    End Select
    driver.Mouse.MoveTo elm
    driver.Wait findElementTimeOut

    Exit Function

Err: '----------------------------
    Rtn = errHandler("commandMouseMoveTo", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Function

Public Sub clearTestResultsConfirm()

    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    If MsgBox("Do you want to clear all of test results? (can't undo this!)", vbOKCancel + vbExclamation + vbDefaultButton2, "test results was initialized") = vbCancel Then
        Exit Sub
    End If

    Call clearTestResults

    Exit Sub

Err: '----------------------------
    Rtn = errHandler("clearTestResultsConfirm", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub
Private Sub clearTestResults()

    Dim LS As ListObject
    Dim R As ListRow
    
    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    Set LS = ActiveSheet.ListObjects(1)
    
    For Each R In LS.ListRows
        Application.StatusBar = "Test results is initializing..." & R.Index & " / " & LS.ListRows.Count
        Call skipTest(R, LS, "")
        R.Range(LS.ListColumns("ErrorMessage").Index) = ""
        R.Range(LS.ListColumns("Memo").Index) = ""
        DoEvents
        DoEvents
    Next R
    
    'ActiveSheet.TextBoxes(1).Text = ""
    Cells(9, 12).Value = ""
    Application.StatusBar = "Ready to run."

    Exit Sub

Err: '----------------------------
    Rtn = errHandler("clearTestResults", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub
Private Sub collectTestResults()

    Dim LS As ListObject
    Dim i As Integer
    Dim rowNum As Long
    
    #If DBG = 0 Then
        On Error GoTo Err
    #End If
    
    Set LS = ActiveSheet.ListObjects("batchRunTBL")
    rowNum = Range("ResultsSummary").Row + 2
    
    For i = 1 To Sheets.Count
        Select Case Sheets(Sheets(i).Name).Name
            Case "BATCH", "LISTBOX_DATA", "UPDATES", "REPORT_RESULTS"
            'ignore theese sheet
            Case Else
            'copy sheet name to listobject
                Cells(rowNum, LS.ListColumns("SheetName").Index) = Sheets(Sheets(i).Name).Name
                Cells(rowNum, LS.ListColumns("Not Tested").Index) = Sheets(Sheets(i).Name).ListObjects(3).Range(2, 2).Text
                Cells(rowNum, LS.ListColumns("Passed").Index) = Sheets(Sheets(i).Name).ListObjects(3).Range(3, 2).Text
                Cells(rowNum, LS.ListColumns("Failed").Index) = Sheets(Sheets(i).Name).ListObjects(3).Range(4, 2).Text
                Cells(rowNum, LS.ListColumns("Skipped").Index) = Sheets(Sheets(i).Name).ListObjects(3).Range(5, 2).Text
                Cells(rowNum, LS.ListColumns("Total").Index) = Sheets(Sheets(i).Name).ListObjects(3).Range(6, 2).Text
                rowNum = rowNum + 1
                DoEvents
                DoEvents
        End Select
    Next i

    Exit Sub

Err: '----------------------------
    Rtn = errHandler("collectTestResults", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub
Public Sub prepTestTarget()

    Dim LS As ListObject
    Dim i As Integer
    Dim rowNum As Long
    
    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    Set LS = ActiveSheet.ListObjects("batchRunTBL")
    rowNum = Range("ResultsSummary").Row + 2
    
    'clear list
    For i = LS.ListRows.Count To 1 Step -1
        LS.ListRows.Item(i).Delete
    Next i
    
    For i = 1 To Sheets.Count
        Select Case Sheets(Sheets(i).Name).Name
            Case "BATCH", "LISTBOX_DATA", "UPDATES", "REPORT_RESULTS"
                'ignore this sheet
            Case Else
                'copy sheet name to listobject
                Cells(rowNum, LS.ListColumns("run target").Index) = "Yes"
                Cells(rowNum, LS.ListColumns("Status").Index) = "Ready to run"
                rowNum = rowNum + 1
                DoEvents
                DoEvents
        End Select
    Next i
    
    Call collectTestResults
    
    Application.StatusBar = "Ready to run."

    Exit Sub

Err: '----------------------------
    Rtn = errHandler("prepTestTarget", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub
Public Sub batchRunTestScriptConfirm()

    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    If MsgBox("Do you want to run batch test script?", vbOKCancel + vbExclamation + vbDefaultButton2, "Run test script") = vbCancel Then
        Exit Sub
    End If
    
    Cells(9, 12).Value = ""
    Call batchRunScript

    Exit Sub

Err: '----------------------------
    Rtn = errHandler("batchRunTestScriptConfirm", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub
Private Sub batchRunScript()

    Dim LS As ListObject
    Dim R As ListRow
    Dim i As Integer
    Dim rowNum As Long
    Dim testTargetSheet As Worksheet

    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    Set LS = ActiveSheet.ListObjects("batchRunTBL")
    
    rowNum = Range("ResultsSummary").Row + 2
    For Each R In LS.ListRows
        
        If LCase(R.Range(LS.ListColumns("run target").Index)) <> "yes" Then
            R.Range(LS.ListColumns("Status").Index) = "Skipped"
            GoTo nextR
        End If
        
        Set testTargetSheet = Worksheets("BATCH")
        testTargetSheet.Activate
        
        Cells(rowNum, LS.ListColumns("Status").Index) = "Testing now"
        Set testTargetSheet = Worksheets(Cells(rowNum, LS.ListColumns("sheetname").Index).Text)
        testTargetSheet.Activate
        
        Call runScript
        
        If Range("ReportResults").Text = "Yes" Then
            'report for each script
            Call reportResults
        End If
        
        Set testTargetSheet = Worksheets("BATCH")
        testTargetSheet.Activate
        Cells(rowNum, LS.ListColumns("Status").Index) = "Finished"
        Cells(rowNum, LS.ListColumns("Lastupdate").Index) = Now()
       
nextR:
        rowNum = rowNum + 1
    Next R
    
    Set testTargetSheet = Worksheets("BATCH")
    testTargetSheet.Activate
    Call collectTestResults
    
    If Range("ReportResults").Text = "Yes" Then
        'report for batch script
        Call reportResults
    End If

    Exit Sub

Err: '----------------------------
    Rtn = errHandler("batchRunScript", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If
    
End Sub

Public Sub reportResults()

    Dim testResults As String
    Dim LS As ListObject
    Dim rowNum As Long
    Dim testTargetSheet  As Worksheet
    
    #If DBG = 0 Then
        On Error GoTo Err
    #End If

    'collect test relusts
    testResults = ""
    testResults = testResults & "/**************************/" & vbCrLf
    testResults = testResults & "Spreadsheetnium TEST RESULT" & vbCrLf
    testResults = testResults & "/**************************/" & vbCrLf
    testResults = testResults & "Test title : " & Range("testTitle").Text & vbCrLf
    testResults = testResults & "sheet name : " & ActiveSheet.Name & vbCrLf & vbCrLf
    
    testResults = testResults & "Browser : " & Range("targetBrowser").Text & vbCrLf
    testResults = testResults & "baseurl : " & Range("baseURL").Text & vbCrLf
    testResults = testResults & "window Width : " & Range("windowSizeW").Text & vbCrLf
    testResults = testResults & "window Height : " & Range("windowSizeH").Text & vbCrLf
    testResults = testResults & "ScreenShot : " & Range("ScreenshotPath").Text & vbCrLf & vbCrLf
    
    testResults = testResults & "/**************************/" & vbCrLf
    testResults = testResults & "Not tested : " & Range("NotTested").Text & vbCrLf
    testResults = testResults & "Passed : " & Range("Passed").Text & vbCrLf
    testResults = testResults & "Failed: " & Range("Failed").Text & vbCrLf
    testResults = testResults & "Skipped : " & Range("Skipped").Text & vbCrLf
    testResults = testResults & "Total : " & Range("Total").Text & vbCrLf
    testResults = testResults & "Progress rate : " & Range("Progressrate").Text & vbCrLf
    testResults = testResults & "/**************************/" & vbCrLf & vbCrLf
    
    
    'testResults = testResults & ActiveSheet.TextBoxes(1).Text & vbCrLf & vbCrLf
    testResults = testResults & Cells(9, 12).Value & vbCrLf & vbCrLf
    
    testResults = testResults & "https://ssugiya.github.io/Spreadsheetnium/" & vbCrLf
    
    #If DBG = 0 Then
        Debug.Print testResults
    #End If
    
    'activate report sheet
    Set testTargetSheet = Worksheets("REPORT_RESULTS")
    testTargetSheet.Activate
    
    'copy results to report sheet
    Set LS = ActiveSheet.ListObjects(1)
    rowNum = Range("TestScript").Row + 3
    Cells(rowNum, LS.ListColumns("Description").Index) = testResults
    
    'run report script
    Call runScript

    Exit Sub

Err: '----------------------------
    Rtn = errHandler("reportResults", Err.Number)
    If Rtn = 0 Then
        Resume Next
    Else
        Call exitProgram
    End If

End Sub


