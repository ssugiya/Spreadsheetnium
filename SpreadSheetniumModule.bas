Attribute VB_Name = "SpreadSheetniumModule"
Option Explicit

#Const DBG = 0

Public driver As New WebDriver
Public Verify As New Selenium.Verify
Public Const findElementTimeOut As Long = 3000
Public Const passedColorCode As Long = 11854022 'RGB(198, 224, 180)
Public Const failedColorCode As Long = 11389944 'RGB(248, 203, 173)
Public Sub runTestScript()

If MsgBox("Do you want to run test script?", vbOKCancel + vbExclamation + vbDefaultButton2, "Run test script") = vbCancel Then
    Exit Sub
Else
    Call runScript
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


Application.StatusBar = "Test script is initializing."

'==========================================
'Initial settings
'==========================================
targetBrowser = Range("targetBrowser").Text
baseURL = Range("baseURL").Text
windowSizeW = Range("windowSizeW").Text
windowSizeH = Range("windowSizeH").Text
screenshotPath = Range("ScreenshotPath").Text
'[TODO] Select Browser Profile?
'[TOTO] HTTP Header?
'[TODO] Install Browser plag-in

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

    'test results
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

'----------------------------
Err:
    R.Range(LS.ListColumns("ErrorMessage").Index) = Now() & vbCrLf & "Err number: " & Err.Number & vbCrLf & Err.Description
    Application.StatusBar = "Test script finished unexpected"

    Call exitProgram


End Sub

Private Sub skipTest(R As ListRow, LS As ListObject, msg As String)

R.Range(LS.ListColumns("ActualResult").Index) = ""
R.Range(LS.ListColumns("Result").Index) = msg
R.Range(LS.ListColumns("LastUpdate").Index) = ""
R.Range(LS.ListColumns("Result").Index).ClearFormats

End Sub

Private Sub exitProgram()

driver.Quit
ActiveWorkbook.Save

End Sub


Private Function commandClick(findMethod, actionTarget As String, R As ListRow, LS As ListObject)

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

End Function

Private Function commandSubmit(findMethod, actionTarget As String, R As ListRow, LS As ListObject)

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

End Function

Private Function commandSendKeys(findMethod As String, actionTarget As String, actionValue As String, R As ListRow, LS As ListObject)

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

End Function

Private Function commandSelect(findMethod As String, actionTarget As String, actionValue As String, R As ListRow, LS As ListObject)

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

End Function

Private Function commandRadio(findMethod As String, actionTarget As String, actionValue As String, R As ListRow, LS As ListObject)


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

End Function

Private Function commandAlert(findMethod As String, actionTarget As String, actionValue As String, R As ListRow, LS As ListObject)

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

End Function

Private Function commandMouseMoveTo(findMethod As String, actionTarget As String, R As ListRow, LS As ListObject)

Dim elm As Object

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

End Function


Public Sub clearTestResults()

Dim LS As ListObject
Dim R As ListRow

If MsgBox("Do you want to clear all of test results? (can't undo this!)", vbOKCancel + vbExclamation + vbDefaultButton2, "test results was initialized") = vbCancel Then
    Exit Sub
End If

Set LS = ActiveSheet.ListObjects(1)

For Each R In LS.ListRows
    Application.StatusBar = "Test results is initializing..." & R.Index & " / " & LS.ListRows.Count
    Call skipTest(R, LS, "")
    R.Range(LS.ListColumns("ErrorMessage").Index) = ""
    R.Range(LS.ListColumns("Memo").Index) = ""
    DoEvents
    DoEvents
Next R

Application.StatusBar = "Ready to run."

End Sub
Private Sub collectTestResults()

Dim LS As ListObject
Dim i As Integer
Dim rowNum As Long

Set LS = ActiveSheet.ListObjects("batchRunTBL")
rowNum = Range("ResultsSummary").Row + 2

For i = 1 To Sheets.Count
    Select Case Sheets(Sheets(i).Name).Name
        Case "BATCH_RUN", "LISTBOX_DATA", "CheckForUpdates"
            'ignore this sheet
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


End Sub
Public Sub prepTestTarget()

Dim LS As ListObject
Dim i As Integer
Dim rowNum As Long

Set LS = ActiveSheet.ListObjects("batchRunTBL")
rowNum = Range("ResultsSummary").Row + 2

'clear list
For i = LS.ListRows.Count To 1 Step -1
    LS.ListRows.Item(i).Delete
Next i

For i = 1 To Sheets.Count
    Select Case Sheets(Sheets(i).Name).Name
        Case "BATCH_RUN", "LISTBOX_DATA", "CheckForUpdates"
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


End Sub
Public Sub batchRunTestScript()

If MsgBox("Do you want to batch run test script?", vbOKCancel + vbExclamation + vbDefaultButton2, "Run test script") = vbCancel Then
    Exit Sub
Else
    Call batchRunScript
End If

End Sub
Private Sub batchRunScript()

Dim LS As ListObject
Dim R As ListRow
Dim i As Integer
Dim rowNum As Long
Dim testTargetSheet As Worksheet

Set LS = ActiveSheet.ListObjects("batchRunTBL")

rowNum = Range("ResultsSummary").Row + 2
For Each R In LS.ListRows
    
    If LCase(R.Range(LS.ListColumns("run target").Index)) <> "yes" Then
        R.Range(LS.ListColumns("Status").Index) = "Skipped"
        GoTo nextR
    End If
    
    Cells(rowNum, LS.ListColumns("Status").Index) = "Testing now"
    Set testTargetSheet = Worksheets(Cells(rowNum, LS.ListColumns("sheetname").Index).Text)
    testTargetSheet.Activate
    
    Call runScript
    
    Set testTargetSheet = Worksheets("BATCH_RUN")
    testTargetSheet.Activate
    Cells(rowNum, LS.ListColumns("Status").Index) = "Finished"
    Cells(rowNum, LS.ListColumns("Lastupdate").Index) = Now()

nextR:
    rowNum = rowNum + 1
Next R

Call collectTestResults

End Sub
