Dim strEmailAddresses As String
Dim FileName, strPath, deno As String
Dim i, n, selectedReportCount As Long
Dim confirmationMsg As VbMsgBoxResult
Dim fileDoesExist As Boolean
Dim objWordApp As Word.Application
Dim objWordDocument As Word.Documet
Dim objSelection As Outlook.Selection
Dim objReport As Outlook.ReportItem
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Public Sub ExtractEmailAddressesFromReports()
    ExceptionAddress = Array("example1@mail.com","example2@mail.com")

    Set ObjSelection = Outlook.Application.ActiveExplorer.Selection
    Call ConfirmationMessage
    Call CreateExcel
    Call CheckFileExist

    If Not (objSelection is Nothing) Then
        i = 0
        n = 1
        On Error Resume Next
        For i = selectedReportCount To 1 Step -1
            Set objReport = objSelection.Item(i)
            Set objWordDocument = objReport.GetInspector.WordEditor
            Set objWordApp = objWordDocument.Application
            Call FindEmailAddress
            While objWordApp.Selection.Find.Found
                strEmailAddresses = objWordApp.Selection.Text
                objWordApp.Selection.Find.Execute
                If IsInArray(strEmailAddresses, ExceptionAddress) = True _
                    Or Len(strEmailAddresses) > 50 Then
                        n = n
                Else
                    Call InputToExcel
                    n = n + 1
                End If
            Wend
            objReport.Close olDiscard
        Next
    End If
    Call RemDuplicate
    Call DisplayExcel
    MsgBox "Completed.", vbInformation, "Extract Email Addresses From Reports"
End Sub

Private Sub ConfirmationMessage()
    selectedReportCount = objSelection.Count
        If selectedReportCount = 1 Then
            deno = "item"
        Else
            deno = "items"
        End If
    confirmationMsg = MsgBox("You selected " & selectedReportCount & " " & deno & "," & vbNewLine & "do you want to continue?", vbOKCancel + vbQuestion, "Extract Email Addresses From Reports"  )
        If confirmationMsg = vbOK Then
            Exit Sub
        Else
            End
        End If
End Sub

Private Sub CreateExcel()
    Set xlApp = CreateObject("Excel.Application")
    FileName = "ExtractedEmailAddresses-" & Format(Date, "YYYYMMDD") & ".xlsx"
    strPath = Environ("USERPROFILE") & "\Desktop\"
    fileDoesExist = Dir(strPath & FileName) > ""
End Sub

Private Sub CheckFileExist()
    If fileDoesExist Then
        Set xlBook = xlApp.Workbooks.Open(strPath & FileName)
        Set xlSheet = xlBook.Sheets(1)
    Else
        Set xlBook = xlApp.Workbooks.Add
        xlBook.SaveAs FileName:=strPath & FileName
        Set xlSheet = xlBook.Sheets(1)
    End If
End Sub

Private Sub FindEmailAddress()
    With objWordApp.Selection.Find
        .Text = "[A-z,0-9,.,_,%,+,-]{1,}\@[A-z,0-9,.,_,%,+,-]{1,}"
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute
    End With
End Sub

Private Sub InputToExcel()
    With xlApp
        With lsBook.Windows(1).Activate
            xlSheet.Cells(n, 1).Value = strEmailAddresses
        End With
    End With
End Sub

Private Sub RemDuplicate()
    With xlBook.Windows(1).Activate
        xlSheet.Range("A:A").RemoveDuplicates Columns:=1
    End With
End Sub

Private Sub DisplayExcel()
    With xlApp
        .Visible = True
        .WindowState = xlMaximized
    End With
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function