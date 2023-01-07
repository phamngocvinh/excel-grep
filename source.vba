' Global variables
Dim resultCount, resultCell, cellChk, commChk, shapeChk, fileCount, processCount As Integer
Dim statusbarStr As String

' Constant
Const folderCell As String = "C2"
Const excludeCell As String = "C8"
Const headerFirstCell As String = "B2"
Const headerRowCell As String = "B2:E2"
Const headerColCell As String = "B:E"
Const headerCellCell As String = "B2"
Const headerValueCell As String = "C2"
Const headerSheetCell As String = "D2"
Const headerFileCell As String = "E2"
Const searchStrCell As String = "C10"

' Browse button click
Sub browse_Click()
    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select path"
        .ButtonName = "Select"
        If .Show = -1 Then ' If OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then
        Range(folderCell).value = sFolder
    End If
End Sub

' Grep button click
Sub grep_Click()
    ' Excel optimize
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.FindFormat.Clear
    
    Dim FileSystem As Object
    Dim HostFolder As String
    
    fileCount = 0
    processCount = 0
    resultCount = 0
    resultCell = 2

    ' Get option checkbox value
    cellChk = ThisWorkbook.Worksheets(1).Shapes("Check Box 3").OLEFormat.Object.value
    commChk = ThisWorkbook.Worksheets(1).Shapes("Check Box 4").OLEFormat.Object.value
    shapeChk = ThisWorkbook.Worksheets(1).Shapes("Check Box 5").OLEFormat.Object.value
    
    ' Create Result sheet
    Call CreateResultSheet
    
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    ' Get number of files
    CountFiles FileSystem.GetFolder(Range(folderCell).value)
    ' Loop through folder and subfolder
    DoFolder FileSystem.GetFolder(Range(folderCell).value)
    
    If resultCount > 0 Then
        MsgBox ("Complete!")
        ' Hide status bar
        Application.statusbar = False
        
        Dim wsr As Worksheet
        Set wsr = ThisWorkbook.Sheets(2)
        wsr.Select
        wsr.Columns(headerColCell).AutoFit
        
        Call AddBorder
        
        ' Scroll to first cell
        Application.GoTo Reference:=Range("A1"), Scroll:=True
    Else
        MsgBox ("Not found!")
    End If
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' Loop through folder and subfolder
Function DoFolder(Folder)
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim searchString, excludeArr
    searchString = Range(searchStrCell).value
    
    ' Get excluded file
    excludeArr = Split(Range(excludeCell).value, ",")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim SubFolder
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
    Next
    
    Dim file
    For Each file In Folder.Files
        ' Update status bar
        processCount = processCount + 1
        statusbarStr = "Process: " & processCount & "/" & fileCount & "    " & file.Path
        Application.statusbar = statusbarStr
        
        ' Ignore excluded file
        If IsInArray(file.name, excludeArr) Then
            GoTo ContinueLoop
        End If
    
        ' Operate on each file
        Dim fileExt As String
        fileExt = fso.GetExtensionName(file)
        If fileExt = "xlsx" Or fileExt = "xls" Then
            Set wb = Workbooks.Open(file)
            For Each ws In ActiveWorkbook.Worksheets
                statusbarStr = statusbarStr & "."
                If cellChk = 1 Then
                    Application.statusbar = statusbarStr
                    Call CellSearch(file, ws, searchString)
                End If
                If commChk = 1 Then
                    Application.statusbar = statusbarStr
                    Call CommentSearch(file, ws, searchString)
                End If
                If shapeChk = 1 Then
                    Application.statusbar = statusbarStr
                    Call ShapeSearch(file, ws, searchString)
                End If
            Next
            wb.Close savechanges:=False
        End If
ContinueLoop:
    Next
End Function

Function CountFiles(Folder)
    
    Dim SubFolder
    For Each SubFolder In Folder.SubFolders
        CountFiles SubFolder
    Next
    Dim file
    For Each file In Folder.Files
        fileCount = fileCount + 1
    Next
End Function

Function CellSearch(file, Worksheet, searchString)
    Dim cl As Range
    
    ' Find first instance on sheet
    Set cl = Worksheet.Cells.Find(What:=searchString, _
        After:=Worksheet.Cells(1, 1), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
    If Not cl Is Nothing Then
        ' if found, remember location
        FirstFound = cl.Address
        ' format found cell
        Do
            resultCell = resultCell + 1
            Call WriteResult(resultCell, Replace(cl.Address, "$", ""), cl.value, Worksheet.name, file.Path)
            ' find next instance
            Set cl = Worksheet.Cells.FindNext(After:=cl)
            ' repeat until back where we started
        Loop Until FirstFound = cl.Address
    End If
End Function

Function ShapeSearch(file, Worksheet, searchString)
    Dim shape As shape
    Dim shapeStr As String
    
    For Each shape In Worksheet.Shapes
        If Not shape.Type = msoComment Then
            On Error Resume Next
            shapeStr = shape.TextFrame.Characters.Text
            On Error GoTo 0
            If Not InStr(shapeStr, searchString) = 0 Then
                resultCell = resultCell + 1
                Call WriteResult(resultCell, ColNumToLetter(shape.TopLeftCell.Column) & shape.TopLeftCell.Row, shapeStr, Worksheet.name, file.Path)
            End If
        End If
    Next
End Function

Function CommentSearch(file, Worksheet, searchString)

    Dim comment As comment
    Dim commentStr As String
    
    For Each comment In Worksheet.Comments
        On Error Resume Next
        commentStr = comment.Text
        On Error GoTo 0
        
        If Not InStr(commentStr, searchString) = 0 Then
            resultCell = resultCell + 1
            Call WriteResult(resultCell, ColNumToLetter(comment.shape.TopLeftCell.Column - 1) & comment.shape.TopLeftCell.Row + 1, commentStr, Worksheet.name, file.Path)
        End If
    Next
    

End Function

Function WriteResult(loc, resCell, resValue, resSheet, resBook)
    resultCount = resultCount + 1
    Dim wsr As Worksheet
    Set wsr = ThisWorkbook.Sheets(2)
    wsr.Activate
    wsr.Range("B" & loc).value = resCell
    wsr.Range("C" & loc).value = resValue
    wsr.Range("D" & loc).value = resSheet
    ActiveSheet.Hyperlinks.Add Anchor:=Range("E" & loc), Address:="file:///" & resBook, SubAddress:="'" & resSheet & "'" & "!" & resCell, TextToDisplay:=resBook
End Function

Function IsSheetExists()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Result")
    On Error GoTo 0
    
    If Not ws Is Nothing Then IsSheetExists = True
End Function

Function ColNumToLetter(colNum)
    ColNumToLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

Function CreateResultSheet()
    If Not IsSheetExists Then
        ThisWorkbook.Sheets.Add(After:=Sheets(1)).name = "Result"
    End If
    
    Dim wsr As Worksheet
    Set wsr = ThisWorkbook.Sheets(2)
    wsr.UsedRange.Delete
    wsr.Range(headerCellCell).value = "Cell"
    wsr.Range(headerValueCell).value = "Value"
    wsr.Range(headerSheetCell).value = "Sheet"
    wsr.Range(headerFileCell).value = "File"
    wsr.Range(headerRowCell).Interior.Color = RGB(46, 52, 64)
    wsr.Range(headerRowCell).Font.Color = vbWhite
    wsr.Range(headerRowCell).Font.Bold = True
    
    ThisWorkbook.Worksheets(1).Select
End Function

Function AddBorder()
    Dim wsr As Worksheet
    Set wsr = ThisWorkbook.Sheets(2)
    Set lastCell = wsr.UsedRange.Cells(wsr.UsedRange.Cells.Count)
    With wsr.Range(headerFirstCell, lastCell).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function


