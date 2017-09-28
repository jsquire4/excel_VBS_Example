Sub autoFormatFeedback()

    Dim raw As Worksheet
    Dim prs As Worksheet
    Dim lastrow As Long, i As Long
    Dim cutOffDate As Date
    Dim today As String
    today = Date

    Set prs = Sheets("Sheet1")
    Set raw = Sheets.Add
    raw.Name = "Temp"

    'Auto add info to second

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;/Users/JSquire/Desktop/feedbackReport.txt", Destination:=Range("A1"))
        .Name = "fb"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .saveData = True
        .AdjustColumnWidth = True
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlMacintosh
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
        .Refresh BackgroundQuery:=False
        .UseListObject = False
    End With
    raw.Select
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.ScrollColumn = 6
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Class Type"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Class Name"
    Range("I1").Select
    
    
    lastrowRaw = raw.Range("A" & Rows.Count).End(xlUp).Row
    lastrowParsed = prs.Range("A" & Rows.Count).End(xlUp).Row

    
    'instigate userId and fbInstanceId so that the loop catches the inequality

    userId = -1
    fbInstanceId = -1

    'add search function to find all feedback objects by id (use find function to find by class id, date, or user.  Push to array, print array of needed feedback objects)
    fbId = 0

    'add iterator to remember row number on the parsed sheet (pr = parsed row)
    pr = -2

    'explained below
    lastColumn = 6

    'Anserable format is not useful, replace question type to get better format
    raw.Columns(6).Replace What:="*isagree*Agree*", Replacement:="radio"
    raw.Columns(6).Replace What:="*rue*alse*", Replacement:="true-false"
    raw.Columns(6).Replace What:="*es*o*", Replacement:="true-false"
    
    For i = 2 To lastrowRaw


        'Get all current values for raw data at i
        curfbInstanceId = raw.Cells(i, 1)
        curUserId = raw.Cells(i, 2)
        curDate = raw.Cells(i, 3)
        curClassId = raw.Cells(i, 4)
        curQ = raw.Cells(i, 5)
        curQType = raw.Cells(i, 6)
        curRes = raw.Cells(i, 7)
        curType = raw.Cells(i, 8)
        curClassName = raw.Cells(i, 9)


        '*** Test if iterator has found new user and feedback instance
        
        If (curfbInstanceId <> fbInstanceId) And (curUserId <> userId) Then
            
            'create new "object" on sheet by getting userId and feedback instance id from raw sheet
            pr = pr + 3
           
            userId = curUserId
            fbInstanceId = curfbInstanceId
            fbId = fbId + 1

            'next row down from current Row
            resPr = pr + 1


            'prefill titles
            
            prs.Cells(pr, 1) = "Date"
                
            prs.Cells(pr, 2) = "Class Name"
                
            prs.Cells(pr, 3) = "Class Id"
            
            prs.Cells(pr, 4) = "User Id"
           
            prs.Cells(pr, 5) = "Feedback Id"

            'fill available data at one row below title (resPr)

            prs.Cells(resPr, 1) = curDate
            
            prs.Cells(resPr, 2) = curClassName
            
            prs.Cells(resPr, 3) = curClassId
           
            prs.Cells(resPr, 4) = curUserId
            
            prs.Cells(resPr, 5) = fbInstanceId
                

            'add iterator for remembering column number; start at 6 because it is one after feeback id column
            pc = 6

            prs.Cells(pr, pc) = curQ
            
            prs.Cells(resPr, pc) = parsedResponse(curQType, curRes)
            
            pc = pc + 1

        '*** Otherwise add to the current object
        
        Else
            'get and parse repsonses based off of ValidAns and Value value and paste under correct object
            
            prs.Cells(pr, pc) = curQ
                
            prs.Cells(resPr, pc) = parsedResponse(curQType, curRes)
                

            pc = pc + 1

            'need to keep track of column range because apparently excel isnt able to
            If pc >= lastColumn Then
                lastColumn = pc
            End If

        End If

    Next i

    lastrowParsed = prs.Range("A" & Rows.Count).End(xlUp).Row

    For i = 6 To lastColumn
        prs.Columns(i).ColumnWidth = 32
    Next i

    For i = 1 To lastrowParsed
    
        If prs.Cells(i, 1) = "Date" Then
            addBorders i, lastColumn, False
        Else
            addBorders i, lastColumn, True
        End If

    Next i

    prs.Cells.WrapText = True
    prs.Cells.VerticalAlignment = xlTop
    prs.Cells.HorizontalAlignment = xlLeft
    prs.Activate
    Application.DisplayAlerts = False
    raw.Delete
    Application.DisplayAlerts = True

    ' added buttons for search functions

    
    prs.Rows("1:3").EntireRow.Insert
    prs.Rows("1:3").RowHeight = 65

    Range("F2").Select
    ActiveSheet.Buttons.Add(46, 40, 296, 133).Select
    Selection.OnAction = "filterDatesFeeback"
    Selection.Characters.Text = "Button 4"
    With Selection.Characters(Start:=1, Length:=8).Font
        .Name = "Helvetica"
        .FontStyle = "Regular"
        .Size = 12
        .StrikeThrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Selection.Characters.Text = "Search Feedback By Date"
    With Selection.Characters(Start:=1, Length:=23).Font
        .Name = "Helvetica"
        .FontStyle = "Regular"
        .Size = 12
        .StrikeThrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("F2").Select
    ActiveSheet.Buttons.Add(384, 47, 319, 126).Select
    Selection.OnAction = "searchByClassFeeback"
    ActiveSheet.Shapes("Button 5").ScaleHeight 1.0555555556, msoFalse, _
        msoScaleFromBottomRight
    Selection.Characters.Text = "Button 5"
    With Selection.Characters(Start:=1, Length:=8).Font
        .Name = "Helvetica"
        .FontStyle = "Regular"
        .Size = 12
        .StrikeThrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Selection.Characters.Text = "Search Feedback by Class"
    With Selection.Characters(Start:=1, Length:=24).Font
        .Name = "Helvetica"
        .FontStyle = "Regular"
        .Size = 12
        .StrikeThrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    ActiveSheet.Buttons.Add(726, 39, 326, 130).Select
    Selection.OnAction = "searchByUserFeeback"
    Selection.Characters.Text = "Button 6"
    With Selection.Characters(Start:=1, Length:=8).Font
        .Name = "Helvetica"
        .FontStyle = "Regular"
        .Size = 12
        .StrikeThrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Selection.Characters.Text = "Search Feedback by User"
    With Selection.Characters(Start:=1, Length:=23).Font
        .Name = "Helvetica"
        .FontStyle = "Regular"
        .Size = 12
        .StrikeThrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    ActiveSheet.Shapes("Button 6").ScaleHeight 1.0384615385, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Button 6").ScaleHeight 0.9851851852, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes.Range(Array("Button 5")).Select
    ActiveSheet.Shapes("Button 5").IncrementLeft -8
    ActiveSheet.Shapes("Button 5").IncrementTop -1
    Range("J2").Select
End Sub

Sub autoFormatGrades()


    Dim ws As Worksheet, raw As Worksheet
        Dim lastrow As Long, i As Long
        Dim cutOffDate As Date, today As Date
        Dim preScores() As long, postScores() As long
        Dim preLow As long, preHigh As long, preAve As long, preCount As Integer
        Dim postLow As long, postHigh As long, postAve As long, postCount As Integer
        Set ws = Sheets("Sheet1")
        Set raw = Sheets.Add

        raw.Name = "gradesTemp"
        raw.Activate
    Range("A1").Select
        With raw.QueryTables.Add(Connection:= _
            "TEXT;/Users/JSquire/Desktop/gradeReport.txt", Destination:=Range("A1"))
            .Name = "grades_1"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .saveData = True
            .AdjustColumnWidth = True
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = xlMacintosh
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
            .Refresh BackgroundQuery:=False
            .UseListObject = False
        End With

        Rows("1:1").Select
            Selection.Delete Shift:=xlUp
            Columns("H:H").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            ActiveCell.FormulaR1C1 = "Post-Test"
            Columns("G:G").Select
            Columns("H:H").EntireColumn.AutoFit
            ActiveCell.FormulaR1C1 = "Pre-Test"
            Range("K3").Select
           
            Columns("I:I").Select
            Selection.NumberFormat = "mm/dd/yy;@"
            raw.Sort.SortFields.Clear
            raw.Sort.SortFields.Add Key:=Range("I:I" _
                ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With raw.Sort
            .SetRange Range("A:I")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        Columns("B:B").Select
        raw.Sort.SortFields.Clear
        raw.Sort.SortFields.Add Key:=Range( _
            "B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With raw.Sort
            .SetRange Range("A:H")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

    raw.Columns(6).Replace What:="*re-?est*", Replacement:="pre"
    raw.Columns(6).Replace What:="*ost-?est*", Replacement:="post"
    lastrow = raw.Range("A" & Rows.Count).End(xlUp).Row

    curClass = -1

    For i = 2 To lastrow

        tName = raw.Cells(i, 6)
        score = raw.Cells(i, 7)
        courseId = raw.Cells(i, 2)

        if (courseId <> curClass) Then
             
            if (i <> 2) then
                writeResults(preScores, postScores)
            end if

            'reset counter and comp variables
            curClass = courseId
            preCount = 0
            postCount = 0

            ReDim preScores(preCount)
            Redim postScores(postCount)
           
            If (tName = "pre") then

                preScores(preCount) = score
                preCount = preCount + 1

            Elseif (tName = "post") then

                postScores(postCount) = score
                postCount = postCount + 1

            End If
            

        Elseif (courseId = curClass) Then

            If (tName = "pre") then

                ReDim Preserve preScores(preCount)
                preScores(preCount) = score
                preCount = preCount + 1

            Elseif (tName = "post") then

                ReDim Preserve postScores(postCount)
                postScores(postCount) = score
                postCount = postCount + 1

            End if      

        End if
            
    Next i
 
End Sub
