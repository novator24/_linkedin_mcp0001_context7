# 📝 Примеры VBA кода для Context7 MCP Server

## 🎯 Введение

Этот файл содержит примеры VBA кода, которые демонстрируют возможности Context7 MCP Server при работе с Visual Basic for Applications.

## 📊 Excel VBA Примеры

### 1. Работа с диапазонами (Range)

```vba
Sub FormatRange()
    Dim rng As Range
    Set rng = Selection
    
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub AutoFitColumns()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.UsedRange.Columns.AutoFit
End Sub

Sub ClearFormats()
    Dim rng As Range
    Set rng = Selection
    
    rng.ClearFormats
End Sub
```

### 2. Работа с листами (Worksheet)

```vba
Sub AddNewSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "NewSheet_" & Format(Now, "yyyymmdd_hhmmss")
End Sub

Sub DeleteEmptySheets()
    Dim ws As Worksheet
    Dim i As Long
    
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        If ws.UsedRange.Cells.Count = 1 And IsEmpty(ws.UsedRange.Cells(1, 1)) Then
            ws.Delete
        End If
    Next i
End Sub

Sub ProtectAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        ws.Protect Password:="password123"
    Next ws
End Sub
```

### 3. Работа с книгами (Workbook)

```vba
Sub SaveAsPDF()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & ThisWorkbook.Name & ".pdf"
    
    ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath
End Sub

Sub CloseAllWorkbooks()
    Dim wb As Workbook
    
    For Each wb In Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            wb.Close SaveChanges:=False
        End If
    Next wb
End Sub

Sub BackupWorkbook()
    Dim backupPath As String
    backupPath = ThisWorkbook.Path & "\Backup\" & ThisWorkbook.Name & "_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    
    ThisWorkbook.SaveCopyAs Filename:=backupPath
End Sub
```

### 4. Работа с диаграммами (Chart)

```vba
Sub CreateChart()
    Dim cht As Chart
    Dim rng As Range
    
    Set rng = Range("A1:B10")
    Set cht = ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Chart
    
    With cht
        .SetSourceData Source:=rng
        .ChartTitle.Text = "Sample Chart"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Categories"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Values"
    End With
End Sub

Sub FormatChart()
    Dim cht As Chart
    
    Set cht = ActiveChart
    
    With cht
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(240, 240, 240)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Bold = True
    End With
End Sub
```

### 5. Работа со сводными таблицами (PivotTable)

```vba
Sub CreatePivotTable()
    Dim pvt As PivotTable
    Dim pvtCache As PivotCache
    Dim rng As Range
    
    Set rng = Range("A1").CurrentRegion
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rng)
    
    Set pvt = ActiveSheet.PivotTables.Add(PivotCache:=pvtCache, TableDestination:=Range("D1"))
    
    With pvt
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Value").Orientation = xlDataField
    End With
End Sub

Sub RefreshAllPivotTables()
    Dim ws As Worksheet
    Dim pvt As PivotTable
    
    For Each ws In ThisWorkbook.Sheets
        For Each pvt In ws.PivotTables
            pvt.RefreshTable
        Next pvt
    Next ws
End Sub
```

## 📝 Word VBA Примеры

### 1. Работа с документами

```vba
Sub CreateNewDocument()
    Dim doc As Document
    Set doc = Documents.Add
    
    With doc
        .Content.Text = "This is a new document created by VBA."
        .SaveAs2 Filename:="C:\Temp\NewDocument.docx"
    End With
End Sub

Sub FormatDocument()
    Dim doc As Document
    Set doc = ActiveDocument
    
    With doc
        .Content.Font.Name = "Arial"
        .Content.Font.Size = 12
        .Content.ParagraphFormat.Alignment = wdAlignParagraphJustify
    End With
End Sub
```

### 2. Работа с текстом

```vba
Sub FindAndReplace()
    Dim findText As String
    Dim replaceText As String
    
    findText = "old text"
    replaceText = "new text"
    
    With Selection.Find
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub InsertTable()
    Dim tbl As Table
    Dim doc As Document
    
    Set doc = ActiveDocument
    Set tbl = doc.Tables.Add(Range:=Selection.Range, NumRows:=3, NumColumns:=3)
    
    With tbl
        .Cell(1, 1).Range.Text = "Header 1"
        .Cell(1, 2).Range.Text = "Header 2"
        .Cell(1, 3).Range.Text = "Header 3"
    End With
End Sub
```

## 🗄️ Access VBA Примеры

### 1. Работа с базами данных

```vba
Sub CreateTable()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tdf = db.CreateTableDef("NewTable")
    
    Set fld = tdf.CreateField("ID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Name", dbText, 50)
    tdf.Fields.Append fld
    
    db.TableDefs.Append tdf
End Sub

Sub ExecuteQuery()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set qdf = db.CreateQueryDef("")
    qdf.SQL = "SELECT * FROM Customers WHERE City = 'London'"
    
    Set rs = qdf.OpenRecordset()
    
    Do While Not rs.EOF
        Debug.Print rs!CustomerName
        rs.MoveNext
    Loop
    
    rs.Close
End Sub
```

### 2. Работа с формами

```vba
Sub OpenForm()
    DoCmd.OpenForm "CustomerForm", acNormal, "", "", acFormEdit, acWindowNormal
End Sub

Sub CloseAllForms()
    Dim frm As Form
    
    For Each frm In Forms
        DoCmd.Close acForm, frm.Name
    Next frm
End Sub

Sub SetFormProperties()
    Dim frm As Form
    
    Set frm = Screen.ActiveForm
    
    With frm
        .Caption = "Updated Form"
        .Width = 8000
        .Height = 6000
    End With
End Sub
```

## 📊 PowerPoint VBA Примеры

### 1. Работа со слайдами

```vba
Sub AddNewSlide()
    Dim sld As Slide
    Dim ppt As Presentation
    
    Set ppt = ActivePresentation
    Set sld = ppt.Slides.Add(ppt.Slides.Count + 1, ppLayoutText)
    
    With sld
        .Shapes.Title.TextFrame.TextRange.Text = "New Slide"
        .Shapes.Item(2).TextFrame.TextRange.Text = "Slide content goes here"
    End With
End Sub

Sub FormatAllSlides()
    Dim sld As Slide
    Dim shp As Shape
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                With shp.TextFrame.TextRange.Font
                    .Name = "Arial"
                    .Size = 14
                    .Color.RGB = RGB(0, 0, 0)
                End With
            End If
        Next shp
    Next sld
End Sub
```

### 2. Работа с анимацией

```vba
Sub AddAnimation()
    Dim sld As Slide
    Dim shp As Shape
    
    Set sld = ActiveWindow.View.Slide
    Set shp = sld.Shapes(1)
    
    With shp.AnimationSettings
        .EntryEffect = ppAnimEffectFade
        .AdvanceMode = ppAdvanceOnClick
        .AdvanceTime = 0
    End With
End Sub
```

## 📧 Outlook VBA Примеры

### 1. Работа с почтой

```vba
Sub SendEmail()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    
    Set olApp = New Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    With olMail
        .To = "recipient@example.com"
        .Subject = "Test Email from VBA"
        .Body = "This is a test email sent using VBA."
        .Send
    End With
End Sub

Sub ProcessInbox()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.Folder
    Dim olItem As Outlook.MailItem
    
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    
    For Each olItem In olFolder.Items
        If olItem.UnRead Then
            Debug.Print olItem.Subject
            olItem.UnRead = False
        End If
    Next olItem
End Sub
```

### 2. Работа с календарем

```vba
Sub CreateAppointment()
    Dim olApp As Outlook.Application
    Dim olAppt As Outlook.AppointmentItem
    
    Set olApp = New Outlook.Application
    Set olAppt = olApp.CreateItem(olAppointmentItem)
    
    With olAppt
        .Subject = "Meeting with Client"
        .Start = Date + TimeValue("10:00:00")
        .End = Date + TimeValue("11:00:00")
        .Body = "Discuss project requirements"
        .Save
    End With
End Sub
```

## 🔧 Утилиты и вспомогательные функции

### 1. Обработка ошибок

```vba
Sub ErrorHandlerExample()
    On Error GoTo ErrorHandler
    
    ' Ваш код здесь
    Dim result As Integer
    result = 10 / 0 ' Это вызовет ошибку
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Произошла ошибка: " & Err.Description, vbCritical
    Resume Next
End Sub
```

### 2. Логирование

```vba
Sub WriteToLog(message As String)
    Dim logFile As String
    Dim fileNum As Integer
    
    logFile = "C:\Temp\vba_log.txt"
    fileNum = FreeFile
    
    Open logFile For Append As fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & message
    Close fileNum
End Sub
```

### 3. Работа с файлами

```vba
Sub ReadTextFile()
    Dim filePath As String
    Dim fileNum As Integer
    Dim lineText As String
    
    filePath = "C:\Temp\data.txt"
    fileNum = FreeFile
    
    Open filePath For Input As fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, lineText
        Debug.Print lineText
    Loop
    Close fileNum
End Sub

Sub WriteTextFile(content As String)
    Dim filePath As String
    Dim fileNum As Integer
    
    filePath = "C:\Temp\output.txt"
    fileNum = FreeFile
    
    Open filePath For Output As fileNum
    Print #fileNum, content
    Close fileNum
End Sub
```

## 🎯 Интеграция с Context7 MCP

### Пример использования в IDE

```typescript
// Поиск VBA библиотек
const searchResult = await client.callTool("resolve-vba-library", {
  libraryName: "Excel.Worksheet",
  officeApp: "Excel"
});

// Получение документации
const docsResult = await client.callTool("get-vba-docs", {
  vbaLibraryId: "/vba/excel-range",
  topic: "formatting"
});
```

### Категории примеров

1. **Beginner** — базовые операции
2. **Intermediate** — работа с объектами
3. **Advanced** — сложная логика и оптимизация

### Рекомендации по использованию

- Всегда используйте обработку ошибок
- Комментируйте сложные участки кода
- Тестируйте код на небольших данных
- Следуйте принципам DRY (Don't Repeat Yourself)
- Используйте осмысленные имена переменных

---

*Эти примеры демонстрируют возможности VBA и показывают, как Context7 MCP Server может помочь разработчикам получить актуальную документацию и примеры кода.* 