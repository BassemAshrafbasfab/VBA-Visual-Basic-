# VBA-Visual-Basic-
Currency Converter
Option Explicit

Sub populate_combo_box()

Dim i As Integer
Dim DateArray As Variant
Dim TODAYDATE As String
Dim combobox As String
Dim tWB As Workbook
Set tWB = ThisWorkbook
tWB.Sheets("currencies").Activate
Application.ScreenUpdating = False
'1-الخطوة ديه عشان احدد انهي شيت و انهي خلية بالظبط عشان مايروحش على شيتات تانية او صفحات تانية
Sheets("names").Select
Range("A1").Select
'2- عشان احدد عدد العناصر اللي موجودة في العمود A
For i = 1 To WorksheetFunction.CountA(Columns("A:A"))
    '3- عشان اضيف كل عنصر من العناصر اللي موجودة في العمود // OFFSET -1 عشان في اول خلية مش محتاج ازاحة في اول مرة
    UserForm1.ComboBox1.AddItem ActiveCell.Offset(i - 1, 0) '& "-" & ActiveCell.Offset(I - 1, 1) '<----- SECOND PART 6
    UserForm1.ComboBox2.AddItem ActiveCell.Offset(i - 1, 0) ' & "-" & ActiveCell.Offset(I - 1, 1)
Next i
'4- عشان اول لما افتح اليوزر فورم يظهر العنصر ده افتراضي
UserForm1.ComboBox1.Text = Range("A1") '& "-" & Range("B1") '<------- SECOND PART 7
UserForm1.ComboBox2.Text = Range("A2") '& "-" & Range("B2")
 ' 8- عشان اقسم التاريخ اللي بيظهر من معادلة now الى ثلاث اجزاء
DateArray = Split(Now())
'9-  هنا بقول ان الخانة دايتبوكس = العنصر الصفري في الاتجاه دايتااراي
UserForm1.DateInput = DateArray(0)
'5- عشان اليوزر فورم يظهر
UserForm1.Show





End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
'get web data ----------------------------------------------------------------------

Sub getdatafromweb()


Dim dateday As String
Dim datemonth As String
Dim dateyear As String
Dim firstslash As Integer
Dim secondsl As Integer
Dim DateInput As Date
Dim leen As Integer
Dim leen1 As Integer

DateInput = UserForm1.DateInput

'تقسيييم التاريخ عشان يسمع في url-------------------
firstslash = InStr(DateInput, "/")
secondsl = InStr(firstslash + 1, DateInput, "/")

datemonth = Left(DateInput, firstslash - 1)
leen = Len(datemonth)
If leen = 1 Then datemonth = "0" & datemonth

dateday = Mid(DateInput, firstslash + 1, (secondsl - firstslash) - 1)
leen1 = Len(dateday)
If leen1 = 1 Then dateday = "0" & dateday

dateyear = Right(DateInput, 4)
'-----------------------------------------------------------------------------
'عشان امسح شيت العملات قبل استيراد البيانات الجديدة
Sheets("currencies").Visible = True
Sheets("currencies").Select
Cells.Select
Selection.ClearContents
'-----------------------------------------------------------------------
' كود استيراد البيانات من الويب
Dim url As String
    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & dateyear & "-" & datemonth & "-" & dateday
    With Worksheets("currencies").QueryTables.Add(Connection:=url, Destination:=Worksheets("currencies").Range("A1"))
        .Name = "My Query"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False

    End With
'-----------------------------------------------------------------------------------------------------------------------
' البحث عن قيم العملات وحفظها في متغيرين
Dim initial_exchange_rate As Double, final_exchange_rate As Double
Range("A100").Select
Do
    Selection.Offset(1, 0).Activate
Loop Until Selection = UserForm1.ComboBox1.Value
Selection.Offset(0, 3).Activate
initial_exchange_rate = Selection
Sheets("currencies").Activate
Range("A100").Select
Do
    Selection.Offset(1, 0).Activate
Loop Until Selection = UserForm1.ComboBox2.Value
Selection.Offset(0, 2).Activate
final_exchange_rate = Selection

'MsgBox (initial_exchange_rate)
'MsgBox (final_exchange_rate)
'--------------------------------------------------------------------------
'الحسبة -----------------
Dim result As Double
Dim ans As String
result = UserForm1.Amount * final_exchange_rate / initial_exchange_rate
UserForm1.Conversion = result
Sheets("currencies").Visible = False
Sheets("names").Activate
ans = MsgBox(result)

End Sub
' عمل المخطط
Sub PLOT()
Dim dateday As String
Dim datemonth As String
Dim dateyear As String
Dim firstslash As Integer
Dim secondsl As Integer
Dim DateInput As Date
Dim leen As Integer
Dim leen1 As Integer
Dim TODAYDATE As String
Dim i As Integer
'عشان اكتب التاريخ من تاريخ النهارده لحد 30 يوم ورا في الخلايا ------
Sheets("PLOTT CHART").Visible = True
Sheets("PLOTT CHART").Select
Application.ScreenUpdating = False
Sheets("currencies").Visible = True
Sheets("PLOTT CHART").Visible = True
Sheets("PLOTT CHART").Activate
Cells.Select
Selection.ClearContents
Sheets("Last 30 Day").Delete


'بمسح الخلايا قبل ما اشتغل في الصفحتين ---------------------------------------------
Sheets("currencies").Select
Cells.Select
Selection.ClearContents
Sheets("PLOTT CHART").Select
Cells.Select
Selection.ClearContents
'------------------------------------------------------------------
Range("A30").Select
TODAYDATE = UserForm1.DateInput
For i = 30 To 1 Step -1
    Range("A" & 30 - i + 1) = DateAdd("d", -i + 1, TODAYDATE)
Next i
'Sheets("PLOTT CHART").Visible = False
'------------------------------------------------------------------------------
Dim exchange_rate_1 As Double
Dim exchange_rate_2 As Double
Dim exchange_name As String

For i = 1 To 30
    Sheets("PLOTT CHART").Select
    TODAYDATE = Range("A1:A30").Cells(i, 1)
    
'تقسيييم التاريخ عشان يسمع في url-------------------
    firstslash = InStr(TODAYDATE, "/")
    secondsl = InStr(firstslash + 1, TODAYDATE, "/")

    datemonth = Left(TODAYDATE, firstslash - 1)
    leen = Len(datemonth)
    If leen = 1 Then datemonth = "0" & datemonth

    dateday = Mid(TODAYDATE, firstslash + 1, (secondsl - firstslash) - 1)
    leen1 = Len(dateday)
    If leen1 = 1 Then dateday = "0" & dateday

    dateyear = Right(TODAYDATE, 4)
'اعوض بالي استورته في العنوان عشان اجيب البيانات --------------------------
    Dim url As String
    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & dateyear & "-" & datemonth & "-" & dateday
    With Worksheets("currencies").QueryTables.Add(Connection:=url, Destination:=Worksheets("currencies").Range("A1"))
        .Name = "My Query"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False

    End With
'------------------------------------------------------------------------------
'العملات بالنسبة للدولار الرقم اللي في ثالث عمود-------------------

Sheets("currencies").Activate
Range("A100").Select
Do
    Selection.Offset(1, 0).Activate
Loop Until Selection = UserForm1.ComboBox1.Value
Selection.Offset(0, 2).Activate
exchange_rate_1 = Selection
Range("A100").Select
Do
    Selection.Offset(1, 0).Activate
Loop Until Selection = UserForm1.ComboBox2.Value
Selection.Offset(0, 2).Activate
exchange_rate_2 = Selection
'-------------------------------------------------------------------
' بعوض عن اللي جبته في الصفحة بتاعت الشارت ---------------
 Sheets("PLOTT CHART").Select
    Cells(i, 2) = exchange_rate_1
    Cells(i, 3) = exchange_rate_2
    Cells(i, 4) = exchange_rate_2 / exchange_rate_1
Next i
'----------------------------------------------------------------
' كود رسم المخطط
Sheets("PLOTT CHART").Activate
Range("A:A,D:D").Select
    Range("D1").Activate
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    ActiveChart.SetSourceData Source:=Range("PLOTT CHART!$A:$A,PLOTT CHART!$D:$D")
    ActiveChart.Location Where:=xlLocationAsNewSheet
     ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Last 30 Day"
    
End Sub
