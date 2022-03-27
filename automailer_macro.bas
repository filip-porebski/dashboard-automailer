Sub StartSub()
    'Show UserForm
    A_Home.Show
End Sub

Sub PrepareFileSub(CompanyCode As String, TopTitle As String, TopSubTitle As String)
    'Declare variables
    Dim ap As Object
    Dim wb, wbkDestination As Workbook
    Dim wsStart, wsDashboard, wsSettings, wsData, wksDestination As Worksheet
    Dim user_name, SelectedCompany As String
    Dim month_digit, month_long, day_shorot, month_short, year_long, myDate As Variant
    Dim alldocuments, std_all, std_1wm, std_2wm, std_3wm, urg_all, urg_1wm, urg_2wm, urg_3wm, inp_all, inp_1wm, inp_2wm, inp_3wm, indexs, duplicates, rejects, utl, reds As Integer
    Dim Attachbacklog As Boolean
    
    Set ap = CreateObject("Excel.Application")
    Set wb = ThisWorkbook
    Set wsDashboard = wb.Sheets("Dashboard")
    Set wsSettings = wb.Sheets("Settings")
    user_name = Application.UserName
    SelectedCompany = CompanyCode

'*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-* WORKING WITH ANOTHER EXCEL DOCUMENT - BACKLOG
    'Open Backlog AS READ ONLY to get the data about documents from it
    Application.ScreenUpdating = False
    
    'Get the variables needed to open the file
    month_digit = Format(Date, "m")
    month_long = Format(Date, "mmmm")
    day_short = Format(Date, "dd")
    month_short = Format(Date, "mm")
    year_long = Format(Date, "yyyy")
    
    'Set your destination
    myDate = day_short & "." & month_short & "." & year_long
    'Change path to your report file
    Set wbkDestination = Workbooks.Open("Y:\" & year_long & "\" & month_digit & ". " & month_long & "\" & myDate & ".xlsx", , 1)
    Set wksDestination = wbkDestination.Sheets("RAW DATA")
    
    alldocuments = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode)
    std_all = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("BU:BU"), "*Std_Prework*")
    std_1wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*NPO*", wksDestination.Range("BU:BU"), "*Std_Prework*")
    std_2wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*2WM*", wksDestination.Range("BU:BU"), "*Std_Prework*")
    std_3wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*3WM*", wksDestination.Range("BU:BU"), "*Std_Prework*")
    urg_all = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("BU:BU"), "*Urg_Prework*", wksDestination.Range("CY:CY"), "")
    urg_1wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*NPO*", wksDestination.Range("BU:BU"), "*Urg_Prework*")
    urg_2wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*2WM*", wksDestination.Range("BU:BU"), "*Urg_Prework*")
    urg_3wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*3WM*", wksDestination.Range("BU:BU"), "*Urg_Prework*")
    inp_all = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("BU:BU"), "*Referr_Input*")
    inp_1wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*NPO*", wksDestination.Range("BU:BU"), "*Referr_Input*")
    inp_2wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*2WM*", wksDestination.Range("BU:BU"), "*Referr_Input*")
    inp_3wm = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*3WM*", wksDestination.Range("BU:BU"), "*Referr_Input*")
    indexs = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("BU:BU"), "*Index*")
    duplicates = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("BU:BU"), "*Duplic*")
    rejects = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("BU:BU"), "*rejct*")
    utl = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("B:B"), "*UTL*")
    reds = Application.WorksheetFunction.CountIfs(wksDestination.Range("G:G"), CompanyCode, wksDestination.Range("CR:CR"), ">=5")
    
    'Check Attach Backlog setting
    If wsSettings.Cells(8, 2).Value = "Yes" And reds > 0 Then
        Attachbacklog = True
    ElseIf wsSettings.Cells(8, 2).Value = "No" Then
        Attachbacklog = False
    End If
    
    'Get the backlog file if selected
    If Attachbacklog = True Then
        Call AttachBacklogSub(SelectedCompany)
    End If
    
    wbkDestination.Close savechanges:=False
    Application.ScreenUpdating = True
'*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-* END OF WORKING WITH ANOTHER EXCEL DOCUMENT - BACKLOG

    'Change the Dashboard data according to user's choice
    With wsDashboard
        .Range("B2").Value = TopTitle
        .Range("B4").Value = TopSubTitle
        .Range("B3").Value = "Report generated: " & Time & " " & day_short & "." & month_short & "." & year_long & " by " & user_name
        
        .Range("C6").Value = alldocuments
        .Range("D6").Value = std_1wm + urg_1wm + inp_1wm
        .Range("E6").Value = std_2wm + urg_2wm + inp_2wm
        .Range("F6").Value = std_3wm + urg_3wm + inp_3wm
        .Range("G6").Value = utl
        .Range("K6").Value = CompanyCode
        
        .Range("H16").Value = reds
        
        .Range("K7").Value = std_all
        .Range("K8").Value = std_1wm
        .Range("K9").Value = std_2wm
        .Range("K10").Value = std_3wm
        
        .Range("K11").Value = urg_all
        .Range("K12").Value = urg_1wm
        .Range("K13").Value = urg_2wm
        .Range("K14").Value = std_3wm
        
        .Range("K15").Value = inp_all
        .Range("K16").Value = inp_1wm
        .Range("K17").Value = inp_2wm
        .Range("K18").Value = inp_3wm
        
        .Range("K20").Value = indexs
        .Range("K21").Value = duplicates
        .Range("K22").Value = rejects
        
        .Range("J22").Value = alldocuments
    End With
End Sub

Sub AttachBacklogSub(SelectedCompany As Variant)
    Dim wbkDestination As Workbook
    Dim xclwbk As Workbook
    
    'Get the variables needed to open the file
    month_m = Format(Date, "m")
    month_mmmm = Format(Date, "mmmm")
    day_dd = Format(Date, "dd")
    month_mm = Format(Date, "mm")
    year_yyyy = Format(Date, "yyyy")
    myDate = day_dd & "." & month_mm & "." & year_yyyy
    DesktopDestination = CreateObject("WScript.Shell").specialfolders("Desktop")
    
    Set wbkDestination = Workbooks(myDate & ".xlsx")
    Set wksDestination = wbkDestination.Sheets("RAW DATA")
    
    Do Until Left(SelectedCompany, 1) <> "0"
        SelectedCompany = Replace(SelectedCompany, "0", "", 1, 1)
    Loop
    
        Sheets("RAW DATA").Select
        Set Source2 = Columns("B:CY").SpecialCells(xlCellTypeVisible)
        Set DBest = Workbooks.Add(xlWBATWorksheet)
    
        Source2.Copy
        With DBest.Sheets(1)
            .Cells(1).PasteSpecial Paste:=8
            .Cells(1).PasteSpecial Paste:=xlPasteValues
            .Cells(1).PasteSpecial Paste:=xlPasteFormats
            .Cells(1).Select
            Application.CutCopyMode = False
        End With
    
    
    
        TempFilePath2 = DesktopDestination & "\"
        TempFileName2 = myDate & " Urgent invoices"
        FileExtStr = ".xlsx": FileFormatNum = 51
    
        With DBest
            .SaveAs TempFilePath2 & TempFileName2 & FileExtStr, FileFormat:=FileFormatNum
        End With
    End With
    
    DBest.Close savechanges:=False
    
End Sub


Sub PrepareFileToSendSub(CompanyCode As String, TopTitle As String, TopSubTitle As String)
    'Declare variables
    Dim ap As Object
    Dim wb As Workbook
    Dim wsStart, wsDashboard, wsSettings, wsData As Worksheet
    Dim DesktopDestination As Variant
    Dim Recipients, Autosend, Exitafter, Attachbacklog As String
    
    Set ap = CreateObject("Excel.Application")
    Set wb = ThisWorkbook
    Set wsDashboard = wb.Sheets("Dashboard")
    Set wsSettings = wb.Sheets("Settings")
    DesktopDestination = CreateObject("WScript.Shell").specialfolders("Desktop")
    Recipients = wsDashboard.Cells(43, 3).Value
    Autosend = wsSettings.Cells(6, 2).Value
    Exitafter = wsSettings.Cells(7, 2).Value
    Attachbacklog = wsSettings.Cells(8, 2).Value

    With wsDashboard
        .Select
        .Range("Dashboard").Activate
        .Range("A1:L31").CopyPicture
        .Range("A1").Select
    End With

    'Paste the copied selected ranges into a temp worksheet
    Set objTempWorkbook = Excel.Application.Workbooks.Add(1)
    Set objTempWorksheet = objTempWorkbook.Sheets(1)
    
    'Paste the picture in Chart area of same dimensions
    Dim rngToPicture As Range
    Set rngToPicture = wsDashboard.Range("Dashboard")
    With objTempWorksheet.ChartObjects.Add(rngToPicture.Left, rngToPicture.Top, rngToPicture.Width, rngToPicture.Height)
        .Activate
        .Chart.Paste
        'Export the chart as PNG File to Temp folder
        objTempWorksheet.Shapes("Chart 1").Line.Visible = msoFalse
        objTempWorksheet.Shapes("Chart 1").Fill.Visible = msoFalse
        .Chart.Export DesktopDestination & "\" & "dashboard_picture" & ".png", "PNG"
    End With
    
    objTempWorkbook.Close savechanges:=False
    ThisWorkbook.Sheets("Start").Select

    'Create a new email
    Set objOutlookApp = CreateObject("Outlook.Application")
    Set objNewEmail = objOutlookApp.CreateItem(olMailItem)
    objNewEmail.HTMLBody = "<body style='background-color: rgb(29, 28, 50); color: rgb(45, 47, 81);'><center><img src='" & DesktopDestination & "\dashboard_picture.png'></center></body>"
    If Attachbacklog = "Yes" Then
        objNewEmail.Attachments.Add DesktopDestination & "\" & Date & " Urgent invoices" & ".xlsx"
    End If
    objNewEmail.Display
    'You can specify the new email recipients, subjects here using the following lines:
    objNewEmail.To = Recipients
    objNewEmail.Subject = Date & " items " & CompanyCode & " " & TopTitle
    objNewEmail.Recipients.ResolveAll
    If Autosend = "Yes" Then
        objNewEmail.Send
    End If

    Kill DesktopDestination & "\" & "dashboard_picture" & ".png"
    If Attachbacklog = "Yes" Then
        Kill DesktopDestination & "\" & Date & " Urgent invoices" & ".xlsx"
    End If

    If Exitafter = "Yes" Then
        ThisWorkbook.Saved = True
        Application.Quit
    End If

End Sub