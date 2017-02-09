
#file path
$filepath = “C:\test\”


#set outlook to open
$o = New-Object -comobject outlook.application
$n = $o.GetNamespace(“MAPI”)
#Get outlook process for close

$op = Get-Process -Name "Outlook" | Where-Object {$_.MainWindowHandle -eq 0}

$Account = $n.Folders | ? { $_.Name -eq 'codydills@protectplus.com'};
$Inbox = $Account.Folders | ? { $_.Name -match 'Inbox' };
$f = $Inbox.Folders | ? { $_.Name -match 'ToExtract' };


#date string to search for in attachment name
$date = Get-Date -Format MMM-dd-yyyy


#now loop through them and grab the attachments
$f.items | foreach {
    $_.attachments | foreach {
    Write-Host $_.filename
    $a = $_.filename
    If ($a.Contains($date)) {
    $_.saveasfile((Join-Path $filepath $a))
          }
  }
}

#Opens file saved from Get-Attach
$Filetoopen= Join-Path $filepath $a

#Opens Excel object
$excel = New-Object -comobject excel.application

#Get Excel process for closing
$ep = Get-Process -Name "Excel" | Where-Object {$_.MainWindowHandle -eq 0}

#Makes Visible - set false for production
$excel.visible = $False

#Opens Source file in excel and saves to $workbook
$Workbook = $excel.Workbooks.Open($Filetoopen)

#$excel.workbooks.Add()

#Select Worksheet
$ws3 = $Workbook.worksheets.item("Sheet1")

#Select Workbook
#$workbook=$excel.workbooks | Where {$_.name -eq $filetopen}


$xlpivotTableVersion15	   = 5
$xlPivotTableVersion12     = 3
$xlPivotTableVersion10     = 1
$xlCount                 = -4112
$xlDescending             = 2
$xlDatabase                = 1
$xlHidden                  = 0
$xlRowField                = 1
$xlColumnField             = 2
$xlPageField               = 3
$xlDataField               = 4    
$xlDirection        = [Microsoft.Office.Interop.Excel.XLDirection]
# R1C1 means Row 1 Column 1 or "A1"
# R65536C5 means Row 65536 Column E or "E65536"

$range1=$ws3.range("B1")
$range1=$ws3.Range($range1,$range1.End($xlDirection::xlDown))
$range2=$ws3.range("C1")
$range2=$ws3.Range($range2,$range2.End($xlDirection::xlDown))
$selection = $ws3.Range($range1, $range2)


$Workbook2=$excel.Workbooks.Add()
$workbook2=$excel.Workbooks.Item("Book1")
$ws4=$Workbook2.Worksheets.item("Sheet1")

$PivotTable = $Workbook2.PivotCaches().Create($xlDatabase,$selection,$xlPivotTableVersion15)
$PivotTable.CreatePivotTable("R1C1","Tables1") | Out-Null 
[void]$ws4.Select()
$ws4.Cells.Item(3,1).Select()
$Workbook2.ShowPivotTableFieldList = $False

$PivotFields = $ws4.PivotTables("Tables1").PivotFields("Model")

$PivotFields.Orientation = $xlRowField
$PivotFields.Orientation = $xlDataField

$PivotFields = $ws4.PivotTables("Tables1").PivotFields("Reason")

$PivotFields.Orientation = $xlRowField
$PivotFields.Orientation = $xlColumnField

$Pivot=$ws4.PivotTables("Tables1")

$Pivot.TableStyle2 = "PivotStyleLight22"
$Pivot.ShowTableStyleRowStripes = $True

$workbook2.Saveas("C:\test\ModelTicketsByReason_Pivot.xlsx")


Stop-Process $ep
Stop-Process $op
