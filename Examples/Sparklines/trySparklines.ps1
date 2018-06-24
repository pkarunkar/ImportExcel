cls

ipmo $PSScriptRoot\..\..\ImportExcel.psd1 -Force

rm .\sparkLines.xlsx -ErrorAction Ignore

$data = @"
field1,field2
15,7
30,33
35,12
28,1
"@ | ConvertFrom-Csv

$data = 1..50 | % {
    [PSCustomObject][Ordered]@{
        Field1 = Get-Random -Minimum -101 -Maximum 101
        Field2 = Get-Random -Minimum 100 -Maximum 201
        Field3 = Get-Random -Minimum -100 -Maximum 1
    }
}

$xl = $data | Export-Excel .\sparkLines.xlsx -PassThru -AutoNameRange -AutoSize

$ws = $xl.Workbook.Worksheets["Sheet1"]

# Column<->Row
#$sg1 = $ws.SparklineGroups.Add('Line', $ws.Cells["D1:D2"], $ws.Cells["A2:B5"])
#$sg1 = $ws.SparklineGroups.Add('Line', $ws.Cells["D1:D1"], $ws.Cells["A2:A5"])
#$sg2 = $ws.SparklineGroups.Add('Line', $ws.Cells["E1:E1"], $ws.Cells["B2:B5"])

$ws.Cells["E1"].Value = "F1 Spklne"
$sg1 = $ws.SparklineGroups.Add('Line', $ws.Cells["E2:E2"], $ws.Cells["field1"])
$ws.Cells["f1"].Value = "F2 Spklne"
$null = $ws.SparklineGroups.Add('Column', $ws.Cells["F2:F2"], $ws.Cells["field2"])
$ws.Cells["G1"].Value = "F3 Spklne"
$null = $ws.SparklineGroups.Add('Line', $ws.Cells["G2:G2"], $ws.Cells["field3"])

#$sg1 = $ws.SparklineGroups.Add('Line', $ws.Cells["A2:A5"], $ws.Cells["B1:C4"])
#$sg1.DisplayEmptyCellsAs = 'Gap'
#$sg1.Type = 'Line'

# Column<->Column
#$sg2 = $ws.SparklineGroups.Add('Column', $ws.Cells["E1:E2"], $ws.Cells["B2:c5"])

# Row<->Column
#$sg3 = $ws.SparklineGroups.Add('Stacked', $ws.Cells["G1:G2"], $ws.Cells["B2:C5"])
#$sg3 = $ws.SparklineGroups.Add('Stacked', $ws.Cells["A10:B10"], $ws.Cells["B1:C4"])
#$sg3.RightToLeft=$true

# Row<->Row
#$sg4 = $ws.SparklineGroups.Add('Line', $ws.Cells["D10:G10"], $ws.Cells["B1:C4"])
#$ws.Cells["A20"].Value = (get-date "2016/12/30").ToShortDateString()
#$ws.Cells["A21"].Value = (get-date "2017/1/31").ToShortDateString()
#$ws.Cells["A22"].Value = (get-date "2017/2/28").ToShortDateString()
#$ws.Cells["A23"].Value = (get-date "2017/3/31").ToShortDateString()

#$sg4.DateAxisRange = $ws.Cells["A20:A23"]

#$sg4.ManualMax = 5
#$sg4.ManualMin = 3

Close-ExcelPackage $xl -Show