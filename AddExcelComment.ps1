<#
      .Synopsis
        Fills values into a [new] row in an Excel spreadsheet. And sets row formats.
      .Example
$xlf="$env:TEMP\testAddComment.xlsx"

rm $xlf -ErrorAction SilentlyContinue

$pkg = gsv | select -First 5 | Export-Excel $xlf -PassThru


$comments=(
    @{
        ExcelPackage=$pkg
        Range="i6"
        Comment ="This is a comment"
    },

    @{
        ExcelPackage=$pkg
        Range="d4"
        Comment ="This is a comment"
    }
)

foreach($comment in $comments) {
    Add-ExcelComment @comment
}

Close-ExcelPackage $pkg -Show
#>
function Add-ExcelComment {
    param(
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        $Worksheetname = "Sheet1",
        $Range,
        $Comment,
        $Author = "n/a"
    )

    $ws = $ExcelPackage.Workbook.Worksheets[$Worksheetname]
    $null = $ws.Cells[$Range].AddComment($Comment, $Author)
}