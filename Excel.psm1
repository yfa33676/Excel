# エクセルを起動
$excel = New-Object -ComObject Excel.Application
# エクセル設定
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.ScreenUpdating = $false

function Open-ExcelBook {
    [OutputType([Microsoft.Office.Interop.Excel.WorkbookClass])]
    param(
        [Parameter(ValueFromPipeLine, Position = 0)]
        [string[]]$Path,
        [switch]$ReadOnly,
        [switch]$OutBook
    )
    process{
        Get-Item -Path $Path | ForEach-Object {
            $book = $excel.Workbooks.Open($_, 0, [bool]$ReadOnly)
            if($OutBook){
                $book
            }
        }
    }
}

function Get-ExcelBook {
    [OutputType([Microsoft.Office.Interop.Excel.WorkbookClass])]
    param(
        [Parameter(Position = 0)]
        [string[]]$Name,
        [switch]$Active
    )
    process{
        if($Active){
            $excel.ActiveWorkbook
        } else {
            if($Name){
                $Name | ForEach-Object {$excel.Workbooks.Item($_)}
            } else {
                $excel.Workbooks
            }
        }
    }
}

function Get-ExcelBookName {
    [OutputType([Microsoft.Office.Interop.Excel.WorkbookClass])]
    param(
        [Parameter(Position = 0)]
        [string[]]$Name,
        [switch]$Active
    )
    process{
        if($Active){
            Get-ExcelBook -Active | ForEach-Object{$_.Name}
        } else {
            Get-ExcelBook -Name $Name | ForEach-Object{$_.Name}
        }
    }
}

function Save-ExcelBook {
    param(
        [Parameter(ValueFromPipeLine, Position = 0)]
        [System.__ComObject]$Workbook = $excel.ActiveWorkbook
    )
    process{
        $Workbook.Save()
    }
}

function Close-ExcelBook {
    param(
        [Parameter(ValueFromPipeLine, Position = 0)]
        [System.__ComObject]$Workbook = $excel.ActiveWorkbook,
        [switch]$SaveChanges
    )
    process{
        $Workbook.Close([bool]$SaveChanges)
    }
}

function Get-ExcelSheet {
    [OutputType([Microsoft.Office.Interop.Excel.WorksheetClass])]
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Workbook = $excel.ActiveWorkbook,
        [Parameter(Position = 0)]
        [object[]]$Index,
        [switch]$Active
    )
    process{
        if($Active){
            $Workbook.ActiveSheet
        } else {
            if($Index){
                $Index | ForEach-Object {$Workbook.Worksheets.Item($_)}
            } else {
                $Workbook.Worksheets
            }
        }
    }
}

function Set-ExcelSheetName {
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Worksheet,
        [Parameter(Position = 0)]
        [string]$Name
    )
    process{
        $Worksheet.Name = $Name
    }
}

function Get-ExcelSheetName {
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Workbook = $excel.ActiveWorkbook,
        [Parameter(Position = 0)]
        [object[]]$Index,
        [switch]$Active
    )
    process{
        if($Active){
            $Workbook | Get-ExcelSheet -Active | ForEach-Object {$_.Name}
        } else {
            $Workbook | Get-ExcelSheet -Index $Index | ForEach-Object {$_.Name}
        }
    }
}

function Get-ExcelRangeValue {
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Worksheet = $excel.ActiveWorkbook.ActiveSheet,
        [Parameter(Position = 0)]
        [object]$Range
    )
    process{
        if($Range){
            $Worksheet.Range($Range).Value2
        } else {
            $Worksheet.UsedRange.Value2
        }
    }
}

function Set-ExcelRangeValue {
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Range,
        [Parameter(Position = 0)]
        [string]$Value
    )
    process{
        $Range.Value2 = $Value
    }
}

function Add-ExcelRangeValue {
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Range,
        [Parameter(Position = 0)]
        [string]$Value,
        [string]$Delimiter
    )
    process{
        if($Range.Value2){
            $Value = $Range.Value2 + $Delimiter + $Value
        }
        $Range.Value2 = $Value
    }
}

function Get-ExcelRange {
    [OutputType([Microsoft.Office.Interop.Excel.Range])]
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Worksheet = $excel.ActiveWorkbook.ActiveSheet,
        [Parameter(Position = 0)]
        [object]$Range
    )
    process{
        if($Range){
            $Worksheet.Range($Range)
        } else {
            $Worksheet.UsedRange
        }
    }
}

function Get-Excel{
    [OutputType([Microsoft.Office.Interop.Excel.ApplicationClass])]
    param(
    )
    process{
        $excel
    }
}

function Set-Excel{
    param(
        [switch]$Visible,
        [switch]$ScreenUpdating,
        [switch]$DisplayAlerts,
        [switch]$Quit
    )
    process{
        $excel.Visible = [bool]$Visible
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = [bool]$ScreenUpdating
        if($Quit){
            $excel.Quit()
        }
    }
}

function ConvertFrom-Excel {
    [OutputType([PSCustomObject])]
    param(
        [Parameter(ValueFromPipeLine)]
        [Microsoft.Office.Interop.Excel.ApplicationClass]$excel = $excel
    )
    process{
        [PSCustomObject]@{
            Visible = $excel.Visible
            DisplayAlerts = $excel.DisplayAlerts
            EnableEvents = $excel.EnableEvents
            ScreenUpdating = $excel.ScreenUpdating
            WorkbooksCount = $excel.Workbooks.Count
            ActiveWorkbookName = $excel.ActiveWorkbook.Name
            ActiveSheetName = $excel.ActiveSheet.Name
            ActiveCellAddress = $excel.ActiveCell.Address(0,0)
            ActiveCellValue = $excel.ActiveCell.Value2
        }
    }
}

function ConvertFrom-ExcelBook {
    [OutputType([PSCustomObject])]
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Workbook
    )
    process{
        [PSCustomObject]@{
            BookName = $Workbook.Name
        }
    }
}

function ConvertFrom-ExcelSheet {
    [OutputType([PSCustomObject])]
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Worksheet,
        [switch]$Value,
        [object]$Range
    )
    process{
        if($Value){
            $Worksheet | Get-ExcelRangeValue -Range $Range | Where-Object {$_ -ne $null} | Foreach-Object {
                [PSCustomObject]@{
                    BookName = $Worksheet.Parent.Name
                    SheetName = $Worksheet.Name
                    Value = $_
                }
            }
        } else {
            [PSCustomObject]@{
                BookName = $Worksheet.Parent.Name
                SheetName = $Worksheet.Name
            }
        }
    }
}

function ConvertFrom-ExcelRange {
    [OutputType([PSCustomObject])]
    param(
        [Parameter(ValueFromPipeLine)]
        [System.__ComObject]$Range,
        [switch]$Text
    )
    process{
        if($Text){
            if($Range.Text){
                [PSCustomObject]@{
                    BookName = $Range.Parent.Parent.Name
                    SheetName = $Range.Parent.Name
                    Address = $Range.Address($false, $false)
                    Text = $Range.Text
                }
            }
        } else {
            if($Range.Value2){
                [PSCustomObject]@{
                    BookName = $Range.Parent.Parent.Name
                    SheetName = $Range.Parent.Name
                    Address = $Range.Address($false, $false)
                    Value = $Range.Value2
                }
            }
        }
    }
}