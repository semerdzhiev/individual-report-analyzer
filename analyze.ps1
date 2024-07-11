param(
  [Parameter(Mandatory=$true, HelpMessage="Directory, which contains the Excel files to analyze.")]
  [ValidateScript({ Test-Path -Type Container $_})]
  [string]$InputDirectory,

  [Parameter(Mandatory=$false, HelpMessage="CSV file to store the results in. If no value is specified for this parameter, no CSV file will be generated.")]
  [string]$CsvOutput,

  [Parameter(Mandatory=$false, HelpMessage="XLSX file to store the results in. If no value is specified for this parameter, the file will not be saved.")]
  [string]$XlsxOutput,

  [Parameter(Mandatory=$false, HelpMessage="Display only the totals, without a detailed information about each exam session.")]
  [switch]$TotalsOnly = $false,

  [Parameter(Mandatory=$false, HelpMessage="The norm for auditory hours per academic year")]
  [int]$AuditoryHoursNorm = 270,

  [Parameter(Mandatory=$false, HelpMessage="The norm for all hours per academic year")]
  [int]$TotalHoursNorm = 360
)


#-----------------------------------------------------------------------
# Excel COM object management functions

$excelObject = $null

function Get-ExcelObject {
  if($null -eq $script:excelObject) {
    $script:excelObject = New-Object -ComObject excel.application
    $script:excelObject.visible = $True
  }
  $script:excelObject
}

function Close-ExcelObject {
  if($null -eq $script:excelObject) {
    $script:excelObject.Quit()
    $script:excelObject = $null
  }
}

#-----------------------------------------------------------------------
# Functions for displaying a progress bar
# Based on Graham Gold's answer in this thread:
# https://stackoverflow.com/a/17625116

function Show-Progress {
  [CmdletBinding()]
  param (
      [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
      [PSObject[]]$InputObject,
      [string]$Activity = "Processing items"
  )

      [int]$TotalItems = $Input.Count
      [int]$Count = 0

      $Input | ForEach-Object {
          $_
          $Count++
          [int]$percentComplete = ($Count/$TotalItems* 100)
          Write-Progress -Activity $Activity -PercentComplete $percentComplete -Status ("Working - " + $percentComplete + "%")
      }
}

#-----------------------------------------------------------------------
# Functions for processing the Excel files and creating the report

function Get-NormFromWorkbook {
  param($ExcelObject, $ExcelFile)

  $workbook = $ExcelObject.Workbooks.Open($ExcelFile.FullName)
  $worksheet = $workbook.Worksheets.Item("1. Отчет")

  $hash = [ordered]@{
    Name          = $worksheet.Range("nПреподавателОтчет").Value2
    TotalHours    = $worksheet.Range("оксОЗ").Value2
    THNormMet     = ''
    AuditoryHours = $worksheet.Range("оксАЗ").Value2
    AHNormMet     = ''
    Semester      = $worksheet.Range("F24").Value2.replace("`n", " ")
    AcademicYear  = $worksheet.Range("D24").Value2
    Source        = $ExcelFile.FullName
  }

  $result = New-Object -Type PSObject -Property $hash
  
  $workbook.Close()

  $result
}

function Get-TotalNorms {
  param($Data)

  $totals = @()

  foreach($entry in $data) {
    $totalSum = 0;
    $auditorySum = 0;

    foreach($record in $entry.Group) {
      $totalSum += $record.TotalHours
      $auditorySum += $record.AuditoryHours
    }
     
    $hash = [ordered]@{
      Name          = $entry.Group[0].Name
      TotalHours    = $totalSum
      THNormMet     = $totalSum -ge $TotalHoursNorm
      AuditoryHours = $auditorySum
      AHNormMet     = $auditorySum -ge $AuditoryHoursNorm
      Semester      = "TOTAL"
      AcademicYear  = $entry.Group[0].AcademicYear
      Source        = "Calculated value"
    }

    $totals += New-Object -Type PSObject -Property $hash
  }

  $totals
}

function Publish-ExcelReport {
  param ($ExcelObject, $ReportData)

  $workbook = $ExcelObject.Workbooks.Add()
  $worksheet = $workbook.Worksheets.Item(1)

  $worksheet.Name = "Report"

  # Collect a list of the names of all properties in the object
  # $properties = ($ReportData[0] | get-member | where-object -Property MemberType -eq 'NoteProperty').Name
  # The above method lists the properties alphabetically,
  # so until we find a fix, we have hardcoded them here:
  $properties = @(
    "Name",
    "TotalHours",
    "THNormMet",
    "AuditoryHours",
    "AHNormMet",
    "Semester",
    "AcademicYear",
    "Source"
  )

  # Print the header
  $col = 1
  foreach($p in $properties) {
    $worksheet.Cells.Item(1, $col++) = $p
  }
  
  # Print the data rows
  $row = 2
  foreach($record in $ReportData) {
    $col = 1
    foreach($p in $Properties) {
      $worksheet.Cells.Item($row, $col++) = $record.$p      
    }
    $row++
  }

  $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null

  if($XlsxOutput) {
    $workbook.SaveAs($XlsxOutput)
  }
}

#-----------------------------------------------------------------------
# Script body

# Get a list of all xlsx files in the target directory
$excelWorkbooks = Get-ChildItem -Path $InputDirectory -Filter *.xlsx -File -Recurse

# Process each workbook and export the norm data
$data = $excelWorkbooks
       | Show-Progress -Activity "Processing Excel files"
       | ForEach-Object { Get-NormFromWorkbook -ExcelObject (Get-ExcelObject) -ExcelFile $_ }
       | Sort-Object -Property @{Expression = "Name"; Descending = $false},
                               @{Expression = "Semester"; Descending = $false}
       | Group-Object -Property Name

# Calculate the total norms
$totals = Get-TotalNorms -Data $data

$result = $totals

if( -not $TotalsOnly ) {
  $result = $totals + ($data | Select-Object -ExpandProperty Group)
            | Sort-Object -Property @{Expression = "Name"; Descending = $false},
                                    @{Expression = "Semester"; Descending = $false}
}

if($CsvOutput) {
  $result | Export-Csv -Path $CsvOutput -UseCulture -NoTypeInformation -Encoding utf8BOM
}

Publish-ExcelReport -ExcelObject (Get-ExcelObject) -ReportData $result

# Close-ExcelObject

$result | Select-Object -Property Name,TotalHours,THNormMet,AuditoryHours,AHNormMet,Semester, AcademicYear | Format-Table
