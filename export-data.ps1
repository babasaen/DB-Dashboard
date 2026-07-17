# Export Excel to JSON
# Usage: powershell -ExecutionPolicy Bypass -File export-data.ps1

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open("$PSScriptRoot\DB_1412.xlsx")
$ws = $wb.Worksheets.Item('Data')
$data = $ws.UsedRange.Value2
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$headers = @(
  'overallStatus','chemicalStatus','mechanicalStatus','zubbCondition','tisCondition',
  'inspectionLot','operation','inspPoint','heatNo','material','materialDesc','matgrade',
  'inspectDate','zubbValidFrom','zubbValidTo','tisValidFrom','tisValidTo',
  'C','Si','Mn','P','S','Ceq','length','massPerMeter','tranIn','tranHigh',
  'sumWidth','angRB','elongation','inspectCount','zubbFailCount','tisFailCount'
)

$records = [System.Collections.Generic.List[object]]::new()
for ($r = 2; $r -le $data.GetLength(0); $r++) {
  $obj = [ordered]@{}
  for ($c = 1; $c -le 33; $c++) {
    $val = $data[$r, $c]
    if ($null -eq $val) { $val = $null }
    elseif ($val -isnot [double] -and $val -isnot [int]) {
      $val = [string]$val
      if ($val -eq '') { $val = $null }
    }
    $obj[$headers[$c - 1]] = $val
  }
  $records.Add([pscustomobject]$obj)
}

$output = @{
  exportedAt   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  sourceFile   = 'DB_1412.xlsx'
  totalRecords = $records.Count
  records      = $records
}

$jsonPath = Join-Path $PSScriptRoot 'data.json'
$output | ConvertTo-Json -Depth 5 -Compress | Set-Content -Path $jsonPath -Encoding UTF8
Write-Host "Exported $($records.Count) records to $jsonPath"
