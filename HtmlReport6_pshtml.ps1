Set-Location "C:\Users\$env:username\OneDrive\Documents\PowerShell\PSUG\SuperMeetup"

# Example using pshtml

$proc = Get-Process | Select-Object -First 10 Id, Name, Handles, StartTime, WorkingSet
$css = 'body{background:#252525;font:87.5%/1.5em Lato,sans-serif;padding:20px}table{border-spacing:1px;border-collapse:collapse;background:#F7F6F6;border-radius:6px;overflow:hidden;max-width:800px;width:100%;margin:0 auto;position:relative}td,th{padding-left:8px}thead tr{height:60px;background:#367AB1;color:#F5F6FA;font-size:1.2em;font-weight:700;text-transform:uppercase}tbody tr{height:48px;border-bottom:1px solid #367AB1;text-transform:capitalize;font-size:1em;&:last-child {;border:0}tr:nth-child(even){background-color:#E8E9E8}'

html {
    head {
        style { $css }
    }
    body { ConvertTo-HTMLtable -Object $proc }
}  | Out-File ExampleReports\report_6.htm

Invoke-Item ExampleReports\report_6.htm