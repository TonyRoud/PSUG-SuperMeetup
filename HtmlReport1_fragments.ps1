Set-Location "C:\Users\$env:username\OneDrive\Documents\PowerShell\PSUG\SuperMeetup"
Remove-Variable HTML -ErrorAction SilentlyContinue

# Example using ConvertTo-Html

$date = Get-Date -Format f
$HTML = ""

$style = @"
<title>Environment Health Report</title>
<style>
body {
    color:#333333;
    font-family:Calibri,Tahoma;
    font-size: 8pt;
}
h1 {
    text-align:center;
    font-size: 16pt;
}
h2 {
    color:#5674fa;
    text-align:center;
    padding-top: 8pt;
    font-size: 12pt;
}
h3 {
    text-align:center;
    font-size: 10pt;
}
p {
    text-align:center;
    font-size: 10pt;
}
table {
    font-family: arial, sans-serif;
    border-collapse: collapse;
    width: 70%;
    margin-left:auto;
    margin-right:auto;
}
table th {
    border: 1px solid #dddddd;
    font-weight:bold;
    color:#000000;
    background-color:#e9f2ff;
    padding: 8px;
}
table tr td {
    border: 1px solid #dddddd;
    text-align: center;
    padding: 8px;
}
table tr:hover {background-color: rgb(236, 234, 234);}
table tr td.red {
    border: 1px solid #e4e4e4;
    font-weight:bold;
    text-align: center;
    color:rgb(255, 14, 14);
    background-color:rgb(253, 237, 237);
}
table tr td.yellow {
    font-weight:bold;
    text-align: center;
    color:rgb(255, 166, 0);
    background-color:rgb(255, 249, 223);
}
table tr td.blank {
    border-top: 1px solid #ffffff;
    border-left: 1px solid #ffffff;
}
</style>
"@

$HTML += "<h1>Environment Health Report</h1><h3>$date</h3>"

# Capture events to variable and filter down to key info before converting
$events = (Get-WinEvent -LogName Application -MaxEvents 200).Where({$level = 'Warning','Error'; $_.LevelDisplayName -in $level -and $_.Message -notmatch 'domain'}) |
    Select-Object TimeCreated, ID, LevelDisplayName, ProviderName, Message -first 10

# Convert to HTML and capture as XML for formatting
$HTML += $events | ConvertTo-Html -Fragment -PreContent '<h2>Critical and Warning Events</h2>'

# Capture process info into variable
$processes = Get-Process | Sort-Object PrivateMemorySize -Descending | Select-Object name, PrivateMemorySize, ProcessName -first 10

# Convert to HTML and capture as XML for formatting
$HTML += $processes | ConvertTo-Html -Fragment -PreContent "<h2>Process Usage</h2>"

# Convert all to HTML in one go and output to file
ConvertTo-Html -Body $HTML -Head $style | out-file ExampleReports\report_1.htm -force

Invoke-Item ExampleReports\report_1.htm