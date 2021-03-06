Remove-Variable HTML -ErrorAction SilentlyContinue

# Example using manually generated HTML with conditional formatting

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
#table {
    font-family: arial, sans-serif;
    border-collapse: collapse;
    width: 70%;
    align: center;
}
#table th {
    border: 1px solid #dddddd;
    font-weight:bold;
    color:#000000;
    background-color:#e9f2ff;
    padding: 8px;
}
#table tr td {
    border: 1px solid #dddddd;
    text-align: center;
    padding: 8px;
}
#table tr:hover {background-color: rgb(236, 234, 234);}
#table tr.red {
    border: 1px solid #e4e4e4;
    font-weight:bold;
    text-align: center;
    color:rgb(255, 14, 14);
    background-color:rgb(253, 237, 237);
}
#table tr td.red {
    border: 1px solid #e4e4e4;
    font-weight:bold;
    text-align: center;
    color:rgb(255, 14, 14);
    background-color:rgb(253, 237, 237);
}
#table tr td.yellow {
    font-weight:bold;
    text-align: center;
    color:rgb(255, 166, 0);
    background-color:rgb(255, 249, 223);
}
</style>
"@

$HTML += "<h1>Environment Health Report</h1><h3>$date</h3>"

$HTML += '<h2>Critical and Warning Events</h2>'

$events = (Get-WinEvent -LogName System -MaxEvents 500).Where({$level = 'Warning','Error'; $_.LevelDisplayName -in $level}) |
    sort-object providername -unique | Sort-Object TimeCreated -Descending | Select-Object TimeCreated, ID, LevelDisplayName, ProviderName -First 10

$HTML += @"
    <table id="table" align="center">
    <tr>
        <th>Time</th>
        <th>ID</th>
        <th>Level</th>
        <th>Provider</th>
    </tr>
"@

Foreach ($evt in $events){

    $class = ""

    if ($evt.LevelDisplayName -match 'Error'){ $Class = ' Class="red"' }

    $HTML += @"
    <tr$class>
        <td>$($evt.TimeCreated)</td>
        <td>$($evt.ID)</td>
        <td>$($evt.LevelDisplayName)</td>
        <td>$($evt.ProviderName)</td>
    </tr>
"@
}

$HTML += "</table>"
$HTML += "<h2>Process Usage</h2>"

$processes = Get-Process | Sort-Object PrivateMemorySize -Descending | Select-Object name, PrivateMemorySize, ProcessName -first 10

$HTML += @"
    <table id="table" align="center">
    <tr>
        <th>Name</th>
        <th>ProcessName</th>
        <th>PrivateMemorySize</th>
    </tr>
"@

Foreach ($process in $processes){

    $class = ""

    if ($process.PrivateMemorySize -gt 200000000){ $Class = ' Class="yellow"' }
    if ($process.PrivateMemorySize -gt 300000000){ $Class = ' Class="red"' }

    $HTML += @"
    <tr>
        <td>$($process.Name)</td>
        <td>$($process.ProcessName)</td>
        <td$class>$($process.PrivateMemorySize)</td>
    </tr>
"@
}

$HTML += "</table>"

ConvertTo-Html -Body $HTML -Head $style | out-file ExampleReports\manual_formatted.htm -force

Invoke-Item ExampleReports\manual_formatted.htm

# Example report fleshed out with additional tables and data
# Invoke-Item ExampleReports\manual_html_full.html