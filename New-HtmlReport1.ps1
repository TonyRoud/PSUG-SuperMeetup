Set-Location "C:\Users\$env:username\OneDrive\Documents\PowerShell\PSUG\SuperMeetup"

$date = Get-Date -Format f
$HTML = ""

$style = @"
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
    width: 80%;
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
#table tr td.green {
    text-align: center;
    color:rgb(0, 61, 18);
    background-color:rgb(219, 255, 230);
}
#table tr td.blank {
    border-top: 1px solid #ffffff;
    border-left: 1px solid #ffffff;
}
</style>
"@

$HTML += "<h1>Environment Health Report</h1><h3>$date</h3>"
$HTML += '<h2>Critical and Warning Events</h2>'

$events = (Get-WinEvent -LogName System -MaxEvents 200).Where({$level = 'Warning','Error'; $_.LevelDisplayName -in $level}) |
    Select-Object TimeCreated, ID, LevelDisplayName, ProviderName, Message -first 10

$HTML += @"
    <table id="table" align="center">
    <tr>
        <th>Time</th>
        <th>ID</th>
        <th>Level</th>
        <th>Provider</th>
        <th>Message</th>
    </tr>
"@

Foreach ($evt in $events){

    $HTML += @"
    <tr>
        <td>$($evt.TimeCreated)</td>
        <td>$($evt.ID)</td>
        <td>$($evt.LevelDisplayName)</td>
        <td>$($evt.ProviderName)</td>
        <td>$($evt.Message)</td>
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
    $HTML += @"
    <tr>
        <td>$($process.Name)</td>
        <td>$($process.ProcessName)</td>
        <td>$($process.PrivateMemorySize)</td>
    </tr>
"@
}

$HTML += "</table>"

ConvertTo-Html -Body $HTML -Head $style | out-file report_1.htm -force

Invoke-Item .\report_1.htm