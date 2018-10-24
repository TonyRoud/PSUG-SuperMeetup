Set-Location 'C:\Users\anthony.roud\Documents\htmltest'

Remove-Variable HTML

# Example using ConvertTo-Html with conditional formatting

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
.red {
    font-weight:bold;
    color:rgb(255, 14, 14);
    background-color:rgb(253, 237, 237);
}
.yellow {
    font-weight:bold;
    color:rgb(255, 166, 0);
    background-color:rgb(255, 249, 223);
}
    </style>
"@

$HTML += "<h1>Environment Health Report</h1><h3>$date</h3>"
$HTML += '<h2>Critical and Warning Events</h2>'

$events = (Get-WinEvent -LogName Application -MaxEvents 200).Where({$level = 'Warning','Error'; $_.LevelDisplayName -in $level -and $_.Message -notmatch 'domain'}) |
    Select-Object TimeCreated, ID, LevelDisplayName, ProviderName, Message -first 10

# Convert to HTML and capture as XML for formatting
[xml]$table1 = $events | ConvertTo-Html -Fragment

# Check each row, skip TH header row
for ($i=1; $i -le $table1.table.tr.count-1; $i++){
    $class = $table1.CreateAttribute("class")

    # Check the value of colum 3 and apply a class to the row if 'Error'
    if ($table1.table.tr[$i].td[2] -match 'Error'){
        $class.value = "red"
        $table1.table.tr[$i].Attributes.Append($class) | Out-Null
    }
}

$HTML += $($table1.InnerXml)

$HTML += "<br><h2>Process Usage</h2><br>"

$processes = Get-Process | Sort-Object PrivateMemorySize -Descending | Select-Object name, PrivateMemorySize -first 10

[xml]$table2 = $processes | ConvertTo-Html -Fragment

# Check each row, skip TH header row
for ($i=1; $i -le $table2.table.tr.count-1; $i++){
    $class = $table2.CreateAttribute("class")

    # Check the value of colum 3 and apply a class to the row if 'Error'
    if (($table2.table.tr[$i].td[1] -as [int]) -ge 300000000){
        $class.value = "red"
        $table2.table.tr[$i].Attributes.Append($class) | Out-Null
    }
    elseif (($table2.table.tr[$i].td[1] -as [int]) -ge 200000000){
        $class.value = "yellow"
        $table2.table.tr[$i].Attributes.Append($class) | Out-Null
    }
}

$HTML += $($table2.InnerXml)

ConvertTo-Html -Body $HTML -Head $style | out-file report_2.htm -force

Invoke-Item .\report_2.htm