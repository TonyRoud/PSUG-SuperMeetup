Remove-Variable HTML -ErrorAction SilentlyContinue

# Example using module pshtmltable - Warren Frame - https://github.com/RamblingCookieMonster/PSHTMLTable

$HTML = ""

#get processes to work with
$processes = Get-Process

#Build HTML header
$HTML = New-HTMLHead -title "Environment Health"

#Add CPU time section with top 10 PrivateMemorySize processes.  This example does not highlight any particular cells
$HTML += "<h3>Process Private Memory Size</h3>"
$HTML += New-HTMLTable -inputObject $($processes | Sort-Object PrivateMemorySize -Descending | Select-Object name, PrivateMemorySize -first 10)

#Add Handles section with top 10 Handle usage.
$formatHTML = New-HTMLTable -inputObject $($processes | Sort-Object handles -descending | Select-Object Name, Handles -first 10)

#build hash table with parameters for Add-HTMLTableColor.  Argument and AttrValue will be modified each time we run this.
$params = @{
    Column = "Handles" #I'm looking for cells in the Handles column
    ScriptBlock = {[double]$args[0] -gt [double]$args[1]} #I want to highlight if the cell (args 0) is greater than the argument parameter (arg 1)
    Attr = "Style" #This is the default, don't need to actually specify it here
}

#Add yellow, orange and red shading
$formatHTML = Add-HTMLTableColor -HTML $formatHTML -Argument 1500 -attrValue "background-color:#FFFF99;" @params
$formatHTML = Add-HTMLTableColor -HTML $formatHTML -Argument 2000 -attrValue "background-color:#FFCC66;" @params
$formatHTML = Add-HTMLTableColor -HTML $formatHTML -Argument 3000 -attrValue "background-color:#FFCC99;" @params

#Add title and table
$HTML += "<h3>Process Handles</h3>"
$HTML += $formatHTML

#gather 20 events from the system log and pick out a few properties
$events = (Get-WinEvent -LogName System -MaxEvents 500).Where({$level = 'Warning','Error'; $_.LevelDisplayName -in $level}) |
    sort-object providername -unique | Sort-Object TimeCreated -Descending | Select-Object TimeCreated, ID, LevelDisplayName, ProviderName -First 10

#Create the HTML table without alternating rows, colorize Warning and Error messages, highlighting the whole row.
$eventTable = $events | New-HTMLTable -setAlternating $false |
    Add-HTMLTableColor -Argument "Error" -Column "LevelDisplayName" -AttrValue "background-color:#ffb7a3;" -WholeRow

$HTML += "<h3>System Warning Events</h3>"
$HTML += $eventTable | Close-HTML

# Add content to file and open in browser
Set-Content ExampleReports\html_module_1.htm $HTML
Invoke-Item ExampleReports\html_module_1.htm
