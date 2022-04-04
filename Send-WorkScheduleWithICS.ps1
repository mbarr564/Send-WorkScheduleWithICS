#Intended for use by restaurants with variable weekly scheduling spreadsheets
#https://en.wikipedia.org/wiki/ICalendar
#https://gist.github.com/nyanhp/20ff0edb7c78cdb08375a15826e47da2
# original script outline:
# 1. convert from native format to csv
# 2. convert csv table into @(PSCustomObject) array
# 3. convert 24hr and add datetime range scheduling NoteProterties to PSCustomObject array
# 4. consume email address ID=value list and add as NoteProperty to PSCustomObject array
# 5. foreach employee scheduling object element {
#     i. generate ICS file from an input PSCustomObject with scheduling NoteProperties}
# 6. foreach employee scheduling object element, generate email body with schedule, attach individual's generated ICS file {
#     i. Send-MailMessage -Body $msgBody -Attachment $ICSfile -Subject "$($employee.Name)'s schedule for week starting $weekStartDate"}
# 7. foreach $employee Where $_.EmailAdddress -eq '' {Send-MailMessage -To $adminEmail -Body $adminErrors}

param ([string]$CSVPath = "$PSScriptRoot\schedule_template.csv", [ValidateSet('CSV','McD','BK','Wend','DQ','TacT','TacB','SubW')][string]$Type = 'CSV')

## Init
[string[]]$adminErrors = @()
$VerbosePreference = 'Continue'

## Import schedule file
## Returned columns: "Name", "EmployeeID", "Dayname (MM/dd/yy)" x7
function Import-ScheduleFile
{
    ## Placeholder: convert native file format or CSV into table-like PSObject
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Importing scheduling file ..."
    if ($Type -ne 'CSV'){throw "Error: NYI"} #switch $Type, transform input schedule file into template format
    return Import-CSV $CSVPath
}
$scheduleTable = Import-ScheduleFile

## Process header column names
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Processing table date columns ..."
$scheduleTable = Import-ScheduleFile
$scheduleDays = @()
[string[]]$columnNames = @(($scheduleTable[0].PSObject.Properties).Name) #Monday (05/16/22)
$columnNames | Where-Object {$_ -match "^[A-Z]{3,6}day .?[0-9]{2}/[0-9]{2}/[0-9]{2}.?$"} | ForEach-Object {
    $scheduleDays += [PSCustomObject]@{
        Name = $_.Split(' ')[0] #Monday
        Date = $_.Split(' ')[1].Trim('(',')')}} #05/16/22
if ($scheduleDays.count -ne 7){throw "Error: Less than 7 scheduled days were found in the table header."}

## Replace table day name/date header with a day names only header
## $scheduleDays now contains the actual dates for each day name, so this drops the date and enables $dayScheduleTable.DayName syntax
$dayScheduleTable = @()
[string[]]$dayColumns = $columnNames | Where-Object {$_ -match "^[A-Z]{3,6}day .?[0-9]{2}/[0-9]{2}/[0-9]{2}.?$"}
$scheduleTable | ForEach-Object {
    $dayScheduleTable += [PSCustomObject]@{
        Name = $_.Name
        EmployeeID = $_.EmployeeID
        "$($dayColumns[0].Split(' ')[0])" = $_."$($dayColumns[0])" #$_."Monday (05/16/22)"
        "$($dayColumns[1].Split(' ')[0])" = $_."$($dayColumns[1])" #$dayColumns array used due to differing week start days
        "$($dayColumns[2].Split(' ')[0])" = $_."$($dayColumns[2])"
        "$($dayColumns[3].Split(' ')[0])" = $_."$($dayColumns[3])"
        "$($dayColumns[4].Split(' ')[0])" = $_."$($dayColumns[4])"
        "$($dayColumns[5].Split(' ')[0])" = $_."$($dayColumns[5])"
        "$($dayColumns[6].Split(' ')[0])" = $_."$($dayColumns[6])"
    }
}

## Process employees into scheduling table PSCustomObject array
## Object properties: Name, EmployeeID, Schedule ([PSCustomObject[]]), VCalendar ([PSCustomObject[]])
$employeeSchedulingTable = @()
$ICSDateFormat = 'yyyyMMddTHHmmssZ'
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Generating @($dayScheduleTable.count) employee scheduling objects ..."
foreach ($scheduleRow in $dayScheduleTable)
{
    [string[]]$ICS = @()
    $ICS += 'BEGIN:VCALENDAR'
    $ICS += 'VERSION:2.0'
    $ICS += 'PRODID:-//PowerShell//github.mbarr564//EN'
    $employeeScheduledDays = @()
    foreach ($scheduleDay in $scheduleDays)
    {
        ## Read cell contents
        [string]$cellTimeSpan = $scheduleRow."$($scheduleDay.Name)"
        if ($cellTimeSpan = 'OFF')
        {
            ## Day OFF today
            $employeeScheduledDays += [PSCustomObject]@{
                Name = $scheduleDay.Name
                Date = $scheduleDay.Date
                Working = $false
            }
        }
        elseif ($cellTimeSpan -match "^[0-9]{2}:[0-9]{2}.?-.?[0-9]{2}:[0-9]{2}$")
        {
            ## Parse scheduled shift cell
            ## Split cell data on hyphen, to get start/end times, then on colon, to get hours/minutes. e.g. 10:30-16:00 below
            $startHour = $cellTimeSpan.Split('-')[0].Trim(' ').Split(':')[0] #[0]:[0] = [Left side of hyphen split]:[Left side of colon split] = 10
            $startMinute = $cellTimeSpan.Split('-')[0].Trim(' ').Split(':')[1] #[0]:[1] = [Left side of hyphen split]:[Right side of colon split] = 30
            $endHour = $cellTimeSpan.Split('-')[1].Trim(' ').Split(':')[0] #[1]:[0] = [Right side of hyphen split]:[Left side of colon split] = 16
            $endMinute = $cellTimeSpan.Split('-')[1].Trim(' ').Split(':')[1] #[1]:[1] = [Right side of hyphen split]:[Right side of colon split] = 00
    
            ## Create datetime objects
            [datetime]$startDateTime = (Get-Date -Date ($scheduleDay.Date) -Hour $startHour -Minute $startMinute)
            if ($endHour -lt $startHour){[datetime]$endDateTime = (Get-Date -Date ($scheduleDay.Date) -Hour $endHour -Minute $endMinute).AddDays(1)} #if ending hour is smaller than start (past midnight), rollover into the next day
            else {[datetime]$endDateTime = (Get-Date -Date ($scheduleDay.Date) -Hour $endHour -Minute $endMinute)}

            ## Add schedule object to employee schedule array
            $employeeScheduledDays += [PSCustomObject]@{
                Name = $scheduleDay.Name
                Date = $scheduleDay.Date
                Working = $true
                Start = $startDateTime
                End = $endDateTime
                Total = (($endDateTime - $startDateTime).TotalHours).ToString("##.#"))
            }

            ## Add schedule event to iCalendar ICS file contents
            $ICS += 'BEGIN:VEVENT'
            $ICS += "UID:$([guid]::NewGuid())"
            $ICS += "CREATED:$((Get-Date).ToUniversalTime().ToString($ICSDateFormat))"
            $ICS += "DTSTAMP:$((Get-Date).ToUniversalTime().ToString($ICSDateFormat))"
            $ICS += "LAST-MODIFIED:$((Get-Date).ToUniversalTime().ToString($ICSDateFormat))"
            $ICS += 'CLASS:PRIVATE'
            $ICS += 'SEQUENCE:0'
            $ICS += "DTSTART:$($startDateTime.ToUniversalTime().ToString($ICSDateFormat))"
            $ICS += "DTEND:$($endDateTime.ToUniversalTime().ToString($ICSDateFormat))"
            $ICS += 'SUMMARY:Work Shift'
            $ICS += 'DESCRIPTION:Work Shift'
            $ICS += 'STATUS:CONFIRMED'
            $ICS += 'TRANSP:OPAQUE'
            $ICS += 'BEGIN:VALARM'
            $ICS += 'ACTION:DISPLAY'
            $ICS += 'DESCRIPTION:Work Shift Reminder'
            $ICS += 'TRIGGER:-P1D'
            $ICS += 'END:VALARM'
            $ICS += 'END:VEVENT'
        }
        else {$adminErrors += "Attention required: cell time span '$cellTimeSpan' for '$($scheduleRow.Name)' ($($scheduleRow.EmployeeID)) on $($scheduleDay.Date) is invalid."}
    }
    $ICS += 'END:VCALENDAR'

    ## Add employee object with above scheduling data properties
    Write-Verbose "[$(Get-Date -f HH:mm:ss.fff)] Adding employee '$($scheduleRow.Name)' ($($scheduleRow.EmployeeID)) to scheduling table ..."
    $employeeSchedulingTable += [PSCustomObject]@{
        Name = $scheduleRow.Name
        EmployeeID = $scheduleRow.EmployeeID
        Schedule = $employeeScheduledDays
        VCalendar = $ICS
    }
}

## Scheduling table to email body / attachment
#$missingEmailNamesAndIDs += "'$($scheduleRow.Name)' ($($scheduleRow.EmployeeID))"
#$adminErrors += "Attention required: missing email addresses for: $($missingEmailNamesAndIDs -join ', '))."
#$employeeSchedulingTable.Schedule | ForEach-Object {$mailBodyHTML += "<ul>$($_.Name): $($_.Start) - $($_.End)</ul>"; $weekTotalHours += $_.Total}
#$employeeSchedulingTable.VCalendar | Out-File -FilePath $fileName -Encoding UTF8 -Force

## Send secure HTML email with ICS attachment
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Sending $($employeeSchedulingTable.count) emails to employees ..."
## ProtonMail w/Bridge, $48/yr:
#https://protonmail.com/bridge
#https://account.protonmail.com/signup?plan=plus&billing=12&currency=USD&language=en
#https://stackoverflow.com/questions/61574549/sending-emails-with-protonmail-in-c-sharp
## import/generate bridge config file
## $mailArguments = @{
#    Credential = $creds
#    To = 'asd'
#    From = $senderAddress
#    Subject = 'asd'
#    ReplyTo = $replyToAddress
#    DeliveryNotificationOption = OnFailure
#    SMTPServer = 'asd'
#    Attachment = 'asd'
#    UseSSL = $true
#}
## Send-MailMessage @mailArguments