<#
.SYNOPSIS
    Intended for use by restaurants that already use weekly scheduling spreadsheets.
    Uses the iCalendar RFC/standard to generate widely compatible calendar events as an ICS file attachment: https://en.wikipedia.org/wiki/ICalendar
.DESCRIPTION
    Requires a ProtonMail paid account ($48/year) to function as a turnkey solution:
    1. Subscribe: https://account.protonmail.com/signup?plan=plus&billing=12&currency=USD&language=en
    2. Install: https://protonmail.com/bridge
    3. Config: ProtonMail Bridge > Settings (Gear Icon) > Advancecd Settings > SMTP connection mode > SSL
.PARAMETER Path
    Path to the folder where all new scheduling files are to be saved/exported.
    The script will process ONLY ONE FILE from this folder location, the NEWEST based on the last modified date.
.PARAMETER Type
    Specifies the type of file being used, which changes how the file is processed: dropping extra columns, formatting cell values, etc.
    If you would like your restaurant(s) to be added to these $Types, supply a short 2 to 4 character name, and a sample scheduling file, as it's saved from your spreadsheet application: and email those to me: mbarr564@protonmail.com
    The scheduling file should have employee names changed, and cell values modified/deleted, as long as the file's exact unmodified formatting remains intact (i.e. don't remove any commas, or columns, and don't remove the top row header).
.PARAMETER Setup
    NYI: Run/rerun first time setup winforms wizard, as seen after double-clicking the batch launcher script.
.EXAMPLE
    PS> .\Send-WorkScheduleWithICS.ps1 -Path "$PSScriptRoot\schedule_template.csv" -Type 'CSV'
.NOTES
    Last update: Tuesday, April 5, 2022 1:15:40 AM
#>

param ([string]$Path = "$PSScriptRoot\schedule_template.csv", [ValidateSet('CSV','McD','BK','Wend','DQ','TacT','TacB','SubW')][string]$Type = 'CSV', [switch]$Setup)

## Init
$VerbosePreference = 'Continue'
[string[]]$adminErrors = @()
[string[]]$configPatterns = @("^127\.0\.0\.1$","^[0-9]{2,5}$","^.*@.*\..*$","^[A-Z0-9-_]{20,32}$","^SSL$") #regex for each config line
[string[]]$protonMailBridgeConfig = @(Get-Content "$PSScriptRoot\..\..\local\Send-WorkScheduleWithICS\protonmail.ini") #todo: winforms: generate protonmail.ini if missing
if (-not($protonMailBridgeConfig)){throw "Error: unable to import ProtonMail Bridge config file from: $PSScriptRoot\protonmail.ini"} #todo: replace bridge password with secure/encrypted string
for ($i = 0; $i -le 4; $i++){if ($protonMailBridgeConfig[$i] -notmatch $configPatterns[$i]){throw "Error: ProtonMail Bridge config line $($i + 1) value '$($protonMailBridgeConfig[$i])' is invalid."}}
$emailTable = Import-CSV "$PSScriptRoot\..\..\local\Send-WorkScheduleWithICS\emails.csv" #todo: winforms: generate emails.csv and add missing addresses
$emailTable = $emailTable | Where-Object {$_.EmailAddress -like "*@*.*"} #filter missing email addresses

## Import schedule file
## Returned columns: "Name", "EmployeeID", "Dayname (MM/dd/yy)" x7
function Import-ScheduleFile
{
    ## Placeholder: convert native file format or CSV into table-like PSObject
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Importing scheduling file ..."
    if ($Type -ne 'CSV'){throw "Error: NYI"} #switch $Type, transform input schedule file into template format: a weekly generic 24hr time span table
    return Import-CSV $Path
}
$scheduleTable = Import-ScheduleFile

## Process header column names
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Processing table date columns ..."
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
        "$($dayColumns[0].Split(' ')[0])" = $_."$($dayColumns[0])" #"Monday" = $_."Monday (05/16/22)"
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
        if ($cellTimeSpan -eq 'OFF')
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
            $startHour = $cellTimeSpan.Split('-')[0].Trim(' ').Split(':')[0]    #[0]:[0] = [Left side of hyphen split]:[Left side of colon split] = 10
            $startMinute = $cellTimeSpan.Split('-')[0].Trim(' ').Split(':')[1] #[0]:[1] = [Left side of hyphen split]:[Right side of colon split] = 30
            $endHour = $cellTimeSpan.Split('-')[1].Trim(' ').Split(':')[0]    #[1]:[0] = [Right side of hyphen split]:[Left side of colon split] = 16
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
                Timespan = $cellTimeSpan
                Start = $startDateTime
                End = $endDateTime
                Total = (($endDateTime - $startDateTime).TotalHours).ToString("##.##")
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
        else {$adminErrors += "Cell time span '$cellTimeSpan' for '$($scheduleRow.Name)' ($($scheduleRow.EmployeeID)) on $($scheduleDay.Date) is invalid."}
    }
    $ICS += 'END:VCALENDAR'

    ## Employee email address
    $employeeEmail = ($emailTable | Where-Object {$_.EmployeeID -eq $employeeRow.EmployeeID}).EmailAddress
    if (-not($employeeEmail)){$employeeEmail = 'Unknown'}

    ## Add employee object with above scheduling data properties
    Write-Verbose "[$(Get-Date -f HH:mm:ss.fff)] Adding employee '$($scheduleRow.Name)' ($($scheduleRow.EmployeeID)) to scheduling table ..."
    $employeeSchedulingTable += [PSCustomObject]@{
        Name = $scheduleRow.Name
        EmployeeID = $scheduleRow.EmployeeID
        Schedule = $employeeScheduledDays
        VCalendar = $ICS
        EmailAddress = $employeeEmail
    }
}

## Build and send secure HTML email with ICS attachment, to all employees in the scheduling table
## $protonMailBridgeConfig[0] IP address, [1] service port, [2] email address, [3] Bridge password
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Sending $($employeeSchedulingTable.count) emails to employees ..."
$password = $protonMailBridgeConfig[3] | ConvertTo-SecureString -AsPlainText -Force #todo: replace bridge password with secure/encrypted string
$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList ($protonMailBridgeConfig[2],$password)
foreach ($employee in $employeeSchedulingTable)
{
    ## Notify admin and skip missing email addresses
    if ($employee.EmailAddress -eq 'Unknown'){$adminErrors += "Email address missing for '$($scheduleRow.Name)', please add EmployeeID $($scheduleRow.EmployeeID)."; continue}
    
    ## Employee email message body
    [string]$totalHours = ($employee.Schedule | Measure-Object -Sum -Property Total).Sum
    [string[]]$msgBody = @()
    $msgBody += "<html><body>Hello $($employee.Name),<br>"
    $msgBody += "Your work schedule for week starting $($scheduleDays[0].Date):<br><ul>"
    $employee.Schedule | ForEach-Object {$msgBody += "<li>$($_.Name) $($_.Date): $($_.Timespan)</li>"} # - Monday 05/16/22: 10:30-16:00
    $msgBody += "</ul><br>You're scheduled for $totalHours total hours.<br>Calendar app work schedule is attached.</body></html>"

    ## Employee email arguments
    $mailArguments = @{
        Credential = $credentials
        To = $employee.EmailAddress
        From = $protonMailBridgeConfig[2]
        Subject = "$($employee.Name) Work Schedule: Week Starting $($scheduleDays[0].Date)"
        Body = $msgBody
        DeliveryNotificationOption = OnFailure
        SMTPServer = $protonMailBridgeConfig[0]
        Port = $protonMailBridgeConfig[1]
        UseSSL = $true
        ErrorAction = Stop
    }

    ## Employee email attachment
    $ICSFilePath = "$($env:TEMP)\calendar.ics"
    if (Test-Path $ICSFilePath){Remove-Item -LiteralPath $ICSFilePath}
    $employee.VCalendar | Out-File -FilePath $ICSFilePath -Encoding UTF8
    if (Test-Path $ICSFilePath){$mailArguments.add('Attachment',$ICSFilePath)}

    ## Email the employee schedule
    try {Send-MailMessage @mailArguments}
    catch {$adminErrors += "Error sending email to '$($employee.EmailAddress)' ($($employee.Name) ($($employee.EmployeeID)))."}
}

## Admin email message body
[string[]]$adminMsgBody = @()
$adminMsgBody += '<html><body>The following issues need your attention:<br><ul>'
$adminErrors | ForEach-Object {$adminMsgBody += "<li>$_</li>"}
$adminMsgBody += '</ul></body></html>'

## Admin email arguments
$adminMailArguments = @{
    Credential = $credentials
    To = $protonMailBridgeConfig[2]
    From = $protonMailBridgeConfig[2]
    Subject = "Attention required: Send-WorkScheduleWithICS: Week Starting $($scheduleDays[0].Date)"
    Body = $adminMsgBody
    DeliveryNotificationOption = OnFailure
    SMTPServer = $protonMailBridgeConfig[0]
    Port = $protonMailBridgeConfig[1]
    UseSSL = $true
}

## Admin email
Send-MailMessage @adminMailArguments

# SIG # Begin signature block
# MIIVpAYJKoZIhvcNAQcCoIIVlTCCFZECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQULbTi+Jjf6kP2HyYrdRm5CMux
# 64mgghIFMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5n
# IFJvb3QgUjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIE
# JHQu/xYjApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7
# fbu2ir29BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGr
# YbNzszwLDO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTH
# qi0Eq8Nq6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv
# 64IplXCN/7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2J
# mRCxrds+LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0P
# OM1nqFOI+rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXy
# bGWfv1VbHJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyhe
# Be6QTHrnxvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXyc
# uu7D1fkKdvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7id
# FT/+IAx1yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQY
# MBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJw
# IDaRXBeF5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUE
# DDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmlj
# YXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3Sa
# mES4aUa1qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+
# BtlcY2fUQBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8
# ZsBRNraJAlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx
# 2jLsFeSmTD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyo
# XZ3JHFuu2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p
# 1FiAhORFe1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG
# 9w0BAQwFADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS0wKwYDVQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIB
# igKCAYEAmyudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxD
# eEDIArCS2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk
# 9vT0k2oWJMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7Xw
# iunD7mBxNtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ
# 0arWZVeffvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZX
# nYvZQgWx/SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+t
# AfiWu01TPhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvr
# n35XGf2RPaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn
# 3UayWW9bAgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaR
# XBeF5jAdBgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYD
# VR0gBBQwEjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RS
# NDYuY3JsMHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEM
# BQADggIBAAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXK
# ZDk8+Y1LoNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWk
# vfPkKaAQsiqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3d
# MapandPfYgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwF
# kvjFV3jS49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZa
# PATHvNIzt+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8b
# kinLrYrKpii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7Ew
# oIJB0kak6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TW
# SenLbjBQUGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg
# 51Tbnio1lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoU
# KD85gnJ+t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGcDCCBNig
# AwIBAgIQVdb9/JNHgs7cKqzSE6hUMDANBgkqhkiG9w0BAQwFADBUMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgQ0EgUjM2MB4XDTIxMTIxNjAwMDAwMFoXDTIy
# MTIxNjIzNTk1OVowUDELMAkGA1UEBhMCVVMxEzARBgNVBAgMCldhc2hpbmd0b24x
# FTATBgNVBAoMDE1pY2hhZWwgQmFycjEVMBMGA1UEAwwMTWljaGFlbCBCYXJyMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAne6XW99iRvph0mHzkgX+e+6i
# mXxytFu35Vw4YC0TSeDqkUCc0PoSyojLc+MKLa/t+32ya1BWmSf1u5Hc55yo9BL3
# dvV7C9HisQ8gB3+Cb+04P+0b/buBor9M7Cu+rJe7RZOVS9bq+CuslCchBejc6tNe
# f+A8b1q9jzjgVvAUpv+dD4asi/KhMYdhDWxI23i0A9XOn8OBrfsu9zQBYGxFX7Is
# Wk+wunMNwN6PPeZ9gFVwHuh5OVXEDIXGVm+N7QTSdTTdLC6w5ttWzVrsKdQM6vZI
# yNuV5x1bQ32cbBdT2oB+R7ODSmuMTxMagfm4lrqjPZKNP91MCRVpbWbv/4/ealte
# KResVeIm+mQbXkWmFWIHgLkXToVDlyWOBFjG0I5rt2p9055FZ7Xpo36Vinvs+JWj
# fgDaYKPEeHJ3AFwdJD6gjVBH9xt0IJlZm7rWiqE+BpsgzxBKJGYzHqBwmWtLFZvG
# 5DdwVKCThFGyoIawT/POm7eBU9tyePv1g95xkzesqHGz854f+w+XXWW/qwAZBMAY
# QnAPLFI1ywJ1GHVkp7xZRaxAOEiId0WG57R/y4h5gtE12nPa07PUrtl3HPClZICE
# 6PP5UimZH2fF2ClwyAoaxXU70yblD6V+gzZ1wgDpDl1FYyDdZmtjtz6zh8MAp9b+
# /Rk2BS3SWH9iUjn0yTECAwEAAaOCAcAwggG8MB8GA1UdIwQYMBaAFA8qyyCHKLjs
# b0iuK1SmKaoXpM0MMB0GA1UdDgQWBBSmkRZEx8ANTjiACrZOUmUhtz5SvTAOBgNV
# HQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzAR
# BglghkgBhvhCAQEEBAMCBBAwSgYDVR0gBEMwQTA1BgwrBgEEAbIxAQIBAwIwJTAj
# BggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwCAYGZ4EMAQQBMEkG
# A1UdHwRCMEAwPqA8oDqGOGh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1
# YmxpY0NvZGVTaWduaW5nQ0FSMzYuY3JsMHkGCCsGAQUFBwEBBG0wazBEBggrBgEF
# BQcwAoY4aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNp
# Z25pbmdDQVIzNi5jcnQwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28u
# Y29tMCIGA1UdEQQbMBmBF21iYXJyNTY0QHByb3Rvbm1haWwuY29tMA0GCSqGSIb3
# DQEBDAUAA4IBgQBNf+rZW+Q5SiL5IjqhOgozqGFuotQ+kpZZz2GGBJoKldd1NT7N
# kEJWWzCd8ezH+iKAd78CpRXlngi/j006WbYuSZetz4Z2bBje1ex7ZcL00Hnh4pSv
# heAkiRgkS5cvbtVVDnK5+AdEre/L71qU1gpZGNs7eqEdp5tBiEcf9d18H0hLHMtd
# 5veYH2zXqwfXo8SNGYRz7CCgDiYSdHDsSE284a/CcUivte/jJe1YmZR/Zueuisti
# fkeqldgFrqc30JztyIU+EVXeNOivA5yihYj5WBz7zMVjnBsmEH0bUdrKImptWzCw
# 2x8dGzziG7jfeYs20gG05Xv4Jd0IBdoxhRMeznT8WhvwifG9aN4IZPDMyfYT9v1j
# 2zx8EbcmhD1aaio9gP18AvBWksa3KvOChA1BQvD7PR5YucZEzoljq10kIjKsLA3U
# te7JSxpXDFC7Ab/xeUYRGIG/x/wyCLRjENe+ryixRy6txVUDkxqDsqngzPVeyvYM
# fjlXjk9R0ZjWwNsxggMJMIIDBQIBATBoMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQK
# Ew9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUg
# U2lnbmluZyBDQSBSMzYCEFXW/fyTR4LO3Cqs0hOoVDAwCQYFKw4DAhoFAKB4MBgG
# CisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
# AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYE
# FCFFWafIoatLM2fhnH8FjzUmcDLjMA0GCSqGSIb3DQEBAQUABIICAG/SegXhG2gj
# DZ/yrtTwbHYp2aoyZzjQNZJvb7r7XUKb2kZbiIiGPEIdI1hDKuLn95b1bz/i2y6j
# 7+lduY+3eAx13esWgB6zaJ3ZpgHe/3zIKpxlDxtDdLGYmlbJcbLsBc+2kILMG8pE
# 5+WQaQvViqI12Etdu0JZZAqsf441VF5JYO+UwESndoEXpNyiIb0pyBYOkKFvROnI
# n7ZyKe1boe5JgCfgaMlIqsUGI3cl/gLUDP6MLG0ZIp8bqrpNoO77W4NBlu5OTaQg
# M/TS+PrO2hco65fzRkOV21EuMPMyZHGltKTTxCJ7zNVJpzMSPuf/ahDUj6zgUHxG
# 9qlveaGEFKIuKGszu0VCKuVH1Z/rApNbWMQ5tZfqa+O32TwmY+OfL/5okq8x+DOd
# oXtw5DJvuivAwlKQ1oue709q+njB8s8XppHBAiQd0ARU3f9EI6YguCm/MgJLx0Yj
# r9V0YwJhag7mPelUff8yd3RgDHoVPoKLtn2jm4R0X/neEtxR3JdwbZewk0Av1995
# eA5dC4FGHo7c/8d0GfClcgE4C/kByl4T69z/duMveK+1W6w1rVoudtry9OE+ZivN
# EgLfAaPTHRISOMKC30/aVCBT4+EMVJbiaCnpKkfZGyzZuUcl2ogqNt81wNx+GXtY
# gAvaGKa1bFWj9XCwbOCqQdeTwRcGhRXS
# SIG # End signature block
