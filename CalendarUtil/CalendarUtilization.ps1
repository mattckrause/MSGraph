<#
Getting room utilization based on calendar appointments. Calculates based on 5day work week.
Script built using Microsoft Graph PowerShell SDK - https://docs.microsoft.com/en-us/graph/powershell/installation
Permission required: Calendars.Read - https://docs.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0&tabs=http
#>

#variables
$reportLocation = "C:\temp\CalendarUtil.csv"
$days = 17 #<- duration for util calculation
$endTime = Get-Date
$weekend = 0
$calendars = @('cal1@domain.com','cal2@domain.com','cal3@domain.com') # <- List of rooms to look up. Could be changed to use data from cmdlet or .csv
[System.Collections.ArrayList]$reportData = @() #<-Setting up the report array list

#Connect to Graph
$appID = "" #<- App ID from application registration in Azure AD
$tenantID = "" #<- Tenant ID from application registration in Azure AD
$certThumb = "" #<- Certificate thumbprint from application registration in Azure AD

Connect-MgGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $certThumb

#calculate weekdays and update $startTime for accurate timeframe
0..$days | foreach-object{
    $b = ($endTime).AddDays(-$_);
    if($b.DayOfWeek -eq "Saturday" -OR $b.DayOfWeek -eq "Sunday")
    {
        $weekend++
    }
}

$startTime = ($endTime).AddDays(-($days+$weekend))
foreach ($calendar in $calendars)
{
    #MS Grapgh command using PowerShell SDK to look up all events on a calendar
    $calendarEvents = Get-MgUserCalendarView -UserId $calendar -StartDateTime $startTime -EndDateTime $endTime
    $calendarEventCount = 0
    $calendarEventDuration = New-TimeSpan
    #Calculate time and number of events per calendar
    foreach ($appointment in $calendarEvents)
    {
        $totalduration = New-TimeSpan
        if($appointment.IsAllDay)
        {
            #if all day use 8 hours
            $totalduration = New-TimeSpan -Hours 8
        }
        else
        {
            #not all day, calculate duration of meeting/appointment
            [system.DateTime]$start = $appointment.Start.DateTime
            [System.DateTime]$end = $appointment.End.DateTime
            $totalDuration = (New-TimeSpan).Add($end.Subtract($start))
        }
        #keep track of number of events
        $calendarEventCount ++
        #add time to total duration for util calculation
        $calendarEventDuration += $totalDuration
    }
    #Calc Utilization based on 8 hour work days
    $calendarUtilization = (($calendarEventDuration.TotalHours) / $([int]$days * 8)).ToString("P")
    
    #record data and add to report
    $roomEvent = [PSCustomObject]@{
        'Calendar'          = $calendar;
        'ReportPeriod'      = $days;
        'TotalAppointments' = $calendarEventCount;
        'TotalHoursBooked'  = $calendarEventDuration;
        'Utilization'       = $calendarUtilization
    }
    $reportData.add($roomEvent) | Out-Null
    $roomEvent=$null
}

#export report data to .csv
$reportData | Export-Csv -Path $reportLocation -NoTypeInformation

Disconnect-MgGraph #<-- Remember to disconnect from MS Graph API when the app is done!
