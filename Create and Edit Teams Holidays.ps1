#This script will create or edit Microsoft Teams holiday schedules

#Prerequisites:
    #Active connection to Skype for Business online PowerShell.

#Author: Clayton Martin, Landis Technologies LLC
#This script is provided "As Is" without any warranty of any kind. In no event shall the author be liable for any damages arising from the use of this script.

$Country = "usa"
$region = "pa"
$holidayType = "public_holiday"
$HolidayYear = "2021"
#Time must be in 15min increments
$StartTime = "00:00"
$EndTime = "23:45"

<#Parameter Options: https://kayaposoft.com/enrico/json/

country:	
    ISO 3166-1 alpha-3 country code or ISO 3166-1 alpha-2 country code
region:	
    Possible values for New Zealand: ISO 3166-2:NZ codes - auk, bop, can, gis, hkb, mbh, mwt, nsn, ntl, ota, stl, tas, tki, wko, wgn, wtc, cit
    Possible values for Australia: ISO 3166-2:AU codes - nsw, qld, sa, tas, vic, wa, act, nt
    Possible values for Canada: ISO 3166-2:CA codes - ab, bc, mb, nb, nl, ns, on, pe, qc, sk, nt, nu, yt
    Possible values for United States of America: ISO 3166-2:US codes - al, ak, az, ar, ca, co, ct, de, fl, ga, hi, id, il, in, ia, ks, ky, la, me, md, ma, mi, mn, ms, mo, mt, ne, nv, nh, nj, nm, ny, nc, nd, oh, ok, or, pa, ri, sc, sd, tn, tx, ut, vt, va, wa, wv, wi, wy, dc
    Possible values for Germany: ISO 3166-2:DE codes - bw, by, be, bb, hb, hh, he, mv, ni, nw, rp, sl, sn, st, sh, th
    Possible values for Great Britain: ISO 3166-2:GB codes - eng, nir, sct, wls
holidayType:
    all - all holiday types
    public_holiday - public holidays
    observance - observances, not a public holidays
    school_holiday - school holidays
    other_day - other important days e.g. Mother's day, Father's day etc
    extra_working_day - extra working days. This day takes place mostly on Saturday or Sunday and is substituted for extra public holiday.
#>


#Get Existing Holiday Schedules
$ExistingHolidays = Get-CsOnlineSchedule

#Get New Holidays
$uri = "https://kayaposoft.com/enrico/json/v2.0?action=getHolidaysForYear&year=" + $HolidayYear + "&country=" + $Country + "&region=" + $region + "&holidayType=" + $holidayType
$Holidays = (Invoke-WebRequest -Method Get -Uri $uri).Content | ConvertFrom-Json
[System.Collections.ArrayList]$formattedHolidays = @()
foreach ($Holiday in $Holidays){
    $name = $Holiday.name.text -replace ",", ""
    $date = $holiday.date.day.ToString() + "/" + $holiday.date.month.ToString() + "/" + $holiday.date.year.ToString()
    $startdate = $date + " " + $StartTime
    $enddate = $date + " " + $EndTime
    $PCO = [PSCustomObject]@{
        Name = $name
        StartDateTime = $startdate
        EndDateTime = $enddate
    }
    $formattedHolidays.Add($pco) | Out-Null
}

#Check If Holiday Exists
[System.Collections.ArrayList]$HolidaysToEdit = @()
foreach ($formattedHoliday in $formattedHolidays){
    foreach ($ExistingHoliday in $ExistingHolidays){
        if ($ExistingHoliday.Name -eq $formattedHoliday.Name){
            $HolidaysToEdit.Add($formattedHoliday) | Out-Null
        }
    }
}
#Remove Existing Holidays From Being Created
for ($i = 0; $i -lt $HolidaysToEdit.Count; $i++){
    $formattedHolidays.Remove($HolidaysToEdit[$i])
}

#Create New Holidays
foreach ($formattedHoliday in $formattedHolidays){
    $name = $formattedHoliday.Name
    Write-Host "Would you like to add the $name Holiday?"
    $ReadHost = Read-Host "( y / n )"
    if ($ReadHost -eq "y"){
        $dtr = New-CsOnlineDateTimeRange -Start $formattedHoliday.StartDateTime -End $formattedHoliday.EndDateTime
        New-CsOnlineSchedule -Name $formattedHoliday.Name -FixedSchedule -DateTimeRanges @($dtr)
    }
}

#Edit Existing Holidays
foreach ($HolidayToEdit in $HolidaysToEdit){
    $name = $HolidayToEdit.Name
    Write-Host "Holiday $name already exists. Would you like to add a new Date and Time Schedule?"
    $ReadHost = Read-Host "( y / n )"
    if ($ReadHost -eq "y"){
        $DateTime = New-CsOnlineDateTimeRange -Start $HolidayToEdit.StartDateTime -End $HolidayToEdit.EndDateTime
        $schedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayToEdit.Name}
        if ($schedule.FixedSchedule.DateTimeRanges.Count -le 10){
            $schedule.FixedSchedule.DateTimeRanges += $DateTime
            Try{
                Set-CsOnlineSchedule -Instance $schedule -ErrorAction stop
            }
            Catch{
                $errorMessage = $_.exception.message
                Write-Host "Holiday $name`: $errorMessage Skipping....." -ForegroundColor Red

            }
        }
        if ($schedule.FixedSchedule.DateTimeRanges.Count -gt 10){
            Write-Host "Cannot add date and time range because there is a maximum of 10 already specified for $name. Please delete some schedules for this holiday and run this script again."
        }
    }
}
