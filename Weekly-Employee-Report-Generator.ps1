$Date = Get-Date
$DayDecrement = 0
$Settings = Get-Content -Path "$PSScriptRoot\Settings.cfg" | ConvertFrom-StringData

switch ($Date.DayOfWeek) {
    'Monday' {
        $DayDecrement = 0
    }
    'Tuesday' {
        $DayDecrement = -1
    }
    'Wednesday' {
        $DayDecrement = -2
    }
    'Thursday' {
        $DayDecrement = -3
    }
    'Friday' {
        $DayDecrement = -4
    }
    'Saturday' {
        $DayDecrement = -5
    }
    'Sunday' {
        $DayDecrement = -6
    }
}
$MondayDate = $Date.AddDays($DayDecrement).ToString("dd.MM.yyyy.")
$SundayDate = $Date.AddDays($DayDecrement + 6).ToString("dd.MM.yyyy.")
$Settings = Get-Content -Path "$PSScriptRoot\Settings.cfg" | ConvertFrom-StringData

$ReportTitle = $Settings.ReportTitle + " $MondayDate - $SundayDate"
$Documents = [environment]::getfolderpath("mydocuments")
$ReportFolderName = $Settings.ReortFolderName
$ReportFullName = "$Documents\$ReportFolderName\$ReportTitle" + "md"

if (-not (Test-Path -Path $ReportFullName)) {
    $DaysOfWeek = @(
        $Settings.MondayName
        $Settings.TuesdayName
        $Settings.WednesdayName
        $Settings.ThursdayName
        $Settings.FridayName
        $Settings.SaturdayName
        $Settings.SundayName
    )

    $Employee = "> " + $Settings.EmployeeName + " - " + $Settings.EmployeeTitle
    
    $ReportBody = "# $ReportTitle`n`n$Employee`n`n------`n`n"
    foreach ($Day in $DaysOfWeek) {
        $ReportBody += "## $Day " + $Date.AddDays($DayDecrement).ToString("dd.MM.yyyy.")
        $DayDecrement ++
        $ReportBody += "`n`n------`n`n"
        for ($Hour = 0; $Hour -lt $Settings.WorkingTime; $Hour ++) {
            $Start = ([datetime]::ParseExact($Settings.StartingTime, "HH:mm", $null)).AddHours($Hour).ToString("HH:mm")
            $End = ([datetime]::ParseExact($Settings.StartingTime, "HH:mm", $null)).AddHours($Hour + 1).ToString("HH:mm")
            $ReportBody += "### $Start - $End`n`n- `n`n"
        }
        $ReportBody += "------`n`n"
    }
    
    New-Item -Path $ReportFullName -Force
    Add-Content -Path $ReportFullName -Value $ReportBody
}

Start-Process $ReportFullName