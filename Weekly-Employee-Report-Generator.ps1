$Date = Get-Date
$DayDecrement = 0

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
$TuesdayDate = $Date.AddDays($DayDecrement + 1).ToString("dd.MM.yyyy.")
$WednesdayDate = $Date.AddDays($DayDecrement + 2).ToString("dd.MM.yyyy.")
$ThursdayDate = $Date.AddDays($DayDecrement + 3).ToString("dd.MM.yyyy.")
$FridayDate = $Date.AddDays($DayDecrement + 4).ToString("dd.MM.yyyy.")
$SaturdayDate = $Date.AddDays($DayDecrement + 5).ToString("dd.MM.yyyy.")
$SundayDate = $Date.AddDays($DayDecrement + 6).ToString("dd.MM.yyyy.")

$Settings = Get-Content -Path "$PSScriptRoot\Settings.cfg" | ConvertFrom-StringData

$ReportTitle = $Settings.ReportTitle + "$MondayDate - $SundayDate"
$Documents = [environment]::getfolderpath("mydocuments")
$ReportFolderName = $Settings.ReortFolderName
$ReportFullName = "$Documents\$ReportFolderName\$ReportTitle" + "md"

if (-not (Test-Path -Path $ReportFullName)) {
    #Defining names of the days in Serbian language
    $Monday = "Ponedeljak"
    $Tuesday = "Utorak"
    $Wednesday = "Sreda"
    $Thursday = "Četvratak"
    $Friday = "Petak"
    $Saturday = "Subota"
    $Sunday = "Nedelja"

    $Employee = "> Zoran Jankov - Administrator sistema i računarskih mreža"

    $Separator = "`n`n------`n`n"
    $WorkingHours = "`n------`n`n### 07:30-08:30`n`n- `n`n### 08:30-09:30`n`n- `n`n### 09:30-10:30`n`n- `n`n### 10:30-11:30`n`n- `n### 11:30-12:30`n`n- `n`n### 12:30-13:30`n`n- `n`n### 13:30-14:30`n`n- `n`n### 14:30-15:30`n`n- "
    
    $ReportBody = "# $ReportTitle`n$Employee"
    $ReportBody += "$Separator## $Monday $MondayDate`n$WorkingHours"
    $ReportBody += "$Separator## $Tuesday $TuesdayDate`n$WorkingHours"
    $ReportBody += "$Separator## $Wednesday $WednesdayDate`n$WorkingHours"
    $ReportBody += "$Separator## $Thursday $ThursdayDate`n$WorkingHours"
    $ReportBody += "$Separator## $Friday $FridayDate`n$WorkingHours"
    $ReportBody += "$Separator## $Saturday $SaturdayDate`n$WorkingHours"
    $ReportBody += "$Separator## $Sunday $SundayDate`n$WorkingHours`n`n------"

    New-Item -Path $ReportFullName -Force
    Add-Content -Path $ReportFullName -Value $ReportBody
}

Start-Process $ReportFullName