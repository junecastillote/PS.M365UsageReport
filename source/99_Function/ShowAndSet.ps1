Function ShowThisModule {
    $Script:thisModule = $MyInvocation.MyCommand.Module
    $Script:thisModule
}

Function Get-M365ReportPeriod {
    [PSCustomObject]@{
        StartDate = $Script:GraphStartDate
        EndDate   = $Script:GraphEndDate
    }
}

Function Set-M365ReportPeriod {
    param (
        [Parameter()]
        [ValidateSet(7, 30, 90, 180)]
        [int]
        $ReportPeriod = 7,

        [Parameter()]
        [switch]
        $Force
    )
    $ProgressPreference = 'SilentlyContinue'
    if ($Script:GraphStartDate -eq '1970-01-01' -or $Force) {
        $temp = (Get-M365ActiveUserCount -ReportPeriod $ReportPeriod)
        $Script:GraphStartDate = [datetime]$temp[0].'Report Date'
        $Script:GraphEndDate = [datetime]$temp[-1].'Report Date'
    }
}