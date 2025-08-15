Function Get-M365GroupProvisioningSummary {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet(7, 30, 90, 180)]
        [int]
        $ReportPeriod = 7
    )
    $ProgressPreference = 'SilentlyContinue'

    $null = Set-M365ReportPeriod -ReportPeriod $ReportPeriod

    try {

        # Get existing groups
        $liveGroups = [System.Collections.Generic.List[System.Object]]@()
        $uri = "https://graph.microsoft.com/beta/groups`?`$filter=groupTypes/any(c:c+eq+'Unified')`&`$select=mailNickname,deletedDateTime,createdDateTime"
        $result = Invoke-MgGraphRequest -Method Get -Uri $uri -ContentType 'application/json' -ErrorAction Stop -OutputType PSObject
        if ($result.value) {
            $liveGroups.AddRange($result.value)
            while ($result.'@odata.nextLink') {
                $result = Invoke-MgGraphRequest -Method Get -Uri $result.'@odata.nextLink' -ContentType 'application/json' -ErrorAction Stop -OutputType PSObject
                $liveGroups.AddRange($result.value)
            }
        }

        # Get deleted groups
        $deletedGroups = [System.Collections.Generic.List[System.Object]]@()
        $uri = "https://graph.microsoft.com/beta/directory/deletedItems/microsoft.graph.group`?`$filter=groupTypes/any(c:c+eq+'Unified')`&`$select=mailNickname,deletedDateTime,createdDateTime"
        $result = Invoke-MgGraphRequest -Method Get -Uri $uri -ContentType 'application/json' -ErrorAction Stop -OutputType PSObject
        if ($result.value) {
            $deletedGroups.AddRange($result.value)
            while ($result.'@odata.nextLink') {
                $result = Invoke-MgGraphRequest -Method Get -Uri $result.'@odata.nextLink' -ContentType 'application/json' -ErrorAction Stop -OutputType PSObject
                $deletedGroups.AddRange($result.value)
            }
        }

        [PSCustomObject]@{
            'Current'       = $liveGroups.count
            'Created'       = ($liveGroups | Where-Object { ([datetime]$_.createdDateTime) -ge $Script:GraphStartDate }).Count
            'Deleted'       = ($deletedGroups | Where-Object { ([datetime]$_.deletedDateTime) -ge $Script:GraphStartDate }).Count
            'Report Period' = $ReportPeriod
            'Start Date'    = ($Script:GraphStartDate).ToString('yyyy-MM-dd')
            'End Date'      = ($Script:GraphEndDate).ToString('yyyy-MM-dd')
        }
    }
    catch {
        SayError $_.Exception.Message
        return $null
    }
}