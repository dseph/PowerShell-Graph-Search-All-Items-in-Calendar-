# PowerShell-Graph-SearchAllItemsInCalendar.ps1
# PowerShell - Graph Calendar query with logging sample. 
# It will pull data for all calendar items based-upon filtering.
# Note, this is not setup to filter by date.  It does use paging and will log results
# Note: Initially generated with Copilot then modified heavily.

# =========================
# 1) CONFIG (fill these in)
# =========================
# You can specify the credentials directly in code here and comment out the injection line below for 
# PowerShell-Graph-SearchAllItemsInCalendar_Credentials.ps1 or uncomment and set the three values below.
# This allows you to directly set the values here OR use the crentials from the seperate file. 
# Sometimes it helps to not show others your credentials while showing others your main code.
#$TenantId     = "YOUR_TENANT_ID"                # TODO - set
#$ClientId     = "YOUR_CLIENT_ID"                # TODO - set
#$ClientSecret = "YOUR_CLIENT_SECRET"            # TODO - set
. "$PSScriptRoot\PowerShell-Graph-SearchAllItemsInCalendar_Credentials.ps1" 

$UserPrincipalName    = "someone@contoso.com"   # TODO - set
$CalendarId = "AAMkAGEyMTcxMWI4LWEyYzAtNGI1.."  # TODO - set to the Calendar's Graph ID
$TopValue = "250"                               # TODO - set result paging size
 
# Optional "search" keyword (local filter). Leave empty to keep everything in the range.
# Note: that the code block which uses $Keyword is commented in section 4 so that it does not process - uncomment it to use.
#$Keyword = ""
#$Keyword = "Project"

# Output paths
$OutCsv  = "C:\Temp\CalendarViewResults.csv"    # TODO - set
$OutJson = "C:\Temp\CalendarViewResults.json"   # TODO - set
$OutLog  = "C:\Temp\CalendarViewRun.log"        # TODO - set

# =============================================
# 2) Get app-only access token (client creds)
# =============================================
# Token endpoint + client_credentials pattern [2](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)[3](https://learn.microsoft.com/en-us/graph/auth-v2-service)[5](https://outlook.office365.com/owa/?ItemID=AAMkADE3ZDEyNzIyLWNmYTEtNDJjNC1iMDcxLWQ1YzRlOTllNThmZgBGAAAAAAAPhJtWhlq%2bR7kfwUCSYEOdBwAR7ZLzvL8SS65OYqkrzuiyAAAAjZ%2fzAADRScJG1g3ITIOZvbEKNMlnAAbp2uqJAAA%3d&exvsurl=1&viewmodel=ReadMessageItem)
$tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$tokenBody = @{
  client_id     = $ClientId
  client_secret = $ClientSecret
  scope         = "https://graph.microsoft.com/.default"
  grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenUri -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
$accessToken   = $tokenResponse.access_token

$headers = @{
  Authorization = "Bearer $accessToken"
  "Accept"      = "application/json"
}

# =========================
# 3) Query 
# =========================

$select ="id,subject,iCalUId,isOrganizer,isOnlineMeeting,onlineMeetingProvider,originalEndTimeZone,recurrence,start,end,subject,onlineMeeting" 
$filter = "isOrganizer eq true and isCancelled eq false"
$uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/calendars/$CalendarId/events?" +
       "&`$select=$select&`$top=$TopValue&`$filter=$filter"
$uri 
 
$ItemCount = 0

 
$all = New-Object System.Collections.Generic.List[object]

while ($uri) {
  try {
    $resp = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$select ="id,subject,iCalUId,isOrganizer,isOnlineMeeting,onlineMeetingProvider,originalEndTimeZone,recurrence,start,end,subject,onlineMeeting" 
    foreach ($ev in $resp.value) {
      # Flatten key fields for CSV
      $all.Add(
            [pscustomobject]@{
            id                      = $ev.id
            subject                  = $ev.subject
            start                   = $ev.start
            end                     = $ev.end
            iCalUid                 = $ev.iCalUId
            isOrganizer             = $ev.isOrganizer
            isOnlineMeeting         = $ev.isOnlineMeeting
            onlineMeetingProvider   = $ev.onlineMeetingProvider
            originalEndTimeZone     = $ev.originalEndTimeZone
            recurrence              = $ev.recurrence
            onlineMeeting           = $ev.onlineMeeting
            }   
      )
        $ItemCount += 1
    }

    # Paging
    $uri = $resp.'@odata.nextLink'
    "[$(Get-Date -Format o)] Retrieved $($resp.value.Count) events; nextLink present: $([bool]$uri)" | Out-File -FilePath $OutLog -Append -Encoding utf8
  }
  catch {
    "[$(Get-Date -Format o)] ERROR calling Graph: $($_.Exception.Message)" | Out-File -FilePath $OutLog -Append -Encoding utf8
    throw
  }
}

$results = $all

Write-Host "Total Items Found: $ItemCount"
# =========================
# 4) Optional local "search"
# =========================
# Note: Uncomment the section below if you want to filter teh results.  Also, If $Keyword is not set then this filtering below will not be done.
<#  
if (-not [string]::IsNullOrWhiteSpace($Keyword)) {
  $results = $all | Where-Object { $_.subject -like "*$Keyword*" }
  "[$(Get-Date -Format o)] Local filter applied: subject contains '$Keyword' -> $($results.Count) matches" | Out-File -FilePath $OutLog -Append -Encoding utf8
}
#>
 
# =========================
# 5) Export / Log outputs
# =========================
$results | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
$results | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutJson -Encoding utf8

"[$(Get-Date -Format o)] Exported $($results.Count) events to:`r`n CSV: $OutCsv`r`n JSON: $OutJson" |
  Out-File -FilePath $OutLog -Append -Encoding utf8

Write-Host "Done. Events exported to:"
Write-Host "  $OutCsv"
Write-Host "  $OutJson"
Write-Host "Run log:"
Write-Host "  $OutLog"
