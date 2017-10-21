##Support functions for Making Auditing Great Again!

##Shrink the prompt down
function Prompt
{
    return ">";
}

##Function to convert the oAuth timestamps into real datetimes.
Function Convert-FromUnixdate ($UnixDate) {

  [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').`

  AddSeconds($UnixDate))

}

function get-oAuthAccessToken()
{
    Param(
        [Parameter(Mandatory=$true)] [String] $resource,
        [Parameter(Mandatory=$true)] [String] $clientID,
        [Parameter(Mandatory=$true)] [String] $clientSecret,
        [Parameter(Mandatory=$true)] [String] $loginURL,
        [Parameter(Mandatory=$true)] [String] $tenantDomain
    )   

    # Retrieve Oauth 2 access token
    $body = @{grant_type="client_credentials";resource=$resource;client_id=$clientID;client_secret=$clientSecret}
    $oauthToken = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body

    return $oauthToken
}


function Get-CurrentAuditSubscription()
{
    Param(
        [Parameter(Mandatory=$true)] [string] $tenantGUID,
        [Parameter(Mandatory=$true)] [Hashtable] $header
    )  
    $result = $false
    #Validate named Subscription is active.
    #If it isn't something has gone wrong with the O365 configuration or it's been removed.
    $subs = Invoke-WebRequest -Headers $header -Uri "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/list" | Select Content

    if ($subs.content.Length -le 2)
    {
        Write-Host "No audit subscriptions returned"
        return $false
    }
    else
    {
        $subsObject = $subs.content | ConvertFrom-Json
        return $subsObject
    }

    return $null
    
}



function fetch-BlobManifest
{
    Param(
        [Parameter(Mandatory=$true)] [Hashtable] $header,
        [Parameter(Mandatory=$true)] [string] $tenantGUID,
        [Parameter(Mandatory=$true)] [string] $auditSuffix,
        [Parameter(Mandatory=$false)] [string] $nextPageURI
    )      

    $requestURI = "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/content$auditSuffix"

    if ($nextPageURI)
    {
        $requestURI = $nextPageURI
    }

    $retryCount = 0
    $blobManifest = $null
    $callSuccess = $false
    do
    {
        $retryCount++
        if ($retryCount -gt 50)
        {
            throw "Maximum retry count hit"
        }
        try
        {
            $blobManifest = Invoke-WebRequest -Headers $header -Uri $requestURI
            $callSuccess = $true
        }
        catch 
        {
            $errorCaught = $error[0].ErrorDetails
            if ($errorCaught.Message)
            {
                $errorJSON = $errorCaught.Message | convertfrom-json
                if ($errorJSON.error.code -eq "AF429")
                {
                    #Need to back off for half a second.
                    write-host "Back off algorithm hit" -ForegroundColor Yellow
                    Write-Output "Back off algorithm hit"
                    start-sleep -m 500
                                
                }
                else
                {
                    throw
                }
            }
            else
            {
                throw
            }

        }
    }
    until ($callSuccess -eq $true)

    return $blobManifest
}

function fetch-BlobContent
{
    Param(
        [Parameter(Mandatory=$true)] [Hashtable] $header,
        [Parameter(Mandatory=$false)] [string] $blobURI
    )      

    $retryCount = 0
    $callSuccess = $false
    do
    {
        $retryCount++
        if ($retryCount -gt 50)
        {
            throw "Maximum retry count hit"
        }
        try
        {
            $blobData = Invoke-WebRequest -Headers $header -Uri $blobURI
            $callSuccess = $true
        }
        catch 
        {
            $errorCaught = $error[0].ErrorDetails
            if ($errorCaught.Message)
            {
                $errorJSON = $errorCaught.Message | convertfrom-json
                if ($errorJSON.error.code -eq "AF429")
                {
                    #Need to back off for half a second.
                    write-host "Back off algorithm hit" -ForegroundColor Yellow
                    Write-Output "Back off algorithm hit"
                    start-sleep -m 500
                                
                }
                else
                {
                    throw
                }
            }
            else
            {
                throw
            }

        }
    }
    until ($callSuccess -eq $true)

    return $blobData
}