. .\SupportFunctions.ps1
$globalConfig = Get-Content makingAuditGreatAgain-config.json -Raw | ConvertFrom-Json


###DEMO
##Here we get the oAuth access token that we use to authenticate our REST requests.
##Code for this is in SupportFunctions.ps1
$oAuth = get-oAuthAccessToken -resource $globalConfig.ResourceAPI -clientID $globalConfig.InvestigationAppID -clientSecret $globalConfig.InvestigationAppSecret -loginURL $globalConfig.LoginURL -tenantDomain $globalConfig.InvestigationTenantDomain
$header  = @{'Authorization'="$($oAuth.token_type) $($oAuth.access_token)"}


Get-CurrentAuditSubscription -tenantGUID $globalConfig.InvestigationTenantGUID -header $header



##Should we need to recreate a subscription we use:
#Invoke-RestMethod -Method Post -Headers $headerParams -Uri "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/start?contentType=Audit.SharePoint"


##Request the Manifest for a given date/time range and ContentType
$extractDate = (get-date).AddDays(-$globalConfig.ProcessLogOffsetDays)
$processDate = $extractDate.toString("yyyy-MM-dd")

##Forcing the date here just for SPS Belgium
$processDate = "2017-10-17"


#REST request is here. Content type and process window.
$auditSuffix = "?contentType=Audit.SharePoint&startTime=" + $processDate + "T00:00:00Z&endTime=" + $processDate + "T23:59:59Z"

$blobManifestConsolidated = @()
$blobList = @()


##Now fetch all of the available Blob Manifests for the given time period
##This may require multiple connections
do
{
    $BlobManifest = fetch-BlobManifest -header $header -tenantGUID $globalConfig.InvestigationTenantGUID -auditSuffix $auditSuffix -nextPageURI $BlobManifest.Headers.NextPageUri
    $blobManifestConsolidated += $blobManifest

}
while ($BlobManifest.Headers.NextPageUri)

foreach($blobManifestRetrieved in $blobManifestConsolidated)
{
    ##now process each manifest for this date
    $blobList += $blobManifestRetrieved.Content | ConvertFrom-Json
}


##Now process the Blob list and get the actual data.

$blobTotal = $blobList.count
$blobCount = 0
$auditRecords = @()
write-host "Blob package contains" $blobTotal "packages for processing"

##Setup our Azure SQL Connection
$connectionString = "Server=tcp:" + $globalConfig.SQLServerURI +
 ",1433;Database=" + $globalConfig.SQLDatabaseName + ";Uid=" + $globalConfig.SqlAccessAccount + ";Pwd=" + $globalConfig.SqlAccessAccountPwd +
  ";Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
$conn = new-object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = $connectionString
$command = new-object System.Data.SqlClient.SqlCommand
$command.Connection = $conn
$recordCount = 0

if ($blobTotal -gt 0)
{
    foreach ($blobDataSource in $blobList)
    {
        $blobCount++
        $statusmsg = "Processing blobs"
        $percentComplete = ($blobCount/$blobTotal)*100
        Write-Progress -Id 1 -Activity "Processing blob $blobCount of $blobTotal" -status $statusmsg -PercentComplete $percentComplete

        ##Depending on how long this process takes, we could have our app token expire.
        ##if this happens we need to refresh it.
        $tokenDate = Convert-FromUnixdate $oAuth.expires_on
        if ($tokenDate -gt (get-date).AddMinutes(5))
        {
            #Token still valid. Continue
        }
        else
        {
            # Token has less than 5 minutes, so refresh it.
            # Retrieve Oauth 2 access token
            write-host "Refreshing oAuth token" -ForegroundColor Yellow
            $oAuth = get-oAuthAccessToken -resource $globalConfig.ResourceAPI -clientID $globalConfig.InvestigationAppID -clientSecret $globalConfig.InvestigationAppSecret -loginURL $globalConfig.LoginURL -tenantDomain $globalConfig.InvestigationTenantDomain
            #Create the header params for this token.
            $header  = @{'Authorization'="$($oAuth.token_type) $($oAuth.access_token)"}
        }

        $blobData = fetch-BlobContent -header $header -blobURI $blobDataSource.contentUri

        $blobObjects = $blobData.Content | ConvertFrom-Json


        ##Now process the Blobobjects and see if we want them or not
        #Or just copy the lot..

        foreach ($auditEntry in $blobObjects)
        {


            #Apply business logic here!
            switch -Wildcard ($auditEntry.ObjectId.ToLower())
            {
                ##If we match the URL with wildcard, save the entry
                ##This could push it directly into SQL
                "https://tenantname.sharepoint.com/sites/audit-test*" {



                $ObjcreationTime = get-date $auditEntry.CreationTime
                $objId = $auditEntry.Id
                $objOperation = $auditEntry.Operation
                $objWorkload = $auditEntry.Workload
                $objClientIp = $auditEntry.ClientIP
                $objEventSource = $auditEntry.EventSource
                $objObjectId = $auditEntry.ObjectId
                $objUserId = $auditEntry.UserId
                $objSite = $auditEntry.Site
                $objFileExtension = $auditEntry.SourceFileExtension

                ##Now insert the values
                $command.commandtext = "INSERT INTO AuditRecords (CreationTime,Id,Operation,Workload,ClientIP,EventSource,ObjectId,UserId,Site,SourceFileExtension)" +
                 " VALUES ('$ObjcreationTime','$objId','$objOperation','$objWorkload','$objClientIp','$objEventSource','$objObjectId','$objUserId','$objSite','$objFileExtension')"

                $command.CommandType = [System.Data.CommandType]::Text

                $conn.Open()
                $command.ExecuteNonQuery()
                $conn.close()
                $recordCount++

                }


                default {
                    ##Default behaviour hit
                    ##Do nothing as it's not a focus site.
                    ##We could write this out to a seperate file if we wanted.
                
                }
            }

        }
    }
}

write-host "$recordCount entries out of $blobTotal written to SQL"