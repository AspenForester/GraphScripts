#Requires -module ImportExcel
#requires -version 7.0

$ReportPath = "C:\yourpath\"
$TenantName = "YourTenant.onmicrosoft.com"
$AppId = "{YourAppIDGUID}"
$CertificateThumprint = 'YourCertThumbprint'

# https://docs.microsoft.com/en-us/graph/api/overview?toc=.%2Fref%2Ftoc.json&view=graph-rest-1.0

# getting the OAuth token is lifted directly from 
# https://adamtheautomator.com/microsoft-graph-api-powershell

#region authentication
$Scope = "https://graph.microsoft.com/.default"
$Certificate = Get-Item "Cert:\CurrentUser\My\$CertificateThumprint"

# Create base64 hash of certificate
$CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())

# Create JWT timestamp for expiration
$StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
$JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
$JWTExpiration = [math]::Round($JWTExpirationTimeSpan, 0)

# Create JWT validity start timestamp
$NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
$NotBefore = [math]::Round($NotBeforeExpirationTimeSpan, 0)

# Create JWT header
$JWTHeader = @{
    alg = "RS256"
    typ = "JWT"
    # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
    x5t = $CertificateBase64Hash -replace '\+', '-' -replace '/', '_' -replace '='
}

# Create JWT payload
$JWTPayLoad = @{
    # What endpoint is allowed to use this JWT
    aud = "https://login.microsoftonline.com/$TenantName/oauth2/token"

    # Expiration timestamp
    exp = $JWTExpiration

    # Issuer = your application
    iss = $AppId

    # JWT ID: random guid
    jti = [guid]::NewGuid()

    # Not to be used before
    nbf = $NotBefore

    # JWT Subject
    sub = $AppId
}

# Convert header and payload to base64
$JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
$EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

$JWTPayLoadToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
$EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

# Join header and Payload with "." to create a valid (unsigned) JWT
$JWT = $EncodedHeader + "." + $EncodedPayload

# Get the private key object of your certificate
$PrivateKey = $Certificate.PrivateKey

# Define RSA signature and hashing algorithm
$RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
$HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

# Create a signature of the JWT
$Signature = [Convert]::ToBase64String(
    $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT), $HashAlgorithm, $RSAPadding)
) -replace '\+', '-' -replace '/', '_' -replace '='

# Join the signature to the JWT with "."
$JWT = $JWT + "." + $Signature

# Create a hash with body parameters
$Body = @{
    client_id             = $AppId
    client_assertion      = $JWT
    client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    scope                 = $Scope
    grant_type            = "client_credentials"
}

$TokenUrl = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"

# Use the self-generated JWT as Authorization
$JWTHeader = @{
    Authorization = "Bearer $JWT"
}

# Splat the parameters for Invoke-Restmethod for cleaner code
$PostSplat = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method      = 'POST'
    Body        = $Body
    Uri         = $TokenUrl
    Headers     = $JWTHeader
}

$TokenRequest = Invoke-RestMethod @PostSplat
#endregion

$Header = @{
    Authorization = "$($TokenRequest.token_type) $($TokenRequest.access_token)"
}


# Now let's do some work.
try
{
    # get Group Lifecycle Policy Info
    $GLCPolicyURI = "https://graph.microsoft.com/v1.0/groupLifecyclePolicies"
    $GLCPolicyRequest = Invoke-RestMethod -uri $GLCPolicyURI -Headers $header -Method Get 
    # $GLCPolicyRequest.value[0].groupLifetimeInDays
    
    # Get all Unified (O365) Groups
    $UnifiedGroupsURI = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c+eq+'Unified')"
    $AllUnifiedgroupsRequest = Invoke-RestMethod -uri $UnifiedGroupsURI -Headers $header -Method Get -ErrorAction Stop 

    # Work through all the pages of results.
    $AllGroups = $AllUnifiedGroupsRequest.value
    while ($null -ne $AllUnifiedGroupsRequest.'@odata.nextLink')
    {
        $AllUnifiedGroupsRequest = Invoke-RestMethod -uri $AllUnifiedGroupsRequest.'@odata.nextLink' -Method Get -Headers $header 
        $AllGroups += $AllUnifiedGroupsRequest.value
    }

    $Teams = $AllGroups | where-object resourceProvisioningOptions -eq "Team"

    #foreach ($team in $teams )
    $TeamsReport = $teams | foreach-object -parallel {
        $team = $_
        # This is where the batch requests would be handy
        $batch = @{requests = @(
                @{
                    id     = 1
                    method = "GET"
                    url    = "/teams/$($team.id)"
                },
                @{
                    id     = 2
                    method = "GET"
                    url    = "/groups/$($team.id)/members"
                },
                @{
                    id     = 3
                    method = "GET"
                    url    = "/groups/$($team.id)/owners"
                }
            )
        }
        $batchjson = $batch | ConvertTo-Json
        $BatchSplat = @{
            Method = "POST"
            URI = 'https://graph.microsoft.com/v1.0/$batch'
            ContentType = "application/json"
            Body = $batchjson
            Headers = $Using:Header
            ErrorAction = "Stop"
        } 
        $batchRequest = Invoke-RestMethod @BatchSplat

        $TeamRequest = ($batchRequest.responses | where-object id -eq 1).body
        $memberRequest = ($batchRequest.responses | where-object id -eq 2).body
        $OwnerRequest = ($batchRequest.responses | Where-Object id -eq 3).body.value # We'll assume for that there are fewer than 100 owners
    
        # Deal with member pagination!
        $allmembers = $memberRequest.value 
        while ($null -ne $memberRequest.'@odata.nextlink')
        {
            $memberRequest = Invoke-RestMethod -Uri $memberRequest.'@odata.nextlink' -Method Get -Headers $Using:Header
            $allmembers += $memberRequest.value
        }

        # Separate the members and the guests
        $Members = $allmembers | Where-Object userPrincipalName -notlike "*EXT*"
        $Guests = $allmembers | Where-Object userPrincipalName -like "*EXT*"

        # Calculate the Expiration Date - Not sure if this is needed!
        $ExpirationDate = $Team.renewedDateTime.addDays($($Using:GLCPolicyRequest).value[0].groupLifetimeInDays)

        # Teams worksheet
        [pscustomobject]@{
            GroupID                           = $teamrequest.id
            DisplayName                       = $TeamRequest.displayName
            Description                       = $TeamRequest.description
            Visibility                        = $Team.visibility
            MailNickName                      = $Team.MailNickName
            Classification                    = $TeamRequest.classification
            Archived                          = $TeamRequest.isArchived
            AllowGiphy                        = $TeamRequest.funSettings.allowGiphy
            GiphyContentRating                = $TeamRequest.funSettings.giphyContentRating
            AllowStickersAndMemes             = $TeamRequest.funSettings.allowStickersAndMemes
            AllowCustomMemes                  = $TeamRequest.funSettings.allowCustomMemes
            AllowGuestCreateUpdateChannels    = $TeamRequest.guestSettings.allowCreateUpdateChannels
            AllowGuestDeleteChannels          = $TeamRequest.guestSettings.allowDeleteChannels
            AllowCreateUpdateChannels         = $TeamRequest.memberSettings.allowCreateUpdateChannels
            AllowDeleteChannels               = $TeamRequest.memberSettings.allowDeleteChannels
            AllowAddRemoveApps                = $TeamRequest.memberSettings.allowAddRemoveApps
            AllowCreateUpdateRemoveTabs       = $TeamRequest.memberSettings.allowCreateUpdateRemoveTabs
            AllowCreateUpdateRemoveConnectors = $TeamRequest.memberSettings.allowCreateUpdateRemoveConnectors
            AllowUserEditMessages             = $TeamRequest.messagingSettings.allowUserEditMessages
            AllowUserDeleteMessages           = $TeamRequest.messagingSettings.allowUserDeleteMessages
            AllowOwnerDeleteMessages          = $TeamRequest.messagingSettings.allowOwnerDeleteMessages
            AllowTeamMentions                 = $TeamRequest.messagingSettings.allowTeamMentions
            AllowChannelMentions              = $TeamRequest.messagingSettings.allowChannelMentions
            ShowInTeamsSearchAndSuggestions   = $TeamRequest.discoverySettings.showInTeamsSearchAndSuggestions
            Owner                             = $OwnerRequest.displayname -join "; "
            OwnerCount                        = $OwnerRequest.Count
            Member                            = $Members.displayName -join "; "
            Guests                            = $Guests.displayName -join "; "
            MemberCount                       = $Allmembers.Count
            WhenCreated                       = $Team.createdDateTime
            RenewedDateTime                   = $Team.renewedDateTime
            ExpirationDate                    = $ExpirationDate 
        }
    } # End Foreach

    # RecycleBin worksheet  --  ?$filter=groupTypes/any(a:a eq 'unified')
    $DeletedGroupsURI = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group"
    $DeletedGroupsRequest = Invoke-RestMethod -uri $DeletedGroupsURI -Headers $header -Method Get 

    # Handle Pagination
    $AllDeletedGroups = $DeletedGroupsRequest.value
    while ($null -ne $DeletedGroupsRequest.'@odata.nextLink')
    {
        $DeletedGroupsRequest = Invoke-RestMethod -uri $DeletedGroupsRequest.'@odata.nextLink' -Method Get -Headers $header -ErrorAction Stop 
        $AllDeletedGroups += $DeletedGroupsRequest.value
    }
    $AllDeletedGroups = $AllDeletedGroups | 
    Where-Object resourceProvisioningOptions -eq "Team" | 
    Select-object -Property Id, Displayname, Description, CreatedDateTime, RenewedDateTime, DeletedDateTime,  
        Mail, MailEnabled, MailNickname, @{N = 'GroupTypes'; E = { $_.GroupTypes } }, Visibility
}
finally
{
    # IF the script is interrupted, it will still output all of the data it has gathered to the point of interruption.
    $Today = Get-Date -Format yyyyMMdd    
    
    $TodaysReport = Join-path -Path $ReportPath -ChildPath "TeamsReport-$Today.xlsx"

    Write-Verbose -Message "$(Get-Date -format "MM/dd/yyyy HH:mm:ss") - Writing reports to Excel..." -verbose
    $TeamsReport | Export-Excel -Path $TodaysReport -WorksheetName 'Teams' -ClearSheet
    $DeletedTeams | Export-Excel -Path $TodaysReport -WorksheetName 'RecycleBin' -ClearSheet 
}
