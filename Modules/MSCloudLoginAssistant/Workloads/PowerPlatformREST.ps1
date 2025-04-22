function Connect-MSCloudLoginPowerPlatformREST
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginPowerPlatformREST'

    # Test authentication to make sure the token hasn't expired
    try
    {
        $uri = "https://" + $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.BapEndpoint + `
               "/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments"
        $headers = @{
            Authorization = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken
        }
        $response = Invoke-WebRequest -Method 'GET' `
                                      -Uri $Uri `
                                      -Headers $headers `
                                      -ContentType 'application/json; charset=utf-8' `
                                      -UseBasicParsing `
                                      -ErrorAction Stop
    }
    catch
    {
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $null
    }

    if (-not $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken)
    {
        try
        {
            if ($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType -eq 'CredentialsWithApplicationId' -or
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType -eq 'Credentials' -or
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType -eq 'CredentialsWithTenantId')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
                Connect-MSCloudLoginPowerPlatformRESTWithUser
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Admin API using AAD App {$ApplicationID}" -Source $source
                Connect-MSCloudLoginPowerPlatformRESTWithCertificateThumbprint
            }
            else
            {
                throw 'Specified authentication method is not supported.'
            }

            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Admin API using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
}

function Connect-MSCloudLoginPowerPlatformRESTWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginPowerPlatformRESTWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId))
    {
        $tenantid = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId
    }
    $username = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Credentials.UserName
    $password = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Credentials.GetNetworkCredential().password

    $clientId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ClientId
    $uri = "$($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl)/{0}/oauth2/token" -f $tenantid
    $body = "resource=aeb86249-8ea3-49e2-900b-54cc8e308f85&client_id=$clientId&grant_type=password&username={1}&password={0}" -f [System.Web.HttpUtility]::UrlEncode($password), $username

    # Request token through ROPC
    try
    {
        $managementToken = Invoke-RestMethod $uri `
            -Method POST `
            -Body $body `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction SilentlyContinue

        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginPowerPlatformRESTWithUserMFA
        }
    }
}
function Connect-MSCloudLoginPowerPlatformRESTWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId))
    {
        $tenantid = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId
    }
    $clientId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ClientId
    $deviceCodeUri = "$($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl)/$tenantId/oauth2/devicecode"

    $body = @{
        client_id = $clientId
        resource  = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AdminUrl
    }
    $DeviceCodeRequest = Invoke-RestMethod $deviceCodeUri `
        -Method POST `
        -Body $body

    Write-Host "`r`n$($DeviceCodeRequest.message)" -ForegroundColor Yellow

    $TokenRequestParams = @{
        Method = 'POST'
        Uri    = "$($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl)/$TenantId/oauth2/token"
        Body   = @{
            grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
            code       = $DeviceCodeRequest.device_code
            client_id  = $clientId
        }
    }
    $TimeoutTimer = [System.Diagnostics.Stopwatch]::StartNew()
    while ([string]::IsNullOrEmpty($managementToken.access_token))
    {
        if ($TimeoutTimer.Elapsed.TotalSeconds -gt 300)
        {
            throw 'Login timed out, please try again.'
        }
        $managementToken = try
        {
            Invoke-RestMethod @TokenRequestParams -ErrorAction Stop
        }
        catch
        {
            $Message = $_.ErrorDetails.Message | ConvertFrom-Json
            if ($Message.error -ne 'authorization_pending')
            {
                throw
            }
        }
        Start-Sleep -Seconds 1
    }
    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
}

function Connect-MSCloudLoginPowerPlatformRESTWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginPowerPlatformRESTWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to PowerPlatformREST using CertificateThumbprint' -Source $source
    $tenantId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId

    try
    {
        $Certificate = Get-Item "Cert:\CurrentUser\My\$($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.CertificateThumbprint)" -ErrorAction SilentlyContinue

        if ($null -eq $Certificate)
        {
            Add-MSCloudLoginAssistantEvent 'Certificate not found in CurrentUser\My, trying LocalMachine\My' -Source $source

            $Certificate = Get-ChildItem "Cert:\LocalMachine\My\$($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.CertificateThumbprint)" -ErrorAction SilentlyContinue

            if ($null -eq $Certificate)
            {
                throw 'Certificate not found in LocalMachine\My nor CurrentUser\My'
            }
        }
        # Create base64 hash of certificate
        $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())

        # Create JWT timestamp for expiration
        $StartDate = (Get-Date '1970-01-01T00:00:00Z' ).ToUniversalTime()
        $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
        $JWTExpiration = [math]::Round($JWTExpirationTimeSpan, 0)

        # Create JWT validity start timestamp
        $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
        $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan, 0)

        # Create JWT header
        $JWTHeader = @{
            alg = 'RS256'
            typ = 'JWT'
            # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
            x5t = $CertificateBase64Hash -replace '\+', '-' -replace '/', '_' -replace '='
        }

        # Create JWT payload
        $JWTPayLoad = @{
            # What endpoint is allowed to use this JWT
            aud = "$($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl)/$TenantId/oauth2/token"

            # Expiration timestamp
            exp = $JWTExpiration

            # Issuer = your application
            iss = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ApplicationId

            # JWT ID: random guid
            jti = [guid]::NewGuid()

            # Not to be used before
            nbf = $NotBefore

            # JWT Subject
            sub = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ApplicationId
        }

        # Convert header and payload to base64
        $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
        $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

        $JWTPayLoadToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
        $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

        # Join header and Payload with "." to create a valid (unsigned) JWT
        $JWT = $EncodedHeader + '.' + $EncodedPayload

        # Get the private key object of your certificate
        $PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))

        # Define RSA signature and hashing algorithm
        $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
        $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

        # Create a signature of the JWT
        $Signature = [Convert]::ToBase64String(
            $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT), $HashAlgorithm, $RSAPadding)
        ) -replace '\+', '-' -replace '/', '_' -replace '='

        # Join the signature to the JWT with "."
        $JWT = $JWT + '.' + $Signature

        # Create a hash with body parameters
        $Body = @{
            client_id             = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ApplicationId
            client_assertion      = $JWT
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            scope                 = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Audience + "/.default"
            grant_type            = 'client_credentials'
        }

        $Url = "$($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl)/$TenantId/oauth2/v2.0/token"

        # Use the self-generated JWT as Authorization
        $Header = @{
            Authorization = "Bearer $JWT"
        }

        # Splat the parameters for Invoke-Restmethod for cleaner code
        $PostSplat = @{
            ContentType = 'application/x-www-form-urlencoded'
            Method      = 'POST'
            Body        = $Body
            Uri         = $Url
            Headers     = $Header
        }

        $Request = Invoke-RestMethod @PostSplat

        # View access_token
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = 'Bearer ' + $Request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Admin API API using Certificate Thumbprint' -Source $source

        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        throw $_
    }
}

function Disconnect-MSCloudLoginPowerPlatformREST
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginPowerPlatformREST'

    if ($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from PowerPlatformREST API' -Source $source
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from PowerPlatformREST API' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to PowerPlatformREST API were found' -Source $source
    }
}
