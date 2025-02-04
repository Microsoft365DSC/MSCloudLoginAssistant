function Connect-MSCloudLoginLicensing
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginLicensing'

    if (-not $Script:MSCloudLoginConnectionProfile.Licensing.AccessToken)
    {
        try
        {
            if ($Script:MSCloudLoginConnectionProfile.Licensing.AuthenticationType -eq 'CredentialsWithApplicationId' -or
                $Script:MSCloudLoginConnectionProfile.Licensing.AuthenticationType -eq 'Credentials' -or
                $Script:MSCloudLoginConnectionProfile.Licensing.AuthenticationType -eq 'CredentialsWithTenantId')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
                Connect-MSCloudLoginLicensingWithUser
            }
            elseif ($Script:MSCloudLoginConnectionProfile.Licensing.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Admin API using AAD App {$ApplicationID}" -Source $source
                Connect-MSCloudLoginLicensingWithCertificateThumbprint
            }
            else
            {
                throw 'Specified authentication method is not supported.'
            }

            $Script:MSCloudLoginConnectionProfile.Licensing.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $true
            $Script:MSCloudLoginConnectionProfile.Licensing.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Admin API using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
}

function Connect-MSCloudLoginLicensingWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginLicensingWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Licensing.TenantId))
    {
        $tenantid = $Script:MSCloudLoginConnectionProfile.Licensing.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Licensing.TenantId
    }
    $username = $Script:MSCloudLoginConnectionProfile.Licensing.Credentials.UserName
    $password = $Script:MSCloudLoginConnectionProfile.Licensing.Credentials.GetNetworkCredential().password

    $clientId = '1950a258-227b-4e31-a9cf-717495945fc2'
    $uri = "$($Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl)/{0}/oauth2/token" -f $tenantid
    $body = "resource=aeb86249-8ea3-49e2-900b-54cc8e308f85&client_id=$clientId&grant_type=password&username={1}&password={0}" -f [System.Web.HttpUtility]::UrlEncode($password), $username

    # Request token through ROPC
    try
    {
        $managementToken = Invoke-RestMethod $uri `
            -Method POST `
            -Body $body `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction SilentlyContinue

        $Script:MSCloudLoginConnectionProfile.Licensing.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Licensing.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginLicensingWithUserMFA
        }
    }
}
function Connect-MSCloudLoginLicensingWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Licensing.TenantId))
    {
        $tenantid = $Script:MSCloudLoginConnectionProfile.Licensing.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Licensing.TenantId
    }
    $clientId = '1950a258-227b-4e31-a9cf-717495945fc2'
    $deviceCodeUri = "$($Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl)/$tenantId/oauth2/devicecode"

    $body = @{
        client_id = $clientId
        resource  = $Script:MSCloudLoginConnectionProfile.Licensing.AdminUrl
    }
    $DeviceCodeRequest = Invoke-RestMethod $deviceCodeUri `
        -Method POST `
        -Body $body

    Write-Host "`r`n$($DeviceCodeRequest.message)" -ForegroundColor Yellow

    $TokenRequestParams = @{
        Method = 'POST'
        Uri    = "$($Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl)/$TenantId/oauth2/token"
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
    $Script:MSCloudLoginConnectionProfile.Licensing.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $true
    $Script:MSCloudLoginConnectionProfile.Licensing.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.Licensing.ConnectedDateTime = [System.DateTime]::Now.ToString()
}

function Connect-MSCloudLoginLicensingWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginLicensingWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Licensing using CertificateThumbprint' -Source $source
    $tenantId = $Script:MSCloudLoginConnectionProfile.Licensing.TenantId

    try
    {
        $Certificate = Get-Item "Cert:\CurrentUser\My\$($Script:MSCloudLoginConnectionProfile.Licensing.CertificateThumbprint)" -ErrorAction SilentlyContinue

        if ($null -eq $Certificate)
        {
            Add-MSCloudLoginAssistantEvent 'Certificate not found in CurrentUser\My, trying LocalMachine\My' -Source $source

            $Certificate = Get-ChildItem "Cert:\LocalMachine\My\$($Script:MSCloudLoginConnectionProfile.Licensing.CertificateThumbprint)" -ErrorAction SilentlyContinue

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
            aud = "$($Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl)/$TenantId/oauth2/token"

            # Expiration timestamp
            exp = $JWTExpiration

            # Issuer = your application
            iss = $Script:MSCloudLoginConnectionProfile.Licensing.ApplicationID

            # JWT ID: random guid
            jti = [guid]::NewGuid()

            # Not to be used before
            nbf = $NotBefore

            # JWT Subject
            sub = $Script:MSCloudLoginConnectionProfile.Licensing.ApplicationID
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
            client_id             = $Script:MSCloudLoginConnectionProfile.Licensing.ApplicationID
            client_assertion      = $JWT
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            scope                 = $Script:MSCloudLoginConnectionProfile.Licensing.Scope
            grant_type            = 'client_credentials'
        }

        $Url = "$($Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl)/$TenantId/oauth2/v2.0/token"

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
        $Script:MSCloudLoginConnectionProfile.Licensing.AccessToken = 'Bearer ' + $Request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Admin API API using Certificate Thumbprint' -Source $source

        $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Licensing.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        throw $_
    }
}

function Disconnect-MSCloudLoginLicensing
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginLicensing'

    if ($Script:MSCloudLoginConnectionProfile.Licensing.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from Licensing API' -Source $source
        $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Licensing API' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Licensing API were found' -Source $source
    }
}
