<#
.SYNOPSIS
    Enterprise Provisioning Script for Hyland Identity Service 5.1.x & OnBase API Server 25.1.x.
.DESCRIPTION
    - Full self-healing and parameterized schema reconciliation architecture.
    - Preserves and self-heals complete factory schemas for:
        * IdS: config\appsettings.json, config\idsconfig.json
        * API Server: config\1_server.json, config\5_scim.json, config\5_tempfs.json, appsettings.json
    - Configures idsconfig.json AdminUiClient, RedirectUris, PostLogoutRedirectUris, and CORS origins.
    - Enforces 'Encrypt=True;TrustServerCertificate=True;' on all SQL connection strings.
    - Configures 1_server.json (Host.Authentication.Provider, TempFileStorage.Address, OnBaseSessionCookie, Tenant).
    - Configures 5_scim.json (Enabled, JwtIssuerOptions, DataSources array with System.Data.SqlClient).
    - Configures 5_tempfs.json (Full TempFS container, BackgroundProcessor, Security, Tenant DDL/SQL array).
    - Registers required IIS modules (AnonymousAuthentication, AspNetCoreModuleV2, HttpCache, UriCache).
    - Grants Read ACL on certificate private keys and MachineKeys to IIS_IUSRS.
#>
[CmdletBinding()]
param(
    # --- Settings File Input ---
    [Parameter(Position = 1)]
    [string]$SettingsFilePath = '.\Settings.json',

    [ValidateSet('Local', 'Remote')]
    [string]$DeploymentMode,

    [string]$TargetMachine,

    [Parameter(Position = 0)]
    [pscredential]$Credential,

    [string]$ConfigurationName,

    # --- Source Packages (UNC Shares - Optional Fallbacks) ---
    [string]$IdSSourcePath,
    [string]$ApiServerSourcePath,
    [string]$DotNetHostingBundlePath,

    # --- ProGet & Version Parameters ---
    [string]$ProGetFeedUrl,
    [string]$IdSVersion,
    [string]$ApiServerVersion,
    [string]$IdSPackageId,
    [string]$ApiServerPackageId,

    [string]$TempFsConfigPackageId,
    [string]$TempFsConfigVersion,

    # --- Database & Tenant Configuration ---
    [string]$SqlInstance,
    [string]$DatabaseName,
    [pscredential]$HsiCredential,

    # --- Identity & Web Endpoints ---
    [string]$IdSAppName,
    [string]$ApiAppName,
    [string]$CertificateThumbprint,

    [string[]]$IdsAdminUsers,
    [string[]]$IdsAdminGroups,
    [string]$IdsProvidersJson,

    [switch]$ForceReinstall
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Helper: StrictMode-safe property getter
function Get-SafeProperty {
    param($TargetObject, [string]$PropertyName, $DefaultValue = $null)
    if ($null -ne $TargetObject -and $TargetObject.PSObject.Properties[$PropertyName]) {
        $val = $TargetObject.$PropertyName
        if ($null -ne $val -and [string]$val -ne '') { return $val }
    }
    return $DefaultValue
}

# ==============================================================================
# Parse Settings.json (Auto-loads .\Settings.json by default)
# ==============================================================================
if ($SettingsFilePath -and (Test-Path $SettingsFilePath)) {
    Write-Host "[Settings] Loading configuration from: $SettingsFilePath" -ForegroundColor Cyan
    $jsonContent = Get-Content -Path $SettingsFilePath -Raw | ConvertFrom-Json

    # Root Settings
    if (-not $TargetMachine) { $TargetMachine = Get-SafeProperty $jsonContent 'machineName' (Get-SafeProperty $jsonContent 'TargetServer' $env:COMPUTERNAME) }
    if (-not $DeploymentMode) { $DeploymentMode = Get-SafeProperty $jsonContent 'deploymentMode' 'Remote' }
    if (-not $ConfigurationName) { $ConfigurationName = Get-SafeProperty $jsonContent 'configurationName' 'UnityInstall' }
    if (-not $DatabaseName) { $DatabaseName = Get-SafeProperty $jsonContent 'datasourceName' 'OnBase' }
    if (-not $SqlInstance) { $SqlInstance = Get-SafeProperty $jsonContent 'sqlInstance' "$TargetMachine\SQL2K22" }
    if (-not $CertificateThumbprint) { $CertificateThumbprint = Get-SafeProperty $jsonContent 'thumbprint' '' }
    if (-not $ProGetFeedUrl) { $ProGetFeedUrl = Get-SafeProperty $jsonContent 'proGetFeedUrl' 'https://proget.onbase.net/nuget/Release' }

    # Nested IdS Settings
    $idpObj = Get-SafeProperty $jsonContent 'idp'
    if ($idpObj) {
        if (-not $IdSAppName) { $IdSAppName = Get-SafeProperty $idpObj 'appName' 'identityservice' }
        if (-not $IdSVersion) { $IdSVersion = Get-SafeProperty $idpObj 'version' '5.1.0' }
        if (-not $IdSPackageId) { $IdSPackageId = Get-SafeProperty $idpObj 'packageId' 'Hyland.Identity.Server.Setup' }
        if (-not $IdsAdminUsers) {
            $u = Get-SafeProperty $idpObj 'adminUsers'
            if ($u) { $IdsAdminUsers = @($u) }
        }
        if (-not $IdsAdminGroups) {
            $g = Get-SafeProperty $idpObj 'adminGroups'
            if ($g) { $IdsAdminGroups = @($g) }
        }
        <#
        if (-not $IdsProvidersJson) {
            $p = Get-SafeProperty $idpObj 'providers'
            if ($p) { $IdsProvidersJson = ConvertTo-Json -InputObject @($p) -Depth 20 -Compress }
        }
        #>
        if (-not $IdsProvidersJson -and $idpObj.providers) {
            $IdsProvidersJson = ConvertTo-Json -InputObject @($idpObj.providers) -Depth 20 -Compress
        }
    }

    # Nested API Server Settings
    $apiObj = Get-SafeProperty $jsonContent 'apiServer'
    if ($apiObj) {
        if (-not $ApiAppName) { $ApiAppName = Get-SafeProperty $apiObj 'appName' 'apiserver' }
        if (-not $ApiServerVersion) { $ApiServerVersion = Get-SafeProperty $apiObj 'version' '25.1.28.1000' }
        if (-not $ApiServerPackageId) { $ApiServerPackageId = Get-SafeProperty $apiObj 'packageId' 'Hyland.ApiServer.Installer' }
    }

    # Nested TempFS Config Tool Settings
    $tfsObj = Get-SafeProperty $jsonContent 'tempFsConfig'
    if ($tfsObj) {
        if (-not $TempFsConfigPackageId) { $TempFsConfigPackageId = Get-SafeProperty $tfsObj 'packageId' 'TempFileService.Configuration' }
        if (-not $TempFsConfigVersion) { $TempFsConfigVersion = Get-SafeProperty $tfsObj 'version' '1.9.0' }
    }

    # HSI Credentials from Settings.json
    if (-not $HsiCredential) {
        $hsiUser = Get-SafeProperty $jsonContent 'hsiUsername' 'hsi'
        $hsiPass = Get-SafeProperty $jsonContent 'hsiPassword' 'wstinol'
        if ($hsiUser -and $hsiPass) {
            $secPass = ConvertTo-SecureString $hsiPass -AsPlainText -Force
            $HsiCredential = New-Object System.Management.Automation.PSCredential($hsiUser, $secPass)
        }
    }
}

# Parameter Defaults if still empty
if (-not $DeploymentMode) { $DeploymentMode = 'Remote' }
if (-not $IdSAppName) { $IdSAppName = 'identityservice' }
if (-not $ApiAppName) { $ApiAppName = 'apiserver' }
if (-not $IdSVersion) { $IdSVersion = '5.1.0' }
if (-not $ApiServerVersion) { $ApiServerVersion = '25.1.28.1000' }
if (-not $TempFsConfigPackageId) { $TempFsConfigPackageId = 'TempFileService.Configuration' }
if (-not $TempFsConfigVersion) { $TempFsConfigVersion = '1.9.0' }
if (-not $DatabaseName) { $DatabaseName = 'OnBase' }

# Parameter Validation
if (-not $SqlInstance) { throw "SqlInstance is required (via CLI parameter or Settings.json)." }
if (-not $HsiCredential) { throw "HsiCredential is required (via CLI parameter or Settings.json)." }

Write-Host "[Settings] TempFS config tool: $TempFsConfigPackageId $TempFsConfigVersion" -ForegroundColor Cyan

# ==============================================================================
# Main Installation & Configuration Engine (Remote Scope)
# ==============================================================================
$installLogic = {
    param(
        [string]$IdSSrc,
        [string]$ApiSrc,
        [string]$DotNetSrc,
        [string]$ProGetFeed,
        [string]$IdSVer,
        [string]$ApiVer,
        [string]$IdSPackage,
        [string]$ApiPackage,
        [string]$SqlInst,
        [string]$DbName,
        [pscredential]$HsiCred,
        [string]$IdSApp,
        [string]$ApiApp,
        [string]$CertThumb,
        [bool]$Force,
        [string]$TempFsCfgPackage,
        [string]$TempFsCfgVer,
        [string[]]$AdminUsers,
        [string[]]$AdminGroups,
        [string]$ProvidersJson
    )
    Write-Host "[DEBUG] Remote ProvidersJson: $(if ($ProvidersJson) { "$($ProvidersJson.Length) chars" } else { 'EMPTY' })" -ForegroundColor Magenta

    # Make ConvertTo-Json write arrays as [...] (stops the PS 5.1 {value,Count} wrapper)
    Remove-TypeData -TypeName System.Array -ErrorAction SilentlyContinue

    # Re-indents JSON text: 2 spaces, one property per line, empty arrays/objects as [] / {}
    function Format-JsonText {
        param([Parameter(Mandatory)][string]$Json, [int]$IndentSize = 2)
        $nl = "`r`n"
        $sb = New-Object System.Text.StringBuilder
        $lvl = 0; $inStr = $false; $esc = $false
        for ($i = 0; $i -lt $Json.Length; $i++) {
            $c = $Json[$i]
            if ($inStr) {                                   # copy string contents unchanged
                [void]$sb.Append($c)
                if ($esc) { $esc = $false }
                elseif ($c -eq [char]'\') { $esc = $true }
                elseif ($c -eq [char]'"') { $inStr = $false }
                continue
            }
            if ([char]::IsWhiteSpace($c)) { continue }      # drop the old formatting
            if ($c -eq [char]'"') { $inStr = $true; [void]$sb.Append($c) }
            elseif ($c -eq [char]'{' -or $c -eq [char]'[') {
                $close = if ($c -eq [char]'{') { [char]'}' } else { [char]']' }
                $j = $i + 1
                while ($j -lt $Json.Length -and [char]::IsWhiteSpace($Json[$j])) { $j++ }
                if ($j -lt $Json.Length -and $Json[$j] -eq $close) {
                    [void]$sb.Append($c).Append($close); $i = $j          # empty: [] or {}
                } else {
                    $lvl++; [void]$sb.Append($c).Append($nl).Append(' ' * ($lvl * $IndentSize))
                }
            }
            elseif ($c -eq [char]'}' -or $c -eq [char]']') {
                $lvl--; [void]$sb.Append($nl).Append(' ' * ($lvl * $IndentSize)).Append($c)
            }
            elseif ($c -eq [char]',') { [void]$sb.Append(',').Append($nl).Append(' ' * ($lvl * $IndentSize)) }
            elseif ($c -eq [char]':') { [void]$sb.Append(': ') }
            else { [void]$sb.Append($c) }
        }
        $sb.ToString()
    }

    # Unescapes, formats, validates, then writes UTF-8 without BOM
    function Save-JsonFile {
        param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Json, [int]$IndentSize = 2)
        $Json = $Json -replace '\\u003c','<' -replace '\\u003e','>' -replace '\\u0027',"'" -replace '\\u0026','&'
        $pretty = Format-JsonText -Json $Json -IndentSize $IndentSize
        $null = $pretty | ConvertFrom-Json            # throws if invalid, so the file is not touched
        [IO.File]::WriteAllText($Path, $pretty + "`r`n", (New-Object System.Text.UTF8Encoding $false))
    }

    $ErrorActionPreference = 'Stop'

    # -------------------------------------------------------------------------
    # Core JSON Utilities
    # -------------------------------------------------------------------------
    function Set-JsonPropertySafe {
        param(
            [Parameter(Mandatory = $true)]
            $TargetObject,
            [Parameter(Mandatory = $true)]
            [string]$PropertyName,
            $PropertyValue
        )
        if ($null -ne $TargetObject) {
            if ($TargetObject.PSObject.Properties[$PropertyName]) {
                $TargetObject.$PropertyName = $PropertyValue
            } else {
                $TargetObject | Add-Member -NotePropertyName $PropertyName -NotePropertyValue $PropertyValue -Force
            }
        }
    }

    function Save-CleanJson {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, Position = 0)]
            [Alias('Path')]
            [string]$FilePath,

            [Parameter(Mandatory, Position = 1)]
            [Alias('JsonObject', 'Json', 'Content', 'Object')]
            [object]$JsonContent,

            [int]$IndentSize = 2,
            [int]$Depth = 100
        )

        $jsonString = if ($JsonContent -is [string]) {
            $JsonContent
        } else {
            ConvertTo-Json -InputObject $JsonContent -Depth $Depth
        }

        Save-JsonFile -Path $FilePath -Json $jsonString -IndentSize $IndentSize
        Write-Host "  [OK] Saved and formatted: $(Split-Path $FilePath -Leaf)" -ForegroundColor Green
    }

    function Read-AndCleanJson {
        param([string]$FilePath)
        if (-not (Test-Path $FilePath)) { return $null }
        try {
            $raw = Get-Content -Path $FilePath -Raw -Encoding UTF8
            $clean = [regex]::Replace($raw, '/\*[\s\S]*?\*/', '')
            $clean = [regex]::Replace($clean, '(?m)^\s*`//.*$`', '')
            $clean = [regex]::Replace($clean, ',\s*([\}\]])', '$1')
            return ($clean | ConvertFrom-Json)
        } catch {
            return $null
        }
    }

    # -------------------------------------------------------------------------
    # Modular Configuration Patch Functions (Using $Config Parameter)
    # -------------------------------------------------------------------------

    # 1. IdS: config\appsettings.json
    function Update-IdSAppSettingsConfig {
        param([string]$FilePath, $Config)

        Write-Host "  Patching $FilePath with target schema..." -ForegroundColor Yellow
        $json = [ordered]@{
            'SigningCertificate' = [ordered]@{
                'Thumbprint' = $Config.CertificateThumbprint
            }
            'KeyEncryptionCertificateThumbprint' = ''
            'KeyFilePersistenceLocation'         = '%LOCALAPPDATA%\ASP.NET\DataProtection-Keys\'
            'IdentityConfigurationFilePath'      = 'config\idsconfig.json'
            'TokenStoreDb' = [ordered]@{
                'ConnectionString'     = $Config.DbConnectionString
                'DatabaseProviderType' = 'SqlServer'
            }
            'Features' = [ordered]@{
                'AdminAPI'                        = $true
                'AllowInitializationUI'           = $false
                'UseDeprecatedPasswordApi'        = $false
                'UseNewSamlModulePath'            = $false
                'TrimSpacesFromUsernameAndUserId' = $true
                'AllowWildcardedCorsOrigins'      = $false
                'UseLegacyUserActiveBehavior'     = $false
                'UseLegacyAPIResources'           = $true
                'EnableFederatedLogout'           = $true
                'TriggerExternalSignOut'          = $true
                'UseManagedSigningCert'           = $true
            }
            'ScimClientSettings' = [ordered]@{
                'Proxy' = 'Default'
            }
            'Hyland.Logging' = [ordered]@{
                'Routes' = [ordered]@{
                    'DiagnosticsConsole' = [ordered]@{
                        'Http'          = 'http://localhost:8989'
                        'minimum-level' = 'Error'
                    }
                    'ErrorEventLog' = [ordered]@{
                        'HylandLog'     = 'Hyland Identity Service'
                        'minimum-level' = 'Error'
                    }
                    'SplunkLogger' = [ordered]@{
                        'Splunk'        = '<splunk_collector_url>'
                        'SplunkToken'   = '<splunk_token>'
                        'minimum-level' = 'Error'
                    }
                }
                'DefaultProfile' = 'Hyland.Identity.Service'
            }
            'HealthChecks' = [ordered]@{
                'Enabled' = $true
                'DetailedEndpointAuthorization' = [ordered]@{
                    'EndpointEnabled' = $false
                    'RequireKey'      = $true
                    'StaticKey'       = ''
                }
            }
        }

        Save-CleanJson -FilePath $FilePath -JsonObject $json -Depth 20
        Write-Host "  [OK] appsettings.json patched successfully." -ForegroundColor Green
    }

    # 2. IdS: config\idsconfig.json (Includes AdminUiClient, RedirectUris, CORS fix)
    function Update-IdSIdsConfig {
        param([string]$FilePath, $Config)

        if (-not (Test-Path $FilePath)) { return }
        Write-Host "  Patching $FilePath in-place..." -ForegroundColor Yellow
        $cfgJson = Read-AndCleanJson -FilePath $FilePath

        if ($cfgJson) {
            Set-JsonPropertySafe -TargetObject $cfgJson -PropertyName 'PublicOrigin' -PropertyValue $Config.PublicOrigin
            Set-JsonPropertySafe -TargetObject $cfgJson -PropertyName 'Issuer' -PropertyValue $Config.IdSBaseUrl

            $adminClientId = 'hyland-identity-admin-ui'
            if (-not $cfgJson.PSObject.Properties['AdminUiClient']) {
                $cfgJson | Add-Member -NotePropertyName 'AdminUiClient' -NotePropertyValue ([ordered]@{}) -Force
            }
            Set-JsonPropertySafe -TargetObject $cfgJson.AdminUiClient -PropertyName 'ClientId' -PropertyValue $adminClientId

            if ($cfgJson.Tenants -and $cfgJson.Tenants.Count -gt 0) {
                $tenant = $cfgJson.Tenants[0]
                if ($tenant.Name -eq '<default>' -or [string]::IsNullOrWhiteSpace($tenant.Name)) {
                    Set-JsonPropertySafe -TargetObject $tenant -PropertyName 'Name' -PropertyValue $Config.DataSourceName
                }
                Set-JsonPropertySafe -TargetObject $tenant -PropertyName 'ScimEndpoint' -PropertyValue $Config.ScimEndpoint

                if ($tenant.Clients -and $tenant.Clients.Count -gt 0) {
                    $client = $tenant.Clients[0]
                    Set-JsonPropertySafe -TargetObject $client -PropertyName 'ClientId' -PropertyValue $adminClientId
                    Set-JsonPropertySafe -TargetObject $client -PropertyName 'ClientName' -PropertyValue 'Hyland Identity Admin UI'
                    Set-JsonPropertySafe -TargetObject $client -PropertyName 'Enabled' -PropertyValue $true
                    Set-JsonPropertySafe -TargetObject $client -PropertyName 'AllowAccessTokensViaBrowser' -PropertyValue $true

                    # Correct OIDC Redirect URIs
                    $redirects = @(
                        "$($Config.IdSBaseUrl)/admin/",
                        "$($Config.IdSBaseUrl)/admin/callback.html",
                        "$($Config.IdSBaseUrl)/admin/silent-renew.html",
                        "https://localhost/$($Config.IdSApp)/admin/",
                        "https://localhost/$($Config.IdSApp)/admin/callback.html"
                    )
                    Set-JsonPropertySafe -TargetObject $client -PropertyName 'RedirectUris' -PropertyValue $redirects

                    # Post-Logout Redirects
                    $postLogoutRedirects = @(
                        "$($Config.IdSBaseUrl)/admin/",
                        "https://localhost/$($Config.IdSApp)/admin/"
                    )
                    if ($client.PSObject.Properties['LogoutSettings']) {
                        Set-JsonPropertySafe -TargetObject $client.LogoutSettings -PropertyName 'PostLogoutRedirectUris' -PropertyValue $postLogoutRedirects
                    }

                    # Allowed CORS Origins
                    $corsOrigins = @(
                        $Config.PublicOrigin,
                        'https://localhost',
                        'http://localhost'
                    )
                    if ($client.PSObject.Properties['SecuritySettings']) {
                        Set-JsonPropertySafe -TargetObject $client.SecuritySettings -PropertyName 'AllowedCorsOrigins' -PropertyValue $corsOrigins
                    }
                }
            }

            Save-CleanJson -FilePath $FilePath -JsonObject $cfgJson -Depth 20
            Write-Host "  [OK] idsconfig.json patched with active AdminUiClient, RedirectUris, and CORS." -ForegroundColor Green
        }
    }

    # 3. API Server: config\1_service.json
    function Update-ApiServer1ServerConfig {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$FilePath,
            [Parameter(Mandatory)][object]$Config
        )

        Write-Host "  Patching $FilePath..." -ForegroundColor Cyan

        # 1. Resolve FQDN
        $fqdn = if ($Config.Fqdn) { $Config.Fqdn } else { $env:COMPUTERNAME }

        # 2. Resolve Provider Address (Identity Service)
        $idsApp = if ($Config.IdSApp) { $Config.IdSApp }
                  elseif ($Config.IdSAppName) { $Config.IdSAppName }
                  else { 'identityservice' }
        $idsApp = $idsApp.Trim('/')

        $providerAddress = if ($Config.IdSBaseUrl) {
            $Config.IdSBaseUrl.TrimEnd('/')
        } elseif ($Config.IdSUrl) {
            $Config.IdSUrl.TrimEnd('/')
        } else {
            "https://$fqdn/$idsApp"
        }

        # 3. Resolve TempFileStorage Address (ApiServer TempFS)
        $apiApp = if ($Config.ApiApp) { $Config.ApiApp }
                  elseif ($Config.ApiAppName) { $Config.ApiAppName }
                  else { 'apiserver' }
        $apiApp = $apiApp.Trim('/')

        $tempFsAddress = if ($Config.TempFileStorageAddress) {
            $Config.TempFileStorageAddress
        } elseif ($Config.TempFsUrl) {
            $Config.TempFsUrl
        } else {
            "https://$fqdn/$apiApp/tempfs/"
        }
        if (-not $tempFsAddress.EndsWith('/')) { $tempFsAddress += '/' }

        # 4. Resolve Database Connection String & Data Source
        $dbConn = if ($Config.DbConnectionString) {
            $Config.DbConnectionString
        } elseif ($Config.OnBaseDbConnectionString) {
            $Config.OnBaseDbConnectionString
        } else {
            "Data Source=$($Config.SqlInstance);Database=$($Config.DatabaseName);User Id=$($Config.DbUser);Password=$($Config.DbPassword);Encrypt=Optional;TrustServerCertificate=True;"
        }

        $dataSourceName = if ($Config.DataSourceName) { $Config.DataSourceName }
                         elseif ($Config.DatabaseName) { $Config.DatabaseName }
                         else { "OnBase" }

        # 5. Build clean, explicit JSON object graph
        $targetConfig = [ordered]@{
            "Host" = [ordered]@{
                "OnBase" = [ordered]@{
                    "ConnectionStringProvider" = "Configuration"
                }
                "Authentication" = [ordered]@{
                    "Provider" = [ordered]@{
                        "Address" = $providerAddress
                    }
                }
            }
            "TempFileStorage" = [ordered]@{
                "Address" = $tempFsAddress
            }
            "OnBaseSessionCookie" = [ordered]@{
                "Secure"       = $true
                "SameSiteMode" = "None"
            }
            "Tenant" = @(
                [ordered]@{
                    "DataSource"       = $dataSourceName
                    "DataProvider"     = "SQL"
                    "ConnectionString" = $dbConn
                }
            )
        }

        Write-Host "    [Provider Address]        -> $providerAddress" -ForegroundColor Gray
        Write-Host "    [TempFileStorage Address] -> $tempFsAddress" -ForegroundColor Gray

        # Save and format cleanly
        Save-CleanJson -FilePath $FilePath -JsonObject $targetConfig
    }

    # 4. API Server: config\5_scim.json
    function Update-ApiServer5ScimConfig {
        param([string]$FilePath, $Config)

        Write-Host "  Patching $FilePath with target schema..." -ForegroundColor Yellow
        $json = [ordered]@{
            'Enabled' = $true
            'JwtIssuerOptions' = [ordered]@{
                'Authority' = $Config.IdSBaseUrl
                'Audience'  = "$($Config.IdSBaseUrl)/resources"
            }
            'DataSources' = @(
                [ordered]@{
                    'Name'             = $Config.DataSourceName
                    'ProviderName'     = 'System.Data.SqlClient'
                    'ConnectionString' = $Config.DbConnectionString
                }
            )
        }

        Save-CleanJson -FilePath $FilePath -JsonObject $json -Depth 20
        Write-Host "  [OK] 5_scim.json patched successfully." -ForegroundColor Green
    }

    # 5. API Server: config\5_tempfs.json
    function Update-ApiServer5TempFsConfig {
        param([string]$FilePath, $Config)

        Write-Host "  Patching $FilePath (Preserving full TempFS factory schema)..." -ForegroundColor Yellow

        $tenantEntries = @(
            [ordered]@{
                'DataSource'       = "Tenant_$($Config.DataSourceName)_DDL"
                'DataProvider'     = 'MSSQL'
                'ConnectionString' = $Config.DbConnectionString
            },
            [ordered]@{
                'DataSource'       = "Tenant_$($Config.DataSourceName)_SQL"
                'DataProvider'     = 'MSSQL'
                'ConnectionString' = $Config.DbConnectionString
            }
        )

        $json = [ordered]@{
            'TempFS' = [ordered]@{
                'IDP' = [ordered]@{
                    'Scope' = [ordered]@{
                        'Name' = 'evolution'
                    }
                }
                'TempFileService' = [ordered]@{
                    'Tenant' = $tenantEntries
                    'BackgroundProcessor' = [ordered]@{
                        'CleanupProcessor' = [ordered]@{
                            'Enabled'             = $true
                            'StartupDelaySeconds' = 0
                            'RunDelaySeconds'     = 300
                            'ErrorBehavior'       = 'LogAndContinue'
                        }
                        'CacheFileNotificationProcessor' = [ordered]@{
                            'Enabled'             = $true
                            'StartupDelaySeconds' = 0
                            'RunDelaySeconds'     = 300
                            'ErrorBehavior'       = 'LogAndContinue'
                        }
                    }
                    'Security' = [ordered]@{
                        'MasterPassword'         = 'AQAAANCMnd8BFdERjHoAwE/Cl+sBAAAADlah/mFUuEuiK3jYOys7qAQAAAACAAAAAAADZgAAwAAAABAAAABXeRfOVtTWF4D27qjlv4yDAAAAAASAAACgAAAAEAAAAMBROjFrjE64vt+pUKaEHogIAAAAR5zGR59n2/QUAAAAfz8z1Fc6bIqiTQJxQPbGkFVcFJA='
                        'PasswordIsEncrypted'    = $true
                        'TenantEncryptionScheme' = 'MasterPasswordKEKEK'
                        'KeyVault'               = $null
                    }
                }
                'IdpOptions' = [ordered]@{
                    'IdP'          = ''
                    'Scope'        = ''
                    'ClientSecret' = ''
                    'ClientId'     = ''
                }
                'KeyVaultEndpoint' = ''
                'KeyVaultDatabase' = [ordered]@{
                    'EncryptionDatabase' = ''
                }
            }
        }

        Save-CleanJson -FilePath $FilePath -JsonObject $json -Depth 20
        Write-Host "  [OK] 5_tempfs.json patched successfully." -ForegroundColor Green
    }

    # 6. API Server: Root appsettings.json
    function Update-ApiServerRootAppSettings {
        param([string]$FilePath, $Config)

        $connStrings = [ordered]@{}
        $connStrings[$Config.DataSourceName] = $Config.DbConnectionString

        $json = [ordered]@{
            'IdentityServer'    = [ordered]@{ 'Authority' = $Config.IdSBaseUrl }
            'ConnectionStrings' = $connStrings
        }
        Save-CleanJson -FilePath $FilePath -JsonObject $json -Depth 10
        Write-Host "  [OK] API Server root appsettings.json synchronized." -ForegroundColor Green
    }

    # -------------------------------------------------------------------------
    # Helper: IIS AppPool Lifecycle Management
    # -------------------------------------------------------------------------
    function Restart-HylandWebAppPoolSafely {
        [CmdletBinding()]
        param([string]$AppPoolName)

        Import-Module WebAdministration -ErrorAction SilentlyContinue
        if (Test-Path "IIS:\AppPools\$AppPoolName") {
            try {
                $poolState = (Get-WebAppPoolState -Name $AppPoolName -ErrorAction SilentlyContinue).Value
                if ($poolState -eq 'Stopped') {
                    Write-Host "  AppPool '$AppPoolName' is stopped. Starting it..." -ForegroundColor Cyan
                    Start-WebAppPool -Name $AppPoolName -ErrorAction Stop
                } else {
                    Write-Host "  AppPool '$AppPoolName' is running. Recycling/Restarting it..." -ForegroundColor Cyan
                    Restart-WebAppPool -Name $AppPoolName -ErrorAction Stop
                }
                Start-Sleep -Seconds 2
            } catch {
                Write-Warning "Could not start/restart AppPool '$AppPoolName': $_"
            }
        }
    }

    # -------------------------------------------------------------------------
    # Helper: IIS Native Modules Verification
    # -------------------------------------------------------------------------
    function Ensure-IISInstalledAndModulesConfigured {
        Write-Host "`n[IIS] Checking IIS prerequisites and required modules on $env:COMPUTERNAME..." -ForegroundColor Cyan

        $w3svc = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
        if ($w3svc) {
            Write-Host "  [IIS] IIS (W3SVC) service is present." -ForegroundColor Green
        } else {
            Write-Host "  [IIS] Installing IIS features..." -ForegroundColor Yellow
            Import-Module ServerManager -ErrorAction SilentlyContinue
            if (Get-Command 'Install-WindowsFeature' -ErrorAction SilentlyContinue) {
                $features = @(
                    'Web-Server', 'Web-WebServer', 'Web-Common-Http', 'Web-Default-Doc',
                    'Web-Static-Content', 'Web-Http-Errors', 'Web-Http-Redirect', 'Web-Health',
                    'Web-Http-Logging', 'Web-Log-Libraries', 'Web-Request-Monitor', 'Web-Performance',
                    'Web-Stat-Compression', 'Web-Dyn-Compression', 'Web-Security', 'Web-Filtering',
                    'Web-Windows-Auth', 'Web-Basic-Auth', 'Web-App-Dev', 'Web-Net-Ext45',
                    'Web-Asp-Net45', 'Web-ISAPI-Ext', 'Web-ISAPI-Filter', 'Web-Mgmt-Tools',
                    'Web-Mgmt-Console', 'Web-Scripting-Tools'
                )
                Install-WindowsFeature -Name $features -IncludeManagementTools -ErrorAction Stop | Out-Null
            }
        }

        Set-Service -Name 'W3SVC' -StartupType Automatic -ErrorAction SilentlyContinue
        Start-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
        Import-Module WebAdministration -ErrorAction Stop

        Write-Host "  [IIS] Verifying AnonymousAuthentication, HttpCache, UriCache, and AspNetCore modules..." -ForegroundColor Cyan
        $globalModules = Get-WebGlobalModule

        # 1. UriCacheModule
        $hasUriCache = $globalModules | Where-Object { $_.Name -eq 'UriCacheModule' }
        if (-not $hasUriCache) {
            $uriCacheDll = "$env:windir\System32\inetsrv\uriCache.dll"
            if (Test-Path $uriCacheDll) {
                Write-Host "  Registering UriCacheModule in IIS Global Modules..." -ForegroundColor Yellow
                New-WebGlobalModule -Name 'UriCacheModule' -Image $uriCacheDll | Out-Null
            }
        }

        # 2. HttpCacheModule
        $hasHttpCache = $globalModules | Where-Object { $_.Name -eq 'HttpCacheModule' }
        if (-not $hasHttpCache) {
            $httpCacheDll = "$env:windir\System32\inetsrv\cachhttp.dll"
            if (Test-Path $httpCacheDll) {
                Write-Host "  Registering HttpCacheModule in IIS Global Modules..." -ForegroundColor Yellow
                New-WebGlobalModule -Name 'HttpCacheModule' -Image $httpCacheDll | Out-Null
            }
        }

        # 3. AnonymousAuthenticationModule
        $hasAnon = $globalModules | Where-Object { $_.Name -eq 'AnonymousAuthenticationModule' }
        if (-not $hasAnon) {
            $anonDll = "$env:windir\System32\inetsrv\authanon.dll"
            if (Test-Path $anonDll) {
                New-WebGlobalModule -Name 'AnonymousAuthenticationModule' -Image $anonDll | Out-Null
            }
        }

        Write-Host "  [IIS] IIS Modules verified." -ForegroundColor Green
    }

    # -------------------------------------------------------------------------
    # Helper: Non-Destructive .NET 8.0 Hosting Bundle Check
    # -------------------------------------------------------------------------
    function Ensure-DotNet8HostingBundle {
        param([string]$LocalOrSharePath)

        Write-Host "`n[1/6] Verifying .NET 8.0 Hosting Bundle on $env:COMPUTERNAME..." -ForegroundColor Cyan
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $hasDotNet8 = $false
        if (Get-Command 'dotnet' -ErrorAction SilentlyContinue) {
            $runtimes = & dotnet --list-runtimes 2>$null
            if ($runtimes -match 'Microsoft\.AspNetCore\.App 8\.0\.') { $hasDotNet8 = $true }
        }

        if (-not $hasDotNet8) {
            $net8Reg = Get-ChildItem 'HKLM:\SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.AspNetCore.App' -ErrorAction SilentlyContinue |
                       Where-Object { $_.PSChildName -like '8.0.*' }
            if ($net8Reg) { $hasDotNet8 = $true }
        }

        if (-not $hasDotNet8) {
            $tempDir = 'C:\Temp'
            if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }
            $installerPath = Join-Path $tempDir 'dotnet-hosting-8.0.x-win.exe'

            if ($LocalOrSharePath -and (Test-Path $LocalOrSharePath)) {
                $installerPath = $LocalOrSharePath
            } else {
                $msUrl = 'https://aka.ms/dotnet/8.0/dotnet-hosting-win.exe'
                Write-Host "  Downloading .NET 8.0 Hosting Bundle from Microsoft..." -ForegroundColor Yellow
                Invoke-WebRequest -Uri $msUrl -OutFile $installerPath -UseBasicParsing
            }

            Write-Host "  Installing .NET 8.0 Hosting Bundle silently..." -ForegroundColor Yellow
            $proc = Start-Process -FilePath $installerPath -ArgumentList '/install', '/quiet', '/norestart', '/log', 'C:\Temp\dotnet8_hosting.log' -Wait -PassThru
            Restart-Service -Name W3SVC -Force
            Write-Host "  .NET 8.0 Hosting Bundle installed (ExitCode: $($proc.ExitCode))." -ForegroundColor Green
        } else {
            Write-Host "  .NET 8.0 AspNetCore Runtime verified." -ForegroundColor Green
        }
    }

    # -------------------------------------------------------------------------
    # Helper: Installed Product Version Discovery
    # -------------------------------------------------------------------------
    function Get-InstalledHylandProductVersion {
        param([string[]]$NamePatterns, [string]$FallbackPath)

        $uninstallRoots = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )

        foreach ($root in $uninstallRoots) {
            $items = Get-ItemProperty -Path $root -ErrorAction SilentlyContinue | Where-Object {
                $displayName = $_.DisplayName
                if (-not $displayName) { return $false }
                foreach ($pattern in $NamePatterns) {
                    if ($displayName -like $pattern) { return $true }
                }
                return $false
            }

            $match = $items | Sort-Object DisplayVersion -Descending | Select-Object -First 1
            if ($match) {
                return [PSCustomObject]@{ Found = $true; DisplayName = $match.DisplayName; DisplayVersion = $match.DisplayVersion }
            }
        }

        if ($FallbackPath -and (Test-Path $FallbackPath)) {
            $candidate = Get-ChildItem -Path $FallbackPath -Recurse -Include '*.exe','*.dll' -ErrorAction SilentlyContinue |
                         Where-Object { $_.VersionInfo -and $_.VersionInfo.ProductVersion } |
                         Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($candidate) {
                return [PSCustomObject]@{ Found = $true; DisplayName = $candidate.Name; DisplayVersion = $candidate.VersionInfo.ProductVersion }
            }
        }

        return [PSCustomObject]@{ Found = $false; DisplayName = $null; DisplayVersion = $null }
    }

    # -------------------------------------------------------------------------
    # Helper: ProGet Native Package Endpoint Downloader
    # -------------------------------------------------------------------------
    function Get-RemoteProGetInstaller {
        param(
            [string]$FeedUrl,
            [string]$PackageId,
            [string]$Version,
            [string]$TargetDir = 'C:\Temp\HylandInstallers'
        )

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $feed = $FeedUrl.TrimEnd('/') -replace '/v3/index\.json$', '' -replace '/v3$', ''

        $packageRoot = Join-Path $TargetDir "$PackageId.$Version"
        $nupkgPath = Join-Path $TargetDir "$PackageId.$Version.nupkg"
        $zipPath = [IO.Path]::ChangeExtension($nupkgPath, '.zip')
        $downloadUrl = "$feed/package/$PackageId/$Version"

        if (-not (Test-Path $TargetDir)) { New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null }

        if (-not (Test-Path $packageRoot)) {
            Write-Host "  [ProGet] Fetching $PackageId $Version from $downloadUrl..." -ForegroundColor Yellow
            Invoke-WebRequest -Uri $downloadUrl -OutFile $nupkgPath -UseDefaultCredentials -UseBasicParsing -ErrorAction Stop
            Copy-Item -Path $nupkgPath -Destination $zipPath -Force
            Expand-Archive -Path $zipPath -DestinationPath $packageRoot -Force
            Remove-Item -Path $zipPath, $nupkgPath -Force -ErrorAction SilentlyContinue
            Write-Host "  [ProGet] Extracted to $packageRoot" -ForegroundColor Green
        }

        return $packageRoot
    }

    function Invoke-IdsOperationalStoreMigration {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$SqlInstance,
            [Parameter(Mandatory)][string]$DatabaseName,
            [Parameter(Mandatory)][pscredential]$HsiCredential,
            [string]$InitUtilityDll = 'C:\Program Files\Hyland\Hyland Identity Service Initialization Utility\Hyland.Identity.Server.Initialization.dll',
            [string]$IdsAppSettings = 'C:\Program Files\Hyland\identityservice\config\appsettings.json',
            [ValidateSet('Mandatory','Optional','Strict','True','False')][string]$Encrypt = 'Optional',
            [string]$TargetMachine
        )

        $migrate = {
            param($dll, $appSettings, $sqlInstance, $db, $user, $pwd, $encrypt)

            if (-not (Test-Path $dll)) { throw "IdS Initialization Utility not found: $dll" }
            if (-not (Test-Path $appSettings)) { throw "IdS appsettings.json not found: $appSettings" }
            if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) { throw "dotnet runtime not found on PATH." }

            $cs = "Data Source=$sqlInstance;database=$db;User Id=$user;Password=$pwd;Encrypt=$encrypt"
            $out = & dotnet $dll create-operational-database `
                        --connectionstring $cs `
                        --providertype SqlServer `
                        --appsettings $appSettings 2>&1

            [pscustomobject]@{
                ExitCode = $LASTEXITCODE
                Output = (($out | Out-String) -replace [regex]::Escape($pwd), '********')
            }
        }

        $argsList = @(
            $InitUtilityDll, $IdsAppSettings, $SqlInstance, $DatabaseName,
            $HsiCredential.UserName, $HsiCredential.GetNetworkCredential().Password, $Encrypt
        )

        Write-Host "  Creating/updating IdS operational store (hsiidp) in $DatabaseName..." -ForegroundColor Cyan

        $result = if ($TargetMachine -and $TargetMachine -ne $env:COMPUTERNAME) {
            Invoke-Command -ComputerName $TargetMachine -ScriptBlock $migrate -ArgumentList $argsList
        } else {
            & $migrate @argsList
        }

        #Write-Verbose $result.Output
        Write-Host $result.Output
        if ($result.ExitCode -ne 0) {
            throw "create-operational-database failed (exit $($result.ExitCode)):`n$($result.Output)"
        }
        Write-Host "  [OK] Operational store migrations applied." -ForegroundColor Green
    }
<#>
function Test-IdsOperationalStore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SqlInstance,
        [Parameter(Mandatory)][string]$DatabaseName,
        [Parameter(Mandatory)][pscredential]$HsiCredential
    )

    # Core tables needed by IdS 5.x+. Version-specific tables are reported, not enforced.
    $required = 'devicecodes','persistedgrants','Keys','ServerSideSessions','PushedAuthorizationRequests'

    $pwd = $HsiCredential.GetNetworkCredential().Password
    $conn = New-Object System.Data.SqlClient.SqlConnection(
        "Data Source=$SqlInstance;Database=$DatabaseName;User Id=$($HsiCredential.UserName);Password=$pwd;Encrypt=True;TrustServerCertificate=True")
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT t.name FROM sys.tables t JOIN sys.schemas s ON s.schema_id = t.schema_id WHERE s.name = 'hsiidp'"
        $reader = $cmd.ExecuteReader()
        $found = while ($reader.Read()) { $reader.GetString(0) }
    } finally { $conn.Close() }

    $missing = $required | Where-Object { $found -notcontains $_ }
    Write-Host "  hsiidp tables found: $($found -join ', ')"
    if ($missing) { throw "IdS operational store incomplete. Missing: $($missing -join ', ')" }
    Write-Host "  [OK] IdS operational store validated." -ForegroundColor Green
}
#>
    function Test-DbSchemaTables {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$SqlInstance,
            [Parameter(Mandatory)][string]$DatabaseName,
            [Parameter(Mandatory)][pscredential]$HsiCredential,
            [Parameter(Mandatory)][string]$Schema,
            [Parameter(Mandatory)][string[]]$RequiredTables
        )
        $pwd = $HsiCredential.GetNetworkCredential().Password
        $conn = New-Object System.Data.SqlClient.SqlConnection(
            "Data Source=$SqlInstance;Database=$DatabaseName;User Id=$($HsiCredential.UserName);Password=$pwd;Encrypt=True;TrustServerCertificate=True")
        try {
            $conn.Open()
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT t.name FROM sys.tables t JOIN sys.schemas s ON s.schema_id = t.schema_id WHERE s.name = @schema"
            [void]$cmd.Parameters.AddWithValue('@schema', $Schema)
            $reader = $cmd.ExecuteReader()
            $found = while ($reader.Read()) { $reader.GetString(0) }
        } finally { $conn.Close() }
        $missing = $RequiredTables | Where-Object { $found -notcontains $_ }
        Write-Host "  [$Schema] $(@($found).Count) tables found."
        if ($missing) { throw "Schema '$Schema' incomplete. Missing: $($missing -join ', ')" }
        Write-Host "  [OK] Schema '$Schema' validated." -ForegroundColor Green
    }

    function Invoke-TempFsDatabaseUpdate {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$TempFsConfigExe,   # path to TempFileService.Configuration.exe
            [string]$TempFsJson = 'C:\Program Files\Hyland\ApiServer\config\5_tempfs.json',
            [string]$TargetMachine
        )
        $update = {
            param($exe, $json)
            if (-not (Test-Path $exe)) { throw "TempFileService.Configuration.exe not found: $exe" }
            if (-not (Test-Path $json)) { throw "5_tempfs.json not found: $json" }
            $out = 'y' | & $exe database update --json-config $json 2>&1
            #$out = 'y' | & $TempFsConfigExe database update --json-config $TempFsJson 2>&1
            [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = ($out | Out-String) }
        }
        Write-Host "  Creating/updating TempFS schema (hsitempfile)..." -ForegroundColor Cyan
        $result = if ($TargetMachine -and $TargetMachine -ne $env:COMPUTERNAME) {
            Invoke-Command -ComputerName $TargetMachine -ScriptBlock $update -ArgumentList $TempFsConfigExe, $TempFsJson
        } else {
            & $update $TempFsConfigExe $TempFsJson
        }
        # Success prints nothing. Errors (including the permission pre-check) print to console / exit 1.
        if ($result.ExitCode -ne 0 -or $result.Output -match 'error|exception|INSUFFICIENT') {
            throw "TempFS 'database update' failed (exit $($result.ExitCode)):`n$($result.Output)"
        }
        Write-Host "  [OK] TempFS migrations applied." -ForegroundColor Green
    }

    function Set-IdsTenantSecurity {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$IdsConfigPath,
            [string]$TenantName = 'OnBase',
            [string[]]$AdminUsers,
            [string[]]$AdminGroups,
            [string]$ProvidersJson
        )

        # Recursively flattens any nested collections, arrays, or {value,Count} wrappers down to leaf objects
        function Get-FlatObjects($inputObject) {
            if ($null -eq $inputObject) { return }
            if ($inputObject -is [System.Collections.IEnumerable] -and $inputObject -isnot [string]) {
                foreach ($item in $inputObject) {
                    Get-FlatObjects $item
                }
            }
            elseif ($inputObject.PSObject.Properties['value'] -and $inputObject.PSObject.Properties['Count']) {
                Get-FlatObjects $inputObject.value
            }
            else {
                $inputObject
            }
        }

        $cfg = Get-Content $IdsConfigPath -Raw | ConvertFrom-Json
        $tn = $cfg.Tenants | Where-Object { $_.Name -eq $TenantName } | Select-Object -First 1
        if (-not $tn) { throw "Tenant '$TenantName' not found in $IdsConfigPath" }

        # Admin users/groups: flatten and deduplicate
        $existingUsers = @(Get-FlatObjects $tn.AdministrativeUsers)
        $existingGroups = @(Get-FlatObjects $tn.AdministrativeGroups)
        $users = @($existingUsers + @($AdminUsers) | Where-Object { $_ } | Select-Object -Unique)
        $groups = @($existingGroups + @($AdminGroups) | Where-Object { $_ } | Select-Object -Unique)
        $tn.AdministrativeUsers = [string[]]$users
        $tn.AdministrativeGroups = [string[]]$groups

        <# Providers: flatten existing + incoming, index strictly by Name as single objects
        $byName = [ordered]@{}
        foreach ($p in @(Get-FlatObjects $tn.Providers)) {
            if ($p -and $p.PSObject.Properties['Name']) {
                $byName[[string]$p.Name] = $p
            }
        }

        if ($ProvidersJson) {
            $parsedIncoming = ConvertFrom-Json -InputObject $ProvidersJson
            foreach ($p in @(Get-FlatObjects $parsedIncoming)) {
                if (-not $p -or -not $p.PSObject.Properties['Name']) { continue }
                $name = [string]$p.Name
                $id = if ($byName.Contains($name) -and $byName[$name].Id) { $byName[$name].Id }
                      elseif ($p.PSObject.Properties['Id'] -and $p.Id) { $p.Id }
                      else { [guid]::NewGuid().ToString() }

                if ($p.PSObject.Properties['Id']) { $p.Id = $id }
                else { Add-Member -InputObject $p -NotePropertyName Id -NotePropertyValue $id }

                $byName[$name] = $p
                Write-Host "  Provider '$name' set (Id $id)"
            }
        }#>#-----------------------------------------------

        Write-Host "  [DEBUG] ProvidersJson: $(if ($ProvidersJson) { "$($ProvidersJson.Length) chars" } else { 'EMPTY' })" -ForegroundColor Magenta
        Write-Host "  [DEBUG] Existing providers: $((@(Get-FlatObjects $tn.Providers) | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor Magenta
        if ($ProvidersJson) {
            Write-Host "  [DEBUG] Incoming providers: $((@(Get-FlatObjects (ConvertFrom-Json -InputObject $ProvidersJson)) | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor Magenta
        }#----------------------------------------------

        # Providers: Settings.json is the source of truth; existing entries only supply Ids
        $existingIds = @{}
        foreach ($p in @(Get-FlatObjects $tn.Providers)) {
            if ($p.Name -and $p.Id) { $existingIds[[string]$p.Name] = [string]$p.Id }
        }

        $byName = [ordered]@{}
        if ($ProvidersJson) {
            $parsedIncoming = ConvertFrom-Json -InputObject $ProvidersJson
            foreach ($p in @(Get-FlatObjects $parsedIncoming)) {
                if (-not $p.Name) { continue }
                $name = [string]$p.Name
                $id = if ($existingIds.ContainsKey($name)) { $existingIds[$name] } else { [guid]::NewGuid().ToString() }
                $p | Add-Member -NotePropertyName Id -NotePropertyValue $id -Force
                $byName[$name] = $p
            }
        } else {
            # No providers passed in: leave the existing ones alone instead of wiping them
            foreach ($p in @(Get-FlatObjects $tn.Providers)) { if ($p.Name) { $byName[[string]$p.Name] = $p } }
        }

        $removed = @($existingIds.Keys | Where-Object { -not $byName.Contains($_) })
        if ($removed.Count) { Write-Host "  [INFO] Removing providers not in Settings.json: $($removed -join ', ')" -ForegroundColor Yellow }

        $parts = foreach ($v in $byName.Values) { ConvertTo-Json -InputObject $v -Depth 20 -Compress }

        # Serialize each provider individually, stripping any accidental outer brackets
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($name in $byName.Keys) {
            $rawObjJson = (ConvertTo-Json -InputObject $byName[$name] -Depth 50).Trim()
            if ($rawObjJson.StartsWith('[') -and $rawObjJson.EndsWith(']')) {
                $rawObjJson = $rawObjJson.Substring(1, $rawObjJson.Length - 2).Trim()
            }
            if ($rawObjJson) {
                $parts.Add($rawObjJson)
            }
        }

        $providersText = '[' + ($parts -join ',') + ']'
        $token = '__IDS_PROVIDERS_PLACEHOLDER__'
        $tn.Providers = $token

        if ($tn.Id -eq '\u003cdefault\u003e') { $tn.Id = '<default>' }

        $json = $cfg | ConvertTo-Json -Depth 100
        $json = $json.Replace('"' + $token + '"', $providersText)

        # Safety validation
        if ($json -match '"Providers"\s*:\s*\[\s*\[') {
            throw 'Providers would be nested - idsconfig.json NOT saved.'
        }

        Save-JsonFile -Path $IdsConfigPath -Json $json

        Write-Host "  [OK] Admin users: $($users -join ', ')" -ForegroundColor Green
        Write-Host "  [OK] Admin groups: $($groups -join ', ')" -ForegroundColor Green
        Write-Host "  [OK] Providers: $($byName.Count)" -ForegroundColor Green
    }

    # =========================================================================
    # Step-by-Step Execution Flow
    # =========================================================================
    $hostName = $env:COMPUTERNAME.ToLower()
    $domain = (Get-CimInstance Win32_ComputerSystem).Domain.ToLower()
    $fqdn = if ($domain -and $domain -ne 'workgroup') { "$hostName.$domain" } else { $hostName }

    Ensure-IISInstalledAndModulesConfigured
    Ensure-DotNet8HostingBundle -LocalOrSharePath $DotNetSrc

    # -------------------------------------------------------------------------
    # 2. SQL TCP/IP Protocol Check
    # -------------------------------------------------------------------------
    Write-Host "`n[2/6] Verifying SQL Server TCP/IP Protocol..." -ForegroundColor Cyan
    $instPart = if ($SqlInst -like '*\*') { $SqlInst.Split('\')[1] } else { 'MSSQLSERVER' }
    $tcpRegs = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like "*$instPart*" }
    foreach ($reg in $tcpRegs) {
        $tcpPath = Join-Path $reg.PSPath 'MSSQLServer\SuperSocketNetLib\Tcp'
        if (Test-Path $tcpPath) { Set-ItemProperty -Path $tcpPath -Name 'Enabled' -Value 1 -Force }
    }
    $sqlSvcName = if ($instPart -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$instPart" }
    Restart-Service -Name $sqlSvcName -Force
    Write-Host "  TCP/IP verified on instance '$SqlInst'." -ForegroundColor Green
    <#
    # -------------------------------------------------------------------------
    # 3. SSL Certificate & IIS 443 Binding + Private Key ACLs for IIS_IUSRS
    # -------------------------------------------------------------------------
    Write-Host "`n[3/6] Configuring SSL Certificate and IIS 443 Binding..." -ForegroundColor Cyan
    $cert = $null
    if ($CertThumb) { $cert = Get-Item "Cert:\LocalMachine\My\$CertThumb" -ErrorAction SilentlyContinue }
    if (-not $cert) {
        $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -like "*$hostName*" -or $_.Subject -like "*$fqdn*" } | Select-Object -First 1
    }
    if (-not $cert) {
        $cert = New-SelfSignedCertificate -DnsName $fqdn, $hostName, 'localhost' `
                                          -CertStoreLocation 'Cert:\LocalMachine\My' `
                                          -KeyUsage DigitalSignature, KeyEncipherment `
                                          -Type SSLServerAuthentication `
                                          -FriendlyName "Hyland_IdS_$hostName"
        $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'LocalMachine')
        $rootStore.Open('ReadWrite')
        $rootStore.Add($cert)
        $rootStore.Close()
    }
    $thumbprint = $cert.Thumbprint
    Write-Host "  Using Certificate Thumbprint: $thumbprint" -ForegroundColor Green

    try {
        $rsaCert = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        if ($rsaCert -and $rsaCert.key -and $rsaCert.key.UniqueKeyContainerName) {
            $keyFileName = $rsaCert.key.UniqueKeyContainerName
            $keyFilePath = "$env:ALLUSERSPROFILE\Microsoft\Crypto\RSA\MachineKeys\$keyFileName"
            if (Test-Path $keyFilePath) {
                Write-Host "  Setting Read permission for IIS_IUSRS on Private Key container: $keyFileName" -ForegroundColor Cyan
                $keyAcl = Get-Acl -Path $keyFilePath
                $keyRule = New-Object System.Security.AccessControl.FileSystemAccessRule('IIS_IUSRS', 'Read', 'Allow')
                $keyAcl.AddAccessRule($keyRule)
                Set-Acl -Path $keyFilePath -AclObject $keyAcl
            }
        }
    } catch {
        Write-Warning "Could not directly bind private key ACL: $_"
    }

    $binding = Get-WebBinding -Name 'Default Web Site' -Protocol 'https' -Port 443 -ErrorAction SilentlyContinue
    if (-not $binding) { New-WebBinding -Name 'Default Web Site' -Protocol 'https' -Port 443 -IPAddress '*' -SslFlags 0 }
    @('IIS:\SslBindings\0.0.0.0!443', 'IIS:\SslBindings\*!443') | ForEach-Object { if (Test-Path $_) { Remove-Item -Path $_ -Force -ErrorAction SilentlyContinue } }
    Get-Item "Cert:\LocalMachine\My\$thumbprint" | New-Item -Path 'IIS:\SslBindings\0.0.0.0!443' -Force | Out-Null
    #>
    # -------------------------------------------------------------------------
    # 3. SSL Certificate Lifecycle (FriendlyName Reuse -> Create Only If Expired)
    # -------------------------------------------------------------------------
    Write-Host "`n[3/6] Configuring SSL Certificate (Reusability & Expiration Check)..." -ForegroundColor Cyan

    $certFriendlyName = "Hyland_Integration_$hostName"
    $cert = $null

    # 1. Explicit Thumbprint Parameter Check
    if ($CertThumb) {
        $candidate = Get-Item "Cert:\LocalMachine\My\$CertThumb" -ErrorAction SilentlyContinue
        if ($candidate -and $candidate.NotAfter -gt (Get-Date)) {
            $cert = $candidate
            Write-Host "  [REUSE] Using supplied valid certificate: $($cert.Thumbprint)" -ForegroundColor Green
        } else {
            Write-Warning "  [WARN] Supplied CertificateThumbprint ($CertThumb) is missing or expired. Searching for existing valid certificate..."
        }
    }

    # 2. Search for existing, unexpired certificate by FriendlyName
    if (-not $cert) {
        $cert = Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
            Where-Object {
                ($_.FriendlyName -eq $certFriendlyName -or $_.FriendlyName -eq "Hyland_IdS_$hostName") -and
                $_.NotAfter -gt (Get-Date)
            } |
            Sort-Object NotAfter -Descending |
            Select-Object -First 1
    }

    # 3. Fallback: Search for existing, unexpired certificate by Subject
    if (-not $cert) {
        $cert = Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
            Where-Object {
                ($_.Subject -like "*$fqdn*" -or $_.Subject -like "*$hostName*") -and
                $_.NotAfter -gt (Get-Date)
            } |
            Sort-Object NotAfter -Descending |
            Select-Object -First 1
    }

    # 4. Decision: Reuse Existing or Generate New 5-Year Certificate
    if ($cert) {
        $daysRemaining = ($cert.NotAfter - (Get-Date)).Days
        Write-Host "  [REUSE] Found existing valid certificate: $($cert.Thumbprint)" -ForegroundColor Green
        Write-Host "          Subject       : $($cert.Subject)" -ForegroundColor Gray
        Write-Host "          Friendly Name : $($cert.FriendlyName)" -ForegroundColor Gray
        Write-Host "          Expires On    : $($cert.NotAfter.ToShortDateString()) ($daysRemaining days remaining)" -ForegroundColor Gray
    } else {
        Write-Host "  [CREATE] No valid certificate found. Generating new 5-year certificate for $fqdn..." -ForegroundColor Yellow
        $cert = New-SelfSignedCertificate `
            -DnsName $fqdn, $hostName, 'localhost', '127.0.0.1' `
            -CertStoreLocation 'Cert:\LocalMachine\My' `
            -NotAfter (Get-Date).AddYears(5) `
            -KeyUsage DigitalSignature, KeyEncipherment, KeyAgreement `
            -Type SSLServerAuthentication `
            -FriendlyName $certFriendlyName

        Write-Host "  [OK] Generated new certificate: $($cert.Thumbprint) (Valid until: $($cert.NotAfter.ToShortDateString()))" -ForegroundColor Green
    }

    $thumbprint = ($cert.Thumbprint -replace '\s', '').ToUpperInvariant()

    # 5. Self-Healing: Ensure Certificate is in Trusted Root (Runs for BOTH Reused and New certs)
    $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'LocalMachine')
    $rootStore.Open('ReadWrite')
    if (-not ($rootStore.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint })) {
        $rootStore.Add($cert)
        Write-Host "  [OK] Synchronized certificate to LocalMachine\Root (Trusted Root Store)." -ForegroundColor Green
    } else {
        Write-Host "  [OK] Certificate already present in LocalMachine\Root." -ForegroundColor Gray
    }
    $rootStore.Close()

    Write-Host "  Using Certificate Thumbprint: $thumbprint" -ForegroundColor Green

    # 3A. Grant IIS_IUSRS read permissions on Certificate Private Key Container
    try {
        $rsaCert = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        if ($rsaCert -and $rsaCert.key -and $rsaCert.key.UniqueKeyContainerName) {
            $keyFileName = $rsaCert.key.UniqueKeyContainerName
            $keyFilePath = "$env:ALLUSERSPROFILE\Microsoft\Crypto\RSA\MachineKeys\$keyFileName"
            if (Test-Path $keyFilePath) {
                Write-Host "  Setting Read permission for IIS_IUSRS on Private Key container..." -ForegroundColor Cyan
                $keyAcl = Get-Acl -Path $keyFilePath
                $keyRule = New-Object System.Security.AccessControl.FileSystemAccessRule('IIS_IUSRS', 'Read', 'Allow')
                $keyAcl.AddAccessRule($keyRule)
                Set-Acl -Path $keyFilePath -AclObject $keyAcl
            }
        }
    } catch {
        Write-Warning "  Could not directly bind private key ACL: $_"
    }

    <#
    # 3B. Remove Hostname/SNI-locked bindings and establish clean Wildcard (*:443:) binding
    Import-Module WebAdministration -ErrorAction SilentlyContinue

    Get-WebBinding -Name 'Default Web Site' -Protocol 'https' -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Remove-WebBinding -Name 'Default Web Site' -Protocol 'https' -BindingInformation $_.bindingInformation -ErrorAction SilentlyContinue
        } catch {}
    }

    if (Test-Path 'IIS:\SslBindings\0.0.0.0!443') {
        Remove-Item -Path 'IIS:\SslBindings\0.0.0.0!443' -Force -ErrorAction SilentlyContinue
    }

    # Rebind wildcard 443 with SslFlags=0 (No SNI lock)
    New-WebBinding -Name 'Default Web Site' -IPAddress '*' -Port 443 -Protocol 'https' -SslFlags 0 | Out-Null
    Get-Item "Cert:\LocalMachine\My\$thumbprint" | New-Item -Path 'IIS:\SslBindings\0.0.0.0!443' -Force | Out-Null

    #>
    # -------------------------------------------------------------------------
    # 3B. Clean IIS 443 Bindings & SSL Certificate Association
    # -------------------------------------------------------------------------
    Import-Module WebAdministration -ErrorAction SilentlyContinue

    # 1. Remove all existing HTTPS bindings on Default Web Site
    Get-WebBinding -Name 'Default Web Site' -Protocol 'https' -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Remove-WebBinding -Name 'Default Web Site' -BindingInformation $_.bindingInformation -ErrorAction SilentlyContinue
        } catch {}
    }

    # 2. Add clean HTTPS bindings
    # Primary FQDN binding (makes IIS Manager display <redacted URL><your server name>.hylandqa.net)
    New-WebBinding -Name 'Default Web Site' -IPAddress '*' -Port 443 -HostHeader $fqdn -Protocol 'https' -SslFlags 0 | Out-Null

    # Secondary Wildcard binding (ensures localhost loopback / IP calls still resolve)
    New-WebBinding -Name 'Default Web Site' -IPAddress '*' -Port 443 -HostHeader '' -Protocol 'https' -SslFlags 0 | Out-Null

    # 3. Safely delete existing HTTP.sys SSL cert registration before rebinding
    netsh http delete sslcert ipport=0.0.0.0:443 2>&1 | Out-Null
    if (Test-Path 'IIS:\SslBindings\0.0.0.0!443') {
        Remove-Item -Path 'IIS:\SslBindings\0.0.0.0!443' -Force -ErrorAction SilentlyContinue
    }

    # 4. Re-bind SSL Certificate cleanly
    $appId = '{4dc3e181-e14b-4a21-b022-59fc669b0914}'
    $netshAdd = netsh http add sslcert ipport=0.0.0.0:443 certhash=$thumbprint appid=$appId 2>&1
    if ($LASTEXITCODE -ne 0 -and $netshAdd -notmatch 'already exists') {
        # Fallback to PowerShell provider if netsh reported unexpected error
        Get-Item "Cert:\LocalMachine\My\$thumbprint" | New-Item -Path 'IIS:\SslBindings\0.0.0.0!443' -Force -ErrorAction SilentlyContinue | Out-Null
    }

    Restart-Service -Name W3SVC -Force
    Write-Host "  [OK] IIS Port 443 bound to $fqdn and 0.0.0.0:443." -ForegroundColor Green

    #Restart-Service -Name W3SVC -Force
    #Write-Host "  [OK] IIS Port 443 wildcard binding bound to 0.0.0.0:443 and W3SVC refreshed." -ForegroundColor Green

    # -------------------------------------------------------------------------
    # 4. Version & Directory-Aware Installation
    # -------------------------------------------------------------------------
    Write-Host "`n[4/6] Evaluating Installed Versions & Directory State..." -ForegroundColor Cyan
    $idpDestDir = "C:\Program Files\Hyland\$IdSApp"
    $shortIdpDestDir = "C:\PROGRA~1\Hyland\$IdSApp"
    $apiDestDir = "C:\Program Files\Hyland\$ApiApp"

    $idpInstalled = Get-InstalledHylandProductVersion -NamePatterns @('*Hyland*Identity*', '*Identity*Provider*', '*Identity*Service*', '*IdS*') -FallbackPath $idpDestDir
    $apiInstalled = Get-InstalledHylandProductVersion -NamePatterns @('*Hyland*API*Server*', '*ApiServer*') -FallbackPath $apiDestDir

    $idpFolderMissing = (-not (Test-Path $idpDestDir)) -or ((Get-ChildItem -Path $idpDestDir -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0)
    $apiFolderMissing = (-not (Test-Path $apiDestDir)) -or ((Get-ChildItem -Path $apiDestDir -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0)

    $installIdS = (-not $idpInstalled.Found) -or ($idpInstalled.DisplayVersion -ne $IdSVer) -or $idpFolderMissing -or $Force
    $installApi = (-not $apiInstalled.Found) -or ($apiInstalled.DisplayVersion -ne $ApiVer) -or $apiFolderMissing -or $Force

    # 4A. Identity Service Installation
    if ($installIdS) {
        $packageFolderOrFile = if ($IdSSrc) { $IdSSrc } else {
            Get-RemoteProGetInstaller -FeedUrl $ProGetFeed -PackageId $IdSPackage -Version $IdSVer
        }

        Import-Module WebAdministration -ErrorAction SilentlyContinue
        foreach ($poolName in @("$($IdSApp)AppPool", "$($ApiApp)AppPool", 'identityproviderAppPool', 'identityserviceAppPool')) {
            if (Test-Path "IIS:\AppPools\$poolName") {
                try {
                    $state = (Get-WebAppPoolState -Name $poolName -ErrorAction SilentlyContinue).Value
                    if ($state -and $state -ne 'Stopped') {
                        Write-Host "  Stopping AppPool '$poolName'..." -ForegroundColor Cyan
                        Stop-WebAppPool -Name $poolName -ErrorAction SilentlyContinue
                        $timeout = 10
                        while ((Get-WebAppPoolState -Name $poolName -ErrorAction SilentlyContinue).Value -ne 'Stopped' -and $timeout -gt 0) {
                            Start-Sleep -Seconds 1
                            $timeout--
                        }
                    }
                } catch {}
            }
        }

        $installerPath = if (Test-Path $packageFolderOrFile -PathType Container) {
            (Get-ChildItem -Path $packageFolderOrFile -Recurse -Include '*.exe','*.msi' |
             Where-Object { $_.Name -notmatch 'nuget|uninstall' } |
             Sort-Object Length -Descending | Select-Object -First 1).FullName
        } else {
            $packageFolderOrFile
        }

        if (-not $installerPath) { throw "No installer found in $packageFolderOrFile" }
        if (-not (Test-Path $idpDestDir)) { New-Item -ItemType Directory -Path $idpDestDir -Force | Out-Null }

        if ($installerPath -like '*.exe') {
            Write-Host "  Installing Identity Service (EXE): $installerPath..." -ForegroundColor Yellow
            $exeArgs = @(
                '/quiet',
                '/norestart',
                '/log', 'C:\Temp\IdS_install.log',
                "APPNAME=$IdSApp",
                "idpAppName=$IdSApp",
                "APPPOOL=$($IdSApp)AppPool",
                "idpAppPool=$($IdSApp)AppPool",
                "APIROOT=$shortIdpDestDir",
                "idpInstallDir=$shortIdpDestDir",
                'APPPOOL_USER=NT AUTHORITY\Network Service',
                'idpAppPoolUser=NT AUTHORITY\Network Service',
                'RETAINCONFIGFILES=yes',
                'retainConfigFiles=yes'
            )
            $p = Start-Process -FilePath $installerPath -ArgumentList $exeArgs -Wait -PassThru
        } else {
            Write-Host "  Installing Identity Service (MSI): $installerPath..." -ForegroundColor Yellow
            $msiArgs = @(
                '/i', "`"$installerPath`"",
                '/qn',
                '/norestart',
                '/l*v', 'C:\Temp\IdS_install.log',
                "APPNAME=$IdSApp",
                "APPPOOL=$($IdSApp)AppPool",
                'WEB_SITE_NAME="Default Web Site"',
                "APIROOT=`"$idpDestDir`"",
                'USE_APPPOOL_IDENTITY=1',
                'DOTNETBUNDLEINSTALLED=1',
                'CheckForDotNetHostBundle=0'
            )
            $p = Start-Process -FilePath 'msiexec.exe' -ArgumentList $msiArgs -Wait -PassThru
        }

        if ($p.ExitCode -notin 0, 3010) {
            throw "IdS installation failed with exit code $($p.ExitCode). Check C:\Temp\IdS_install.log"
        }
        Write-Host "  Identity Service installed successfully." -ForegroundColor Green
    } else {
        Write-Host "  [SKIP] Identity Service '$($idpInstalled.DisplayVersion)' is already current." -ForegroundColor Green
    }

    # 4B. API Server Installation (Passing WiX arguments & SKIPCONFIGUREIIS=1)
    if ($installApi) {
        $resolvedApiSrc = if ($ApiSrc) { $ApiSrc } else {
            $apiPkg = Get-RemoteProGetInstaller -FeedUrl $ProGetFeed -PackageId $ApiPackage -Version $ApiVer
            (Get-ChildItem -Path $apiPkg -Recurse -Include '*.msi','*.exe' |
             Where-Object { $_.Name -notmatch 'nuget|uninstall' } |
             Sort-Object Length -Descending | Select-Object -First 1).FullName
        }

        Write-Host "  Installing API Server from $resolvedApiSrc..." -ForegroundColor Yellow
        if ($resolvedApiSrc -like '*.msi') {
            $msiArgs = @(
                '/i', "`"$resolvedApiSrc`"",
                '/qn',
                '/norestart',
                '/l*v', 'C:\Temp\API_install.log',
                "APIROOT=`"$apiDestDir`"",
                'SKIPCONFIGUREIIS=1'
            )
            $p = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru
        } else {
            $apiArgs = @(
                '/quiet',
                '/norestart',
                '/log', 'C:\Temp\API_install.log',
                "APPNAME=$ApiApp",
                "APPPOOL=$($ApiApp)AppPool",
                'WEB_SITE_NAME="Default Web Site"',
                "APIROOT=C:\PROGRA~1\Hyland\$ApiApp",
                'APPPOOL_USER=NT AUTHORITY\Network Service',
                'SKIPCONFIGUREIIS=1',
                'BYPASS_CPP_2010=1',
                'BYPASS_CPP_2012=1'
            )
            $p = Start-Process -FilePath $resolvedApiSrc -ArgumentList $apiArgs -Wait -PassThru
        }

        if ($p.ExitCode -notin 0, 3010) {
            throw "API Server installation failed with exit code $($p.ExitCode). Check C:\Temp\API_install.log"
        }
        Write-Host "  API Server installed successfully." -ForegroundColor Green
    } else {
        Write-Host "  [SKIP] API Server '$($apiInstalled.DisplayVersion)' is already current." -ForegroundColor Green
    }

    # SEQUENCE FIX: moved the existing pool-configuration loop before the ACL block.
    # Its contents and existing-pool reconciliation behavior are unchanged.
    # Configure IIS AppPools & Web Applications
    foreach ($pool in @("$($IdSApp)AppPool", "$($ApiApp)AppPool")) {
        if (-not (Test-Path "IIS:\AppPools\$pool")) {
            $np = New-Item "IIS:\AppPools\$pool"
            $np.managedRuntimeVersion = ''
            $np.processModel.identityType = 'ApplicationPoolIdentity'
            $np | Set-Item
        } else {
            Set-ItemProperty "IIS:\AppPools\$pool" -Name 'managedRuntimeVersion' -Value ''
            Set-ItemProperty "IIS:\AppPools\$pool" -Name 'processModel.identityType' -Value 'ApplicationPoolIdentity'
        }
    }

    # --------------------------------------------------------------------------
    # Ensure Write/Modify Permissions on ApiServer and IdS Directories
    # --------------------------------------------------------------------------
    Write-Host "  Configuring directory permissions for ApiServer and IdS..." -ForegroundColor Cyan

    Import-Module WebAdministration -ErrorAction SilentlyContinue

    # 1. Resolve physical paths directly on the target machine (or fallback to defaults)
    $apiPath = (Get-WebApplication -Site 'Default Web Site' | Where-Object { $_.Path -match $ApiApp }).PhysicalPath
    if (-not $apiPath) { $apiPath = $apiDestDir }

    $idsPath = (Get-WebApplication -Site 'Default Web Site' | Where-Object { $_.Path -match $IdSApp }).PhysicalPath
    if (-not $idsPath) { $idsPath = $idpDestDir }

    $apiAppPool = "$($ApiApp)AppPool"
    $idsAppPool = "$($IdSApp)AppPool"

    # 2. Build list of target folders (Root, Config, Logs, and TempFS)
    $foldersToConfigure = @(
        @{ Path = $apiPath;                             AppPool = $apiAppPool },
        @{ Path = (Join-Path $apiPath 'config');         AppPool = $apiAppPool },
        @{ Path = (Join-Path $apiPath 'logs');           AppPool = $apiAppPool },
        @{ Path = $idsPath;                             AppPool = $idsAppPool },
        @{ Path = (Join-Path $idsPath 'config');         AppPool = $idsAppPool },
        @{ Path = (Join-Path $idsPath 'logs');           AppPool = $idsAppPool },
        @{ Path = 'C:\ProgramData\Hyland\TempFS';        AppPool = $apiAppPool }
    )

    $inherit = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
    $prop = [System.Security.AccessControl.PropagationFlags]'None'

    foreach ($entry in $foldersToConfigure) {
        if (-not (Test-Path $entry.Path)) {
            New-Item -ItemType Directory -Path $entry.Path -Force | Out-Null
        }

        $acl = Get-Acl $entry.Path

        # BUILTIN\Administrators & SYSTEM (Full Control)
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule('BUILTIN\Administrators', 'FullControl', $inherit, $prop, 'Allow')))
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule('NT AUTHORITY\SYSTEM', 'FullControl', $inherit, $prop, 'Allow')))

        # BUILTIN\IIS_IUSRS (Full Control / Write)
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule('BUILTIN\IIS_IUSRS', 'FullControl', $inherit, $prop, 'Allow')))

        # Specific Application Pool Identity (Modify)
        if ($entry.AppPool) {
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("IIS AppPool\$($entry.AppPool)", 'Modify', $inherit, $prop, 'Allow')))
        }

        Set-Acl -Path $entry.Path -AclObject $acl
    }

    Write-Host "  Directory permissions applied successfully." -ForegroundColor Green

    if (-not (Test-Path $idpDestDir)) { New-Item -ItemType Directory -Path $idpDestDir -Force | Out-Null }
    if (-not (Test-Path $apiDestDir)) { New-Item -ItemType Directory -Path $apiDestDir -Force | Out-Null }

    if (-not (Test-Path "IIS:\Sites\Default Web Site\$IdSApp")) {
        New-WebApplication -Name $IdSApp -Site 'Default Web Site' -PhysicalPath $idpDestDir -ApplicationPool "$($IdSApp)AppPool" | Out-Null
    }
    if (-not (Test-Path "IIS:\Sites\Default Web Site\$ApiApp")) {
        New-WebApplication -Name $ApiApp -Site 'Default Web Site' -PhysicalPath $apiDestDir -ApplicationPool "$($ApiApp)AppPool" | Out-Null
    }

    # -------------------------------------------------------------------------
    # 5. Parameterized Schema Reconciliation Architecture
    # -------------------------------------------------------------------------

    Write-Host "`n[5/6] Reconciling Configuration Files via Modular Parameter Engine..." -ForegroundColor Cyan

    $idpConfigDir = Join-Path $idpDestDir 'config'
    if (-not (Test-Path $idpConfigDir)) { New-Item -ItemType Directory -Path $idpConfigDir -Force | Out-Null }

    $apiConfigDir = Join-Path $apiDestDir 'config'
    if (-not (Test-Path $apiConfigDir)) { New-Item -ItemType Directory -Path $apiConfigDir -Force | Out-Null }

    # Connection string standard
    $password = $HsiCred.GetNetworkCredential().Password
    $dbConnStr = "Data Source=$SqlInst;Database=$DbName;User Id=$($HsiCred.UserName);Password=$password;Encrypt=True;TrustServerCertificate=True;"

    # Assemble Centralized Patch Arguments Object
    $ConfigPatchArgs = [PSCustomObject]@{
        Fqdn                   = $fqdn
        HostNameUpper          = $hostName.ToUpper()
        IdSApp                 = $IdSApp
        ApiApp                 = $ApiApp
        IdSBaseUrl             = "https://$fqdn/$IdSApp"
        ApiBaseUrl             = "https://$fqdn/$ApiApp"
        PublicOrigin           = "https://$fqdn"
        ScimEndpoint           = "https://$fqdn/$ApiApp/onbase/scim"
        TempFileStorageAddress = "https://$fqdn/$ApiApp/tempfs/"
        SqlInstance            = $SqlInst
        DatabaseName           = $DbName
        DataSourceName         = $DbName
        DbUser                 = $HsiCred.UserName
        DbConnectionString     = $dbConnStr
        CertificateThumbprint  = $thumbprint
    }

    # Execute Dedicated Patch Functions with $Config
    $idsAppFile = Join-Path $idpConfigDir 'appsettings.json'
    $idsConfigFile = Join-Path $idpConfigDir 'idsconfig.json'
    $svcFile = Join-Path $apiConfigDir '1_server.json'
    $scimFile = Join-Path $apiConfigDir '5_scim.json'
    $tempfsFile = Join-Path $apiConfigDir '5_tempfs.json'
    $rootAppFile = Join-Path $apiDestDir 'appsettings.json'

    Update-IdSAppSettingsConfig     -FilePath $idsAppFile    -Config $ConfigPatchArgs
    Update-IdSIdsConfig             -FilePath $idsConfigFile -Config $ConfigPatchArgs
    Set-IdsTenantSecurity -IdsConfigPath 'C:\Program Files\Hyland\identityservice\config\idsconfig.json' `
        -AdminUsers $AdminUsers -AdminGroups $AdminGroups -ProvidersJson $ProvidersJson
    Update-ApiServer1ServerConfig  -FilePath $svcFile       -Config $ConfigPatchArgs
    Update-ApiServer5ScimConfig     -FilePath $scimFile      -Config $ConfigPatchArgs
    Update-ApiServer5TempFsConfig   -FilePath $tempfsFile    -Config $ConfigPatchArgs
    Update-ApiServerRootAppSettings -FilePath $rootAppFile   -Config $ConfigPatchArgs

    # Clean up obsolete flat root file if present
    $oldRoot = Join-Path $idpDestDir 'idpconfig.json'
    if (Test-Path $oldRoot) { Remove-Item -Path $oldRoot -Force -ErrorAction SilentlyContinue }

    # -------------------------------------------------------------------------
    # 5D. Apply Security ACLs to Directories and Config Files
    # -------------------------------------------------------------------------
    Write-Host "  Applying IIS_IUSRS and MachineKey ACL permissions..." -ForegroundColor Yellow
    foreach ($dir in @($idpDestDir, $idpConfigDir, $apiDestDir, $apiConfigDir)) {
        if (Test-Path $dir) {
            $acl = Get-Acl $dir
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule('IIS_IUSRS', 'ReadAndExecute, Synchronize', 'ContainerInherit, ObjectInherit', 'None', 'Allow')
            $acl.SetAccessRule($rule)
            Set-Acl -Path $dir -AclObject $acl
        }
    }
    Get-ChildItem -Path @($idpConfigDir, $apiConfigDir) -Filter '*.json' | ForEach-Object {
        $acl = Get-Acl $_.FullName
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule('IIS_IUSRS', 'Write, Read', 'Allow')))
        Set-Acl -Path $_.FullName -AclObject $acl
    }
    $machineKeysDir = Join-Path $env:ProgramData 'Microsoft\Crypto\RSA\MachineKeys'
    if (Test-Path $machineKeysDir) {
        $kacl = Get-Acl $machineKeysDir
        $kacl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule('IIS_IUSRS', 'Read', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
        Set-Acl -Path $machineKeysDir -AclObject $kacl
    }

    # -------------------------------------------------------------------------
    # 5E. Database schemas: hsiidp (IdS) + hsitempfile (TempFS) - before AppPool restart
    # -------------------------------------------------------------------------
    Write-Host "  [5E] TempFS tool package: $TempFsCfgPackage $TempFsCfgVer"
    Write-Host "`n[5E] Creating/validating hsiidp and hsitempfile schemas..." -ForegroundColor Cyan

    # --- IdS (hsiidp) ---
    Invoke-IdsOperationalStoreMigration -SqlInstance $SqlInst -DatabaseName $DbName `
        -HsiCredential $HsiCred -IdsAppSettings $idsAppFile
    Test-DbSchemaTables -SqlInstance $SqlInst -DatabaseName $DbName -HsiCredential $HsiCred `
        -Schema 'hsiidp' -RequiredTables 'devicecodes','persistedgrants','Keys','ServerSideSessions','PushedAuthorizationRequests'

    # --- TempFS (hsitempfile) ---
    <#$tempFsExe = Join-Path $apiConfigDir 'TempFileService.Configuration.exe'
    if (-not (Test-Path $tempFsExe)) {
        $tempFsExe = Get-ChildItem -Path $apiDestDir -Recurse -Filter 'TempFileService.Configuration.exe' -ErrorAction SilentlyContinue |
                     Select-Object -First 1 -ExpandProperty FullName
    }
    if (-not $tempFsExe) { throw "TempFileService.Configuration.exe not found under $apiDestDir. Copy it to $apiConfigDir." }
    #>

    # --- TempFS config tool: download from ProGet if not already present ---
    $tempFsExe = Join-Path $apiConfigDir 'TempFileService.Configuration.exe'
    if (-not (Test-Path $tempFsExe)) {
        if (-not $TempFsCfgVer) {
            throw "TempFileService.Configuration.exe missing and no -TempFsConfigVersion supplied."
        }
        $pkgRoot = Get-RemoteProGetInstaller -FeedUrl $ProGetFeed -PackageId $TempFsCfgPackage -Version $TempFsCfgVer
        $exeDir = Get-ChildItem -Path $pkgRoot -Recurse -Filter 'TempFileService.Configuration.exe' -ErrorAction SilentlyContinue |
                  Select-Object -First 1 -ExpandProperty DirectoryName
        if (-not $exeDir) { throw "TempFileService.Configuration.exe not found inside $TempFsCfgPackage $TempFsCfgVer ($pkgRoot)." }

        # Copy the exe plus its DLLs/runtime files
        Copy-Item -Path (Join-Path $exeDir '*') -Destination $apiConfigDir -Recurse -Force
    }
    Write-Host "  Using TempFS config tool: $tempFsExe" -ForegroundColor Green

    # Hyland order: 'security set-global-scheme' must run BEFORE 'database update'. Add that call here.

    Invoke-TempFsDatabaseUpdate -TempFsConfigExe $tempFsExe -TempFsJson $tempfsFile
    Test-DbSchemaTables -SqlInstance $SqlInst -DatabaseName $DbName -HsiCredential $HsiCred `
        -Schema 'hsitempfile' -RequiredTables 'storagesystems','uncstoragesystems','s3storagesystems','namespaces','tenantsettings','shorttermtempfiles','keks','deks'

    # Still manual: security set-tenant-scheme -> storage-scriptable add (UNC/S3) -> storage-scriptable set-active
    # -------------------------------------------------------------------------
    # 5F. Safe AppPool Restart
    # -------------------------------------------------------------------------
    Restart-HylandWebAppPoolSafely -AppPoolName "$($IdSApp)AppPool"
    Restart-HylandWebAppPoolSafely -AppPoolName "$($ApiApp)AppPool"

    # -------------------------------------------------------------------------
    # 6. Verification with Warm-up Delay & Explicit TLS 1.2
    # -------------------------------------------------------------------------
    Write-Host "`n[6/6] Validating Endpoint Health (Waiting 5s for AppPool warm-up)..." -ForegroundColor Cyan
    Start-Sleep -Seconds 5

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

    try {
        $resp = Invoke-WebRequest -Uri "https://$fqdn/$ApiApp/onbase/core/healthcheck" -Method Get -UseBasicParsing -TimeoutSec 20
        Write-Host "  [OK] API Server Healthcheck: https://$fqdn/$ApiApp/onbase/core/healthcheck (Status: $($resp.StatusCode))" -ForegroundColor Green
    } catch { Write-Warning "  [WARN] API Server check: $_" }

    try {
        $resp = Invoke-WebRequest -Uri "https://$fqdn/$IdSApp/admin" -Method Head -UseBasicParsing -TimeoutSec 20
        Write-Host "  [OK] Identity Service Admin portal reachable: https://$fqdn/$IdSApp/admin (Status: $($resp.StatusCode))" -ForegroundColor Green
    } catch { Write-Warning "  [WARN] IdS portal check: $_" }

    Write-Host "`nSetup Completed Successfully on $fqdn!" -ForegroundColor Green
}
# ==============================================================================
# Router
# ==============================================================================
if ($DeploymentMode -eq 'Remote') {
    if (-not $TargetMachine -or -not $Credential) {
        throw "TargetMachine and Credential are required for Remote mode."
    }

    $invokeParams = @{
        ComputerName = $TargetMachine
        Credential = $Credential
        ConfigurationName = $ConfigurationName
        ScriptBlock = $installLogic
        ArgumentList = @(
            $IdSSourcePath,
            $ApiServerSourcePath,
            $DotNetHostingBundlePath,
            $ProGetFeedUrl,
            $IdSVersion,
            $ApiServerVersion,
            $IdSPackageId,
            $ApiServerPackageId,
            $SqlInstance,
            $DatabaseName,
            $HsiCredential,
            $IdSAppName,
            $ApiAppName,
            $CertificateThumbprint,
            $ForceReinstall.IsPresent,
            $TempFsConfigPackageId,
            $TempFsConfigVersion,
            $IdsAdminUsers,
            $IdsAdminGroups,
            $IdsProvidersJson
        )
    }
    Write-Host "[DEBUG] Outer IdsProvidersJson: $(if ($IdsProvidersJson) { "$($IdsProvidersJson.Length) chars" } else { 'EMPTY' })" -ForegroundColor Magenta
    Invoke-Command @invokeParams
} else {
    & $installLogic `
        $IdSSourcePath `
        $ApiServerSourcePath `
        $DotNetHostingBundlePath `
        $ProGetFeedUrl `
        $IdSVersion `
        $ApiServerVersion `
        $IdSPackageId `
        $ApiServerPackageId `
        $SqlInstance `
        $DatabaseName `
        $HsiCredential `
        $IdSAppName `
        $ApiAppName `
        $CertificateThumbprint `
        $ForceReinstall.IsPresent `
        $TempFsConfigPackageId `
        $TempFsConfigVersion `
        $IdsAdminUsers `
        $IdsAdminGroups `
        $IdsProvidersJson
}
