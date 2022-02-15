<#
.SYNOPSIS
    Rotate Bitlocker Recovery keys using Intune - via MS Graph API.
.DESCRIPTION
    This script will invoke the recovery key rotation using the same process as clicking on the "Rotate Recovery Key" button in the Endpoint Management portal, but in bulk
.PARAMETER TenantID
    Specify the Azure AD tenant ID.
.PARAMETER ClientID
    Specify the service principal, also known as app registration, Client ID (also known as Application ID).
.PARAMETER State
    Specify -TenantID and -ClientID, or edit the script and add hard code it
.EXAMPLE
    # Rotate Bitlocker Recovery Keys for all devices in estate: NOTE: Windows 10 Build 1909 and above required
    .\RotateBitlockerKeys-Parallel.ps1 -TenantID "<tenant_id>" -ClientID "<client_id>"
.NOTES
    FileName:    RotateBitlockerKeys-Parallel.ps1
    Author:      Christopher Baxter
    Contact:     GitHub - https://github.com/christopherbaxter
    Created:     2021-11-01
    Updated:     2022-02-03

    Depending on the size of your estate and the speed of your connection, this script may take a significant amount of time to run. Make sure that your elevated rights in AzureAD\Intune have an appropriate amount of time for this to function.
#>
#Requires -Modules "MSAL.PS","PoshRSJob","ImportExcel","JoinModule"
[CmdletBinding(SupportsShouldProcess = $TRUE)]
param(
    #PLEASE make sure you have specified your details below, else edit this and use the switches\variables in command line.
    [parameter(Mandatory = $TRUE, HelpMessage = "Specify the Azure AD tenant ID.")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantID,

    [parameter(Mandatory = $TRUE, HelpMessage = "Specify the service principal, also known as app registration, Client ID (also known as Application ID).")]
    [ValidateNotNullOrEmpty()]
    [string]$ClientID
)
Begin {}
Process {
    
    #############################################################################################################################################
    # Functions
    #############################################################################################################################################

    function New-AuthenticationHeader {
        <#
        .SYNOPSIS
            Construct a required header hash-table based on the access token from Get-MsalToken cmdlet.
        .DESCRIPTION
            Construct a required header hash-table based on the access token from Get-MsalToken cmdlet.
        .PARAMETER AccessToken
            Pass the AuthenticationResult object returned from Get-MsalToken cmdlet.
        .NOTES
            Author:      Nickolaj Andersen
            Contact:     @NickolajA
            Created:     2020-12-04
            Updated:     2020-12-04
            Version history:
            1.0.0 - (2020-12-04) Script created
        #>
        param(
            [parameter(Mandatory = $tRUE, HelpMessage = "Pass the AuthenticationResult object returned from Get-MsalToken cmdlet.")]
            [ValidateNotNullOrEmpty()]
            [Microsoft.Identity.Client.AuthenticationResult]$AccessToken
        )
        Process {
            # Construct default header parameters
            $AuthenticationHeader = @{
                "Content-Type"  = "application/json"
                "Authorization" = $AccessToken.CreateAuthorizationHeader()
                "ExpiresOn"     = $AccessToken.ExpiresOn.LocalDateTime
            }
    
            # Amend header with additional required parameters for bitLocker/recoveryKeys resource query
            $AuthenticationHeader.Add("ocp-client-name", "My App")
            $AuthenticationHeader.Add("ocp-client-version", "1.2")
    
            # Handle return value
            return $AuthenticationHeader
        }
    }

    function Invoke-MSGraphOperation {
        <#
        .SYNOPSIS
            Perform a specific call to Graph API, either as GET, POST, PATCH or DELETE methods.
            
        .DESCRIPTION
            Perform a specific call to Graph API, either as GET, POST, PATCH or DELETE methods.
            This function handles nextLink objects including throttling based on retry-after value from Graph response.
            
        .PARAMETER Get
            Switch parameter used to specify the method operation as 'GET'.
            
        .PARAMETER Post
            Switch parameter used to specify the method operation as 'POST'.
            
        .PARAMETER Patch
            Switch parameter used to specify the method operation as 'PATCH'.
            
        .PARAMETER Put
            Switch parameter used to specify the method operation as 'PUT'.
            
        .PARAMETER Delete
            Switch parameter used to specify the method operation as 'DELETE'.
            
        .PARAMETER Resource
            Specify the full resource path, e.g. deviceManagement/auditEvents.
            
        .PARAMETER Headers
            Specify a hash-table as the header containing minimum the authentication token.
            
        .PARAMETER Body
            Specify the body construct.
            
        .PARAMETER APIVersion
            Specify to use either 'Beta' or 'v1.0' API version.
            
        .PARAMETER ContentType
            Specify the content type for the graph request.
            
        .NOTES
            Author:      Nickolaj Andersen & Jan Ketil Skanke & (very little) Christopher Baxter
            Contact:     @JankeSkanke @NickolajA
            Created:     2020-10-11
            Updated:     2020-11-11
    
            Version history:
            1.0.0 - (2020-10-11) Function created
            1.0.1 - (2020-11-11) Tested in larger environments with 100K+ resources, made small changes to nextLink handling
            1.0.2 - (2020-12-04) Added support for testing if authentication token has expired, call Get-MsalToken to refresh. This version and onwards now requires the MSAL.PS module
            1.0.3.Custom - (2020-12-20) Added aditional error handling. Not complete, but more will be added as needed. Christopher Baxter
        #>
        param(
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Switch parameter used to specify the method operation as 'GET'.")]
            [switch]$Get,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST", HelpMessage = "Switch parameter used to specify the method operation as 'POST'.")]
            [switch]$Post,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH", HelpMessage = "Switch parameter used to specify the method operation as 'PATCH'.")]
            [switch]$Patch,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT", HelpMessage = "Switch parameter used to specify the method operation as 'PUT'.")]
            [switch]$Put,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE", HelpMessage = "Switch parameter used to specify the method operation as 'DELETE'.")]
            [switch]$Delete,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Specify the full resource path, e.g. deviceManagement/auditEvents.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [string]$Resource,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Specify a hash-table as the header containing minimum the authentication token.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [System.Collections.Hashtable]$Headers,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST", HelpMessage = "Specify the body construct.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [ValidateNotNullOrEmpty()]
            [System.Object]$Body,
    
            [parameter(Mandatory = $FALSE, ParameterSetName = "GET", HelpMessage = "Specify to use either 'Beta' or 'v1.0' API version.")]
            [parameter(Mandatory = $FALSE, ParameterSetName = "POST")]
            [parameter(Mandatory = $FALSE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $FALSE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $FALSE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("Beta", "v1.0")]
            [string]$APIVersion = "Beta",
    
            [parameter(Mandatory = $FALSE, ParameterSetName = "GET", HelpMessage = "Specify the content type for the graph request.")]
            [parameter(Mandatory = $FALSE, ParameterSetName = "POST")]
            [parameter(Mandatory = $FALSE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $FALSE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $FALSE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("application/json", "image/png")]
            [string]$ContentType = "application/json"
        )
        Begin {
            # Construct list as return value for handling both single and multiple instances in response from call
            $GraphResponseList = New-Object -TypeName "System.Collections.ArrayList"
            $Runcount = 1
    
            # Construct full URI
            $GraphURI = "https://graph.microsoft.com/$($APIVersion)/$($Resource)"
            Write-Verbose -Message "$($PSCmdlet.ParameterSetName) $($GraphURI)"
        }
        Process {
            # Call Graph API and get JSON response
            do {
                try {
                    # Determine the current time in UTC
                    $UTCDateTime = (Get-Date).ToUniversalTime()
    
                    # Determine the token expiration count as minutes
                    $TokenExpireMins = ([datetime]$Headers["ExpiresOn"] - $UTCDateTime).Minutes
    
                    # Attempt to retrieve a refresh token when token expiration count is less than or equal to 10
                    if ($TokenExpireMins -le 10) {
                        Write-Verbose -Message "Existing token found but has expired, requesting a new token"
                        $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -Silent -ForceRefresh
                        $Headers = New-AuthenticationHeader -AccessToken $AccessToken
                    }
    
                    # Construct table of default request parameters
                    $RequestParams = @{
                        "Uri"         = $GraphURI
                        "Headers"     = $Headers
                        "Method"      = $PSCmdlet.ParameterSetName
                        "ErrorAction" = "Stop"
                        "Verbose"     = $VerbosePreference
                    }
    
                    switch ($PSCmdlet.ParameterSetName) {
                        "POST" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                        "PATCH" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                        "PUT" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                    }
    
                    # Invoke Graph request
                    $GraphResponse = Invoke-RestMethod @RequestParams
                    
                    # Handle paging in response
                    if ($GraphResponse.'@odata.nextLink') {
                        $GraphResponseList.AddRange($GraphResponse.value) | Out-Null
                        $GraphURI = $GraphResponse.'@odata.nextLink'
                        Write-Verbose -Message "NextLink: $($GraphURI)"
                    }
                    else {
                        # NextLink from response was null, assuming last page but also handle if a single instance is returned
                        if (-not([string]::IsNullOrEmpty($GraphResponse.value))) {
                            $GraphResponseList.AddRange($GraphResponse.value) | Out-Null
                        }
                        else {
                            $GraphResponseList.Add($GraphResponse) | Out-Null
                        }
                        
                        # Set graph response as handled and stop processing loop
                        $GraphResponseProcess = $TRUE
                    }
                }
                catch [System.Exception] {
                    $ExceptionItem = $PSItem
                    if ($ExceptionItem.Exception.Response.StatusCode -like "429") {
                        # Detected throttling based from response status code
                        $RetryInsecond = $ExceptionItem.Exception.Response.Headers["Retry-After"]
    
                        # Wait for given period of time specified in response headers
                        Write-Verbose -Message "Graph is throttling the request, will retry in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "Unauthorized") {
                        Write-Verbose -Message "Your Account does not have the relevent privilege to read this data. Please Elevate your account or contact the administrator"
                        $Script:PIMExpired = $tRUE
                        $GraphResponseProcess = $TRUE
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "GatewayTimeout") {
                        # Detected Gateway Timeout
                        $RetryInsecond = 30
    
                        # Wait for given period of time specified in response headers
                        Write-Verbose -Message "Gateway Timeout detected, will retry in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "NotFound") {
                        Write-Verbose -Message "The Device data could not be found"
                        $Script:StatusResult = $ExceptionItem.Exception.Response.StatusCode
                        $GraphResponseProcess = $TRUE
                    }
                    elseif ($PSItem.Exception.Message -like "*Invalid JSON primitive*") {
                        $Runcount++
                        $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -Silent -ForceRefresh
                        $Headers = New-AuthenticationHeader -AccessToken $AccessToken
                        if ($Runcount -ge 10) {
                            Write-Verbose -Message "An Unrecoverable Error occured - Error: Invalid JSON primitive"
                            $GraphResponseProcess = $TRUE
                        }
                        $RetryInsecond = 5
                        Write-Verbose -Message "Invalid JSON Primitive detected, Trying again in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                    }
                    elseif ($PSItem.Exception.Message -like "*null-valued*") {
                        $Runcount++
                        $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -Silent -ForceRefresh
                        $Headers = New-AuthenticationHeader -AccessToken $AccessToken
                        
                        if ($Runcount -ge 10) {
                            Write-Verbose -Message "An Unrecoverable Error occured - Error: You cannot call a method on a null-valued expression"
                            $GraphResponseProcess = $TRUE
                        }
                        $RetryInsecond = 5
                        Write-Verbose -Message "Null-Valued Expression detected, renewing roken and trying again in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                    }
                    else {
                        try {
                            # Read the response stream
                            $StreamReader = New-Object -TypeName "System.IO.StreamReader" -ArgumentList @($ExceptionItem.Exception.Response.GetResponseStream())
                            $StreamReader.BaseStream.Position = 0
                            $StreamReader.DiscardBufferedData()
                            $ResponseBody = ($StreamReader.ReadToEnd() | ConvertFrom-Json)
                            
                            switch ($PSCmdlet.ParameterSetName) {
                                "GET" {
                                    # Output warning message that the request failed with error message description from response stream
                                    Write-Warning -Message "Graph request failed with status code $($ExceptionItem.Exception.Response.StatusCode). Error message: $($ResponseBody.error.message)"
    
                                    # Set graph response as handled and stop processing loop
                                    $GraphResponseProcess = $TRUE
                                }
                                default {
                                    # Construct new custom error record
                                    $SystemException = New-Object -TypeName "System.Management.Automation.RuntimeException" -ArgumentList ("{0}: {1}" -f $ResponseBody.error.code, $ResponseBody.error.message)
                                    $ErrorRecord = New-Object -TypeName "System.Management.Automation.ErrorRecord" -ArgumentList @($SystemException, $ErrorID, [System.Management.Automation.ErrorCategory]::NotImplemented, [string]::Empty)
    
                                    # Throw a terminating custom error record
                                    $PSCmdlet.ThrowTerminatingError($ErrorRecord)
                                }
                            }
    
                            # Set graph response as handled and stop processing loop
                            $GraphResponseProcess = $TRUE
                        }
                        catch [System.Exception] {
                            if ($PSItem.Exception.Message -like "*Invalid JSON primitive*") {
                                $Runcount++
                                if ($Runcount -ge 10) {
                                    Write-Verbose -Message "An Unrecoverable Error occured - Error: Invalid JSON primitive"
                                    $GraphResponseProcess = $TRUE
                                }
                                $RetryInsecond = 5
                                Write-Verbose -Message "Invalid JSON Primitive detected, Trying again in $($RetryInsecond) seconds"
                                Start-Sleep -second $RetryInsecond
                                
                            }
                            else {
                                Write-Warning -Message "Unhandled error occurred in function. Error message: $($PSItem.Exception.Message)"
    
                                # Set graph response as handled and stop processing loop
                                $GraphResponseProcess = $TRUE
                            }
                        }
                    }
                }
            }
            until ($GraphResponseProcess -eq $TRUE)
    
            # Handle return value
            return $GraphResponseList
            
        }
    }

    #############################################################################################################################################
    # Variables
    #############################################################################################################################################

    $Script:PIMExpired = $null
    $Script:StatusResult = $null
    $FileDate = Get-Date -Format 'yyyy_MM_dd'
    [string]$Resource = "deviceManagement/managedDevices"
    $FilePath = "C:\Temp\BitlockerKeyEscrow\"
    $FailedReportFileName = "KeyRotationRequestFailure-$($FileDate).xlsx"
    $KeyRotationCompletionReportFileName = "KeyRotationRequest-$($FileDate).xlsx"
    $KeyRotationFailureFile = "$($FilePath)$($FailedReportFileName)"
    $KeyRotationReportFile = "$($FilePath)$($KeyRotationCompletionReportFileName)"
        
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

    #############################################################################################################################################
    # Get Auth token and create Authentication header
    #############################################################################################################################################

    #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    Try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
    catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
    
    #############################################################################################################################################
    # Extract list of Intune Device IDs from MSGraph - Or supply your own list of IntuneDeviceIDs
    #############################################################################################################################################

    $StartTime = Get-Date -Format 'yyyy/MM/dd HH:mm'
    $IntuneDevIDs = [System.Collections.ArrayList]::new()
    $IntuneDevIDs = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'" -Headers $AuthenticationHeader -Verbose:$VerbosePreference | Where-Object { $_.azureADDeviceId -ne "00000000-0000-0000-0000-000000000000" } | Select-Object id | Sort-Object id)

    # This will be if you want to specify the IntuneDeviceIDs

    #$InputFile = "$($FilePath)InputFiles\TargetedIntuneDeviceIDs.csv"
    #$IntuneDevIDs | Export-Csv $InputFile -NoTypeInformation -Delimiter ";"
    #$IntuneDevIDs = @(Import-Csv -Path $InputFile)
    #$IntuneDevIDs = $IntuneDevIDs[0..999] # I use this for testing on a small number of devices
    
    #############################################################################################################################################
    # Split the IntuneDeviceID array into smaller chunks (parts) and set Throttle limit for parallel processing
    #############################################################################################################################################

    [int]$parts = 30
    $PartSize = [Math]::Ceiling($IntuneDevIDs.count / $parts)
    $SplitDevicelist = @()
    for ($i = 1; $i -le $parts; $i++) {
        $start = (($i - 1) * $PartSize)
        $end = (($i) * $PartSize) - 1
        if ($end -ge $IntuneDevIDs.count) { $end = $IntuneDevIDs.count }
        $SplitDevicelist += , @($IntuneDevIDs[$start..$end])
    }

    $ThrottleLimit = 30
    
    #############################################################################################################################################
    # Specify the ScriptBlock for the PoshRSJob
    #############################################################################################################################################

    $ScriptBlock = {
        param (
            [System.Collections.Hashtable]$AuthenticationHeader, [string]$Resource, [string]$cID, [string]$tID, [string]$APIVersion
        )
        if (-not($Runcount)) {
            $Runcount = 0
        }

        $Runcount++
        # This will refresh the authtoken and the authentication header after every 1000 'runs\cycles'
        if ($Runcount -ge 1000) {
            #$AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ErrorAction Stop
            #$AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

            #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
            try { $AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ForceRefresh -Silent -ErrorAction Stop }
            catch { $AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ErrorAction Stop }
            if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
            $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

            $Runcount = 0 
        }
        
        $GraphURI = "https://graph.microsoft.com/$($APIVersion)/$($Resource)/$($_)/rotateBitLockerKeys"
        $RequestParams = @{
            "Uri"         = $GraphURI
            "Headers"     = $AuthenticationHeader
            "Method"      = "POST"
            "ErrorAction" = "Stop"
            "Verbose"     = $VerbosePreference
        }

        # Invoke Graph request
        try {
            Invoke-RestMethod @RequestParams
            if ($?) {
                #$Requested = "Successfull"
                $obj = New-Object psobject
                $obj | Add-Member -Name ID -type noteproperty -Value $_
                $obj | Add-Member -Name KeyRotationRequest -type noteproperty -Value "Successful"
                Return $obj
            }
        }
        catch {
            #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
            try { $AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ForceRefresh -Silent -ErrorAction Stop }
            catch { $AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ErrorAction Stop }
            if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
            $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
            $RequestParams = @{
                "Uri"         = $GraphURI
                "Headers"     = $AuthenticationHeader
                "Method"      = "POST"
                "ErrorAction" = "Stop"
                "Verbose"     = $VerbosePreference
            }
            try {
                Invoke-RestMethod @RequestParams
                if ($?) {
                    #$Requested = "Successfull"
                    $obj = New-Object psobject
                    $obj | Add-Member -Name ID -type noteproperty -Value $_
                    $obj | Add-Member -Name KeyRotationRequest -type noteproperty -Value "Successful"
                    Return $obj
                }
            }
            catch {
                #$Requested = "Failed"
                $obj = New-Object psobject
                $obj | Add-Member -Name ID -type noteproperty -Value $_
                $obj | Add-Member -Name KeyRotationRequest -type noteproperty -Value "Failed"
                Return $obj
            }
        }
    }

    #############################################################################################################################################
    # Foreach loop to run through all items in each 'split' array
    #############################################################################################################################################

    # Processing Time is 4 Hours
    $ExtractionReport = [System.Collections.ArrayList]::new()
    $Counter = 0
    Foreach ($i in $SplitDevicelist) {
        # Get authentication token
        #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
        try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
        catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
        if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
        $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
        
        $Counter++
                
        Write-Host "Rotation Round number $($Counter) of $($parts)"

        # This is the little code snippet that allows the parallel processing of the data. Remember that $i is actually an array.
        $RawExtractionReport = @($i.id | Start-RSJob -ScriptBlock $ScriptBlock -Throttle $ThrottleLimit -ArgumentList $AuthenticationHeader, $Resource, $ClientID, $TenantID, "Beta" | Wait-RSJob -ShowProgress | Receive-RSJob)
        Get-RSJob | Remove-RSJob -Force

        $ExtractionReport += $RawExtractionReport
    }

    #############################################################################################################################################
    # Process the results and create an array with the failed extractions for retry
    #############################################################################################################################################

    $ExtractionReport = @($ExtractionReport | Where-Object { ($_.ID -notlike "Proxy*") -and ($_.KeyRotationRequest -Match "Successful" -or $_.KeyRotationRequest -Match "Failed") })
    $CompletionCheckArray = @($IntuneDevIDs | LeftJoin-Object $ExtractionReport -On ID)
    $FailedExtraction = @($CompletionCheckArray | Where-Object { ($_.KeyRotationRequest -notlike "Successful" -and $_.KeyRotationRequest -notlike "Failed") } | Select-Object ID)
    $CompletionCheckArray = @($CompletionCheckArray | Sort-Object id )

    #############################################################################################################################################
    # Split the failed extraction array into smaller chunks (parts) and set Throttle limit for parallel processing
    #############################################################################################################################################

    [int]$Fparts = 10
    $PartSize = [Math]::Ceiling($FailedExtraction.count / $Fparts)
    $FailedDevicelist = @()
    for ($i = 1; $i -le $Fparts; $i++) {
        $start = (($i - 1) * $PartSize)
        $end = (($i) * $PartSize) - 1
        if ($end -ge $FailedExtraction.count) { $end = $FailedExtraction.count }
        $FailedDevicelist += , @($FailedExtraction[$start..$end])
    }
    
    #############################################################################################################################################
    # Foreach loop to run through all items in each 'split' array - more of the same above
    #############################################################################################################################################

    $FailedExtractionReport = [System.Collections.ArrayList]::new()
    $FCounter = 0
    Foreach ($f in $FailedDevicelist) {
        # Get authentication token
        #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
        try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
        catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
        if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
        $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
        $FCounter++
                
        Write-Host "Extraction Round number $($FCounter) of $($Fparts)"

        $RawFailedExtractionReport = @($f.id | Start-RSJob -ScriptBlock $ScriptBlock -Throttle $ThrottleLimit -ArgumentList $AuthenticationHeader, $Resource, $ClientID, $TenantID, "Beta" | Wait-RSJob -ShowProgress | Receive-RSJob)
        Get-RSJob | Remove-RSJob -Force

        $FailedExtractionReport += $RawFailedExtractionReport
    }

    #############################################################################################################################################
    # Generate the reports and export them
    #############################################################################################################################################

    if ($FailedExtractionReport.count -gt 0) {
        $FailedExtractionReport = @($FailedExtractionReport | Where-Object { ($_.ID -notlike "Proxy*") -and ($_.KeyRotationRequest -like "Successful" -or $_.KeyRotationRequest -like "Failed") })
        $FailedCompletionCheckArray = @($FailedExtraction | LeftJoin-Object $FailedExtractionReport -On ID)
        $RepeatedExtractionFailures = @($FailedCompletionCheckArray | Where-Object { ($_.KeyRotationRequest -notlike "Successful" -and $_.KeyRotationRequest -notlike "Failed") } | Select-Object ID)
        $FailedCompletionCheckArray = @($FailedCompletionCheckArray | Sort-Object id)
        $RawReportingArray = @($CompletionCheckArray | LeftJoin-Object $FailedCompletionCheckArray -On id)
        $ReportingArray = @($RawReportingArray | Select-Object @{Name = "IntuneDeviceID"; Expression = { $_.id } }, @{Name = "KeyRotationRequest"; Expression = { if ($_.KeyRotationRequest -like "Successful") { "Successful" }elseif ($_.KeyRotationRequest -like "Failed") { "Failed" }else { "Failed" } } })
        $RepeatedExtractionFailures | Export-Excel $KeyRotationFailureFile -Verbose:$VerbosePreference -ClearSheet -AutoSize -AutoFilter
    }
    else {
        $ReportingArray = @($CompletionCheckArray | Select-Object @{Name = "IntuneDeviceID"; Expression = { $_.id } }, @{Name = "KeyRotationRequest"; Expression = { if ($_.KeyRotationRequest -like "Successful") { "Successful" }elseif ($_.KeyRotationRequest -like "Failed") { "Failed" }else { "Failed" } } })
    }

    $ReportingArray | Export-Excel $KeyRotationReportFile -Verbose:$VerbosePreference -ClearSheet -AutoSize -AutoFilter

    $EndTime = Get-Date -Format 'yyyy/MM/dd HH:mm'
    Write-Host "Process Started at $($StartTime) and completed at $($EndTime)"

}