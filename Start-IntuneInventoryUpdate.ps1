
function Get-GraphToken ($appClientId, $appClientSecret) {
    Write-Verbose '***Requesting GraphToken...' 

    # o365 tenant
    $tenantId = 'episerver99.onmicrosoft.com'
  
    # build parameters to make the API call
    $irmParams = @{
        Method = 'Post'
        Uri    = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        Body   = @{
            client_id     = $appClientId
            client_secret = $appClientSecret
            grant_type    = 'client_credentials'
            scope         = 'https://graph.microsoft.com/.default'
        }
    }
    # make the API call and store the resulting token
    $OAuthResponse = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    # if the token request is successful return the token
    if ($OAuthResponse) {
        Write-Verbose '***Token request successful' 
        return $OAuthResponse.access_token

    }
    else {
        # if request is not successful, write an error
        Write-Error 'Token request NOT successful'
    }

}

function Get-ManagedDevices ($graphToken) {
    Write-Verbose '***Getting Intune Managed Devices...' 

    # API parameters to request sign in logs
    $irmParams = @{
        Headers     = @{
            'Content-Type'  = 'application\json'
            'Authorization' = "Bearer $graphToken"
        }
        Method      = 'GET'
        ContentType = 'application/json'
    }
    
    # url for intune managed device data
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
 
    $result = [System.Collections.Generic.List[System.Object]]::New()
    #$return = [System.Collections.Generic.List[System.Object]]::New()
    While ($uri) {
        #Perform pagination if next page link (odata.nextlink) returned.
        $Response = Invoke-WebRequest @irmParams -Uri $uri | ConvertFrom-Json
        if ($Response.value) {
            foreach ($device in $Response.value) {
                if ($device.managedDeviceOwnerType -eq 'company') {
                    $result.Add($device)
                }
            }
        }
        $uri = $Response.'@odata.nextlink'
    }
    
    # filter out only company managed devcies
    

    return $result
}

function Get-AzureUserLocation ($graphToken, $azureUpn) {
    # parameters:
    #   $graphToken - access token as requested from Get-GraphToken function
    #

    # API parameters to request sign in logs
    $irmParams = @{
        Headers     = @{
            'Content-Type'  = 'application\json'
            'Authorization' = "Bearer $graphToken"
        }
        Method      = 'GET'
        ContentType = 'application/json'
        # because requesting specific user, upn is included in URI
        Uri         = "https://graph.microsoft.com/v1.0/users/$azureUpn"
    }
    
    # make call to Azure Graph
    $response = Invoke-WebRequest @irmParams | ConvertFrom-Json
    if ($response) {
        # return user with location if user is found
        # user should always be found because user upn is coming directly from
        # intune
        return $response
    }
}

function Invoke-SnipeItApiThrottleManagement {
    <#
        This function tracks API usage to avoid exceeding the permitted 120
        calls per minute. Overview:
        - add the current date/time to the tracking array
        - filter out any date/times that aren't in the past minute
        - if the count is >= to 120, identiy the oldest call and the number of
            seconds until it drops off the list
        - sleep that number of seconds then resume normally
    #>
    
    # add a new call to the list
    $global:apiCalls += $(Get-Date)
 
    # filter out calls from only the past minute
    $global:apiCalls = foreach ($call in $global:apiCalls) {
        # if NOW less call dateTime is under 1 minute
        if (( (Get-Date) - $call ).Minutes -lt 1) {
            # include this call in the updated list
            $call
        }
    }

    Write-Verbose "***API Monitor: $($global:apicalls.count)"

    # if the number of api calls in the past minute exceeds 120
    if ($global:apiCalls.count -ge 115) {
        # need to throttle
        # determine number of seconds until oldest call expires
        # + 1 seconds to ensure we always stay under the throttle limit
        $secondsToSleep = 61 - ((Get-Date) - $global:apiCalls[0]).seconds
        Write-Verbose "*** API Monitor: $($global:apicalls.count) calls. Sleeping $secondsToSleep seconds..." -Verbose
        $ctr = 0
        do {
            $ctr++
            Write-Verbose $ctr -Verbose
            Start-Sleep -Seconds 1
        } until ($ctr -eq $secondsToSleep)

        # resume processing normally
        Write-Verbose "*** API Monitor: resuming..." -Verbose
    }

}

function Get-SnipeItData($snipeitToken, $apiEndpoint, $searchQuery) {
    # parameters to make api call
    $irmParams = @{
        Headers = @{
            Accept        = 'application/json'
            Authorization = "Bearer $snipeitToken"
        }
        Method  = 'GET'
    }

    # base url
    $baseUri = 'https://snipeit.internal.optimizely.com/api/v1/' + $apiEndpoint + '?'
    
    # if a manufacturer is being queried add the search parameter
    if ($searchQuery) {
        $baseUri += "search=$searchQuery&"
    }
    
    $return = [System.Collections.Generic.List[System.Object]]::New()
    $offset = 0
    $limit = 500
    do {
        # add the remaining parameters    
        $uriParams = (
            "offset=$offset", # offset to start pulling from
            "limit=$limit", # number of entries to pull this request
            'sort=created_at', # field to sort by, sorty by created date
            'order=desc', # sort ascending or descending
            'deleted=false', # return only deleted users
            'all=false'         # return deleted users with active users
        ) -join "&"             # join all parameters with ampersand

        # build complete url
        $Uri = $baseUri + $uriParams

        # retrieve $limit number of rows from snipeit
        $response = Invoke-RestMethod @irmParams -Uri $Uri
        Invoke-SnipeItApiThrottleManagement

        if ($response.messages -eq 'Server Error') {
            $response
        }

        $rows = $response.rows | Sort-Object -Property id
        
        # add each row to the return object
        foreach ($row in $rows) {
            $return.Add($row)
        }
        
        # increase the offset by the limit
        $offset += $limit

    } until ($offset -gt $response.total)
    
    return $return

}

function New-SnipeItManufacturer($snipeitToken, $manufacturer) {
    # parameters to make api call
    $irmParams = @{
        Headers = @{
            Accept         = 'application/json'
            Authorization  = "Bearer $snipeitToken"
            'Content-Type' = 'application/json'
        }
        Method  = 'POST'
        Uri     = 'https://snipeit.internal.optimizely.com/api/v1/manufacturers'
        Body    = @{
            name = $manufacturer
        } | ConvertTo-Json
    }
       
    # make api call to add manufacturer   
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        Write-Error $error[0].Exception
        $response
    }
    else {
        return $response.payload
    }
    
}

function New-SnipeItLocation($snipeitToken, $locationName) {
    # parameters to make api call
    $irmParams = @{
        Headers = @{
            Accept         = 'application/json'
            Authorization  = "Bearer $snipeitToken"
            'Content-Type' = 'application/json'
        }
        Method  = 'POST'
        Uri     = 'https://snipeit.internal.optimizely.com/api/v1/locations'
        Body    = @{
            name = $locationName
        } | ConvertTo-Json
    }
       
    # make api call to add location   
    Write-Verbose "Creating New Location: $locationName" -Verbose
    $irmParams
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        Write-Error $error[0].Exception
        $response.messages
    }
    else {
        Write-Verbose "***New location created successfully" 
        return $response.payload
    }
    
}

function New-SnipeItModel($snipeitToken, $intuneModel) {
    # parameters to make api call
    $irmParams = @{
        Headers = @{
            Accept         = 'application/json'
            Authorization  = "Bearer $snipeitToken"
            'Content-Type' = 'application/json'
        }
        Method  = 'POST'
        Uri     = 'https://snipeit.internal.optimizely.com/api/v1/models'
        Body    = @{
            name            = $intuneModel.name
            category_id     = $intuneModel.category_id
            manufacturer_id = $intuneModel.manufacturer_id
            model_number    = $intuneModel.model
            fieldset_id     = $intuneModel.fieldset_id

        } | ConvertTo-Json
    }
       
    # make api call to add manufacturer   
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        Write-Error $error[0].Exception
        $response
    }
    else {
        return $response.payload
    }
    
}

function New-SnipeItAssetCheckout($snipeItToken, $assetId, $userId) {
    # parameters to make api call
    $irmParams = @{
        Headers = @{
            Accept         = 'application/json'
            Authorization  = "Bearer $snipeitToken"
            'Content-Type' = 'application/json'
        }
        Method  = 'POST'
        Uri     = "https://snipeit.internal.optimizely.com/api/v1/hardware/$assetId/checkout"
        Body    = @{
            checkout_to_type = 'user'
            status_id        = '4'        
            assigned_user    = $userId
            note             = 'Checked out by Intune->SnipeIt Sync Script'
        } | ConvertTo-Json
    }
       
    # make api call to checkout asset   
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        $irmParams.Body | ConvertTo-Json
        
        Write-Error $error[0].Exception
        $response.messages
    }
    else {
        return $response.payload
    }
}

function New-SnipeItAssetCheckin($snipeItToken, $assetId) {
    # parameters to make api call
    $irmParams = @{
        Headers = @{
            Accept         = 'application/json'
            Authorization  = "Bearer $snipeitToken"
            'Content-Type' = 'application/json'
        }
        Method  = 'POST'
        Uri     = "https://snipeit.internal.optimizely.com/api/v1/hardware/$assetId/checkin"
        Body    = @{
            status_id = '4'        
            note      = 'Checked in by Intune->SnipeIt Sync Script'
        } | ConvertTo-Json
    }
       
    # make api call to add manufacturer   
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        Write-Error $error[0].Exception
        $response
    }
    else {
        return $response.payload
    }
}

function New-SnipeItAsset($snipeitToken, $asset) {
    # parameters to make api call
    $irmParams = @{
        Headers = @{
            Accept         = 'application/json'
            Authorization  = "Bearer $snipeitToken"
            'Content-Type' = 'application/json'
        }
        Method  = 'POST'
        Uri     = 'https://snipeit.internal.optimizely.com/api/v1/hardware'
        Body    = @{
            status_id       = '4' # In Production
            model_id        = $asset.model_id
            name            = $asset.name
            serial          = $asset.serial
            last_audit_date = $asset.last_audit_date
            

        } | ConvertTo-Json
    }
       
    # make api call to add manufacturer   
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        Write-Error $error[0].Exception
        $response
    }
    else {
        return $response.payload
    }
    
}

function Set-SnipeItAsset {
    param(
        $snipeitToken,
        $asset_id, 
        $status_id, 
        $model_id, 
        $notes, 
        $last_checkout, 
        $assigned_to, 
        $company_id, 
        $serial, 
        $order_number,
        $warranty_months,
        $purchase_cost,
        $purchase_date,
        $requestable,
        $archived,
        $rtd_location_id,
        $name,
        $location_id
    )

    # check each parameter for a value and if present, add it to body
    [pscustomobject]$body = @{}
    if ($status_id) { $body.Add('status_id', $status_id) } 
    if ($model_id) { $body.Add('model_id', $model_id) } 
    if ($notes) { $body.Add('notes', $notes) } 
    if ($last_checkout) { $body.Add('last_checkout', $last_checkout) } 
    if ($assigned_to) { $body.Add('assigned_to', $assigned_to) } 
    if ($company_id) { $body.Add('company_id', $company_id) } 
    if ($serial) { $body.Add('serial', $serial) } 
    if ($order_number) { $body.Add('order_number', $order_number) }
    if ($warranty_months) { $body.Add('warranty_months', $warranty_months) }
    if ($purchase_cost) { $body.Add('purchase_cost', $purchase_cost) }
    if ($purchase_date) { $body.Add('purchase_date', $purchase_date) }
    if ($requestable) { $body.Add('requestable', $requestable) }
    if ($archived) { $body.Add('archived', $archived) }
    if ($rtd_location_id) { $body.Add('rtd_location_id', $rtd_location_id) }
    if ($name) { $body.Add('name', $name) }
    if ($location_id) { $body.Add('location_id', $location_id) }

    
    # parameters to make api call
    $irmParams = @{
        Headers = @{
            Accept         = 'application/json'
            Authorization  = "Bearer $snipeitToken"
            'Content-Type' = 'application/json'
        }
        Method  = 'PUT'
        Uri     = "https://snipeit.internal.optimizely.com/api/v1/hardware/$asset_id"
        Body    = $body | ConvertTo-Json
    }
    
    Write-Verbose "Asset ID: $asset_id"     
    Write-Verbose "Location ID: $location_id"     
    Write-Verbose $($irmParams | ConvertTo-Json ) 
    Write-Verbose "Body: $($body | ConvertTo-Json )" 

    # make api call to update asset   
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        Write-Error $error[0].Exception
        $response
    }
    else {
        Write-Verbose 'Asset update successful.' 
        return $response.payload
    }
    
}

function Update-SnipeItAssets ($snipeitToken, $intuneDevices) { 

    # loop through all intune devices and process them in 
    # to more usable data
    Write-Verbose "Processing $($intuneDevices.count) devices..." -Verbose
    $ctr = 0
    foreach ($device in $intuneDevices) {
        $ctr++
        # clear re-used variables
        $snipeItDevice = $null
        $manufacturer = $null
        $snipeitModel = $null
        $snipeitLocation = $null

        
        Write-Verbose "***Processing $($device.deviceName) ($ctr/$($intuneDevices.count))" -Verbose
        
        ##################
        ## Validate User #######################################################
        
        # match the intune user to a snipeit user for asset assgnment
        $snipeItUser = $null
        $snipeItUser = Get-SnipeItData -snipeitToken $snipeitToken -apiEndpoint 'users' -searchQuery $device.EmailAddress
        
        if (-not $snipeItUser) {
            # unable to match user, need to run an ldap sync

        }
        Write-Verbose "***User Identified: $($snipeItUser.name)" 
        ## END Validate User ###################################################
        ########################################################################

        ###############################
        ## Validate Device in Snipeit ###################################
        ## check if device has a serial number listed
        $snipeItDevice = $null
        if ($device.serialNumber) {
            # if serial number present, query snipeit with serial
            $snipeItDevice = Get-SnipeItData -snipeitToken $snipeitToken -apiEndpoint 'hardware' -searchQuery $device.serialnumber
        }
        else {
            # if serial not present, query snipeit with asset name
            $snipeItDevice = Get-SnipeItData -snipeitToken $snipeitToken -apiEndpoint 'hardware' -searchQuery $device.deviceName
        }
        ## END Validate Device #################################################
        ########################################################################
         

        ###############################        
        ## BEGIN IF Device does exist ##########################################
        if (-not $snipeItDevice) {
            Write-Verbose "***Device does NOT exist, creating..." 
            # device does not exist, create it
            # verify manufacturer or create

            ##########################################
            ## BEGIN Validate or create Manufacturer ###########################
            # query snipeit for manufacturer
            $manufacturer = $null
            $gsidParams = @{
                snipeitToken = $snipeitToken
                apiEndpoint  = 'manufacturers'
                searchQuery  = $device.manufacturer
            }
            $manufacturer = Get-SnipeItData @gsidParams

            # fix to match manufacturer by name if multiple matches exist
            if ($manufacturer.count -gt 1) {
                $manufacturer = $manufacturer | Where-Object {
                    $_.name -eq $device.manufacturer
                }
            }
            
            # if not found
            if (-not $manufacturer) {
                Write-Verbose "***Manufacturer: $($device.manufacturer) not found, creating..." 
                # manufacturer doesn't exist, create it
                $nsimParams = @{
                    snipeitToken = $snipeitToken
                    manufacturer = $device.manufacturer
                }
                $manufacturer = New-SnipeItManufacturer @nsimParams
                

                # fix to match manufacturer by name if multiple matches exist
                if ($manufacturer.count -gt 1) {
                    $manufacturer = $manufacturer | Where-Object {
                        $_.name -eq $device.manufacturer
                    }
                }
                
            }
            ## END Validate manufacturer #######################################
            ####################################################################
            
            #################################
            ## BEGIN Valide or create model ####################################
            # manufacturer exists, verify model
            Write-Verbose "***Manufacturer: $($device.manufacturer) exists, checking model..." 

            $snipeitModel = $null
            $gsidParams = @{
                snipeitToken = $snipeitToken
                apiENDpoint  = 'models'
                searchQuery  = $device.model
            }
            $snipeitModel = Get-SnipeItData @gsidParams

            # if model isn't found create it
            if (-not $snipeitModel) {
                Write-Verbose "***Model: $($device.model) not found, creating..." 
                
                # model doesn't exist, create it
                $model = @{
                    Name            = $device.model
                    Model           = $device.model
                    category_id     = '2'
                    manufacturer_id = $manufacturer.id
                    fieldset_id     = '1' 
                }
                $snipeitModel = New-SnipeItModel -snipeitToken $snipeitToken -intuneModel $model
                
            }
            ## END Valide or create model ######################################
            ####################################################################
            
            ######################################
            ## BEGIN Validate or create location ###############################
            Write-Verbose "***Model exists, checking location..." 
            
            # get user location from Azure Graph
            $userLocation = $null
            $gaulParams = @{
                graphToken = $graphToken
                azureUpn   = $device.userPrincipalName
            }
            $userLocation = (Get-AzureUserLocation @gaulParams).officeLocation

            # check snipeit to see if location exists
            $snipeitLocation = $null
            $gsidParams = @{
                snipeitToken = $snipeitToken
                apiEndpoint  = 'locations'
                searchQuery  = $userLocation
            }
            $snipeitLocation = Get-SnipeItData @gsidParams

            # if model isn't found create it
            if (-not $snipeitLocation) {
                Write-Verbose "***Location: $userLocation not found, creating..." 
                
                # model doesn't exist, create it
                $snipeitLocation = New-SnipeItLocation -snipeitToken $snipeitToken -location $userLocation
            }
            ## END Validate or create location #################################
            ####################################################################

            ###########################
            ## BEGIN Create new asset ##########################################
            # model exists create asset
            $snipeItDevice = New-SnipeItAsset -snipeitToken $snipeitToken -asset @{
                model_id        = $snipeitModel.id
                name            = $device.deviceName
                serial          = $device.serialNumber
                last_audit_date = $device.lastSyncDateTime
                location_id     = $snipeitLocation.id
            }
            Write-Verbose "***New Asset:" 
            $snipeItDevice | Format-Table id, name, serial, location, lastCheckout
            
            ## END Create new asset ############################################
            ####################################################################
        }
        
        Write-Verbose "***Devices exists, verifying user..." 
        
        # user snipeit shows currently assigned
        $snipeitUser = $snipeItDevice.assigned_to

        # user location from azure graph
        $intuneLocation = Get-AzureUserLocation -graphToken $graphToken -azureUpn $device.userPrincipalName
        Write-Verbose "***Intune Location: $intuneLocation" 

        # compare intune email address to snipeit user
        if ($device.emailAddress -ne $snipeitUser.username) {
            Write-Verbose "***Intune Assignee: $($device.emailAddress) != $($snipeItUser.username)" 
            Write-Verbose "***Checking in asset..." 
            # if user doesn't match
            # need to re-assign asset in snipeit
            # unsassign asset if it's assigned
            if ($snipeItDevice.assigned_to) {
                New-SnipeItAssetCheckin -snipeItToken $snipeitToken -assetId $snipeItDevice.id       
            }
            
            # use UPN to get correct user from snipeit
            # assign asset and leave note
            $gsidParams = @{
                snipeitToken = $snipeitToken
                apiEndpoint  = 'users'
                searchQuery  = $device.emailAddress
            }
            $snipeitUser = Get-SnipeItData @gsidParams | Where-Object {
                # fix for snipeit somehow returning multiple users
                $_.email -eq $device.emailaddress
            }
            Write-Verbose "*** Checking out asset..." -Verbose
            Write-Verbose "*** Intune Email: $($device.emailaddress)"
            New-SnipeItAssetCheckout -snipeItToken $snipeitToken -assetId $snipeItDevice.id -userId $snipeItUser.id
        }

        # compare intune user location to snipeit asset location
        $snipeitDeviceLocation = $null
        $snipeitDeviceLocation = Get-SnipeItData -snipeitToken $snipeitToken -apiEndpoint 'locations' | Where-Object {
            $_.id -eq $snipeItDevice.location.id
        }
        #Write-Verbose "SnipeIt Location: $($snipeitLocation | ConvertTo-Json)" -Verbose
        #Write-Verbose "Intune Location: $($intuneLocation.officeLocation)" -Verbose
        
        if ($intuneLocation.officeLocation -ne $snipeitDeviceLocation.name) {
            Write-Verbose "Intune Location: $($intuneLocation.officeLocation) does not match $($snipeitDeviceLocation.name)" -Verbose
            
            # if location doesn't match verify location exists
            $gsidParams = @{
                snipeitToken = $snipeitToken
                apiEndpoint  = 'locations'
                searchQuery  = $intunelocation.officeLocation
            }
            $snipeitLocation = Get-SnipeItData @gsidParams | Where-Object {
                $_.name -eq $intuneLocation.officeLocation
            }

            # verify resulting name matches exactly (avoids partial matches)
            if (-not $snipeitLocation) {
                # if location doesn't exist, create it
                $snipeitLocation = New-SnipeItLocation -snipeitToken $snipeitToken -location $intuneLocation.officeLocation
            }

            # update asset location
            
            Write-Verbose "*** Calling Asset update with location: $($snipeitLocation.id)" -Verbose
            Set-SnipeItAsset -snipeitToken $snipeitToken -asset_id $snipeItDevice.id -location_id $snipeitLocation.id
        }
    }
    ## END If Device does exist ########################################
    ####################################################################

    <# #  checkout asset if it isn't assigned to a user
        if (-not $snipeItDevice.assigned_to) {
            New-SnipeItAssetCheckout -snipeItToken $snipeitToken -assetId $snipeItDevice.id -userId $snipeItUser.id
            
        }
        else {
            Write-Verbose '***Device already assigned...' 
        } #>
}

<#
To Do:
- get jamf import working
- recurring schedule with proper credential storage
- handle override or not overriding a user assignment that doesn't match
- setup a dedicated snipeit user so it doesn't reflect my user account checking
    out all assets
#>

#$VerbosePreference = 'Continue'
[System.DateTime[]]$global:apiCalls = @()

#########################################
## prepare tokens for Graph and Snipeit ########################################
#

# Graph client id from stored secure string
$appClientId = [pscredential]::New(
    'user',
    $(Get-Content .\graphAppClientId | ConvertTo-SecureString)).GetNetworkCredential().Password


# Graph client secret from stored secure string
$appClientSecret = [pscredential]::New(
    'user',
    $(Get-Content .\graphAppClientSecret | ConvertTo-SecureString)).GetNetworkCredential().Password


# retrieve graph token with client id and secret
$graphToken = Get-GraphToken $appClientId $appClientSecret

# snipeit token from stored secure string
$snipeItToken = [PSCredential]::New(
    'user',
    $(Get-Content -Path .\apiKey | ConvertTo-SecureString)).GetNetworkCredential().Password

## End Prepare tokens ##########################################################
#######################

# retrieve intune devices with graph token
$intuneDevices = Get-ManagedDevices($graphToken)

# run the sync
Update-SnipeItAssets -snipeitToken $snipeItToken -intuneDevices $intuneDevices
