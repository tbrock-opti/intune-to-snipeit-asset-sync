
function Get-GraphToken ($appClientId, $appClientSecret) {
    Write-Verbose 'Requesting GraphToken...'

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
        Write-Verbose 'Token request successful'
        return $OAuthResponse.access_token

    }
    else {
        # if request is not successful, write an error
        Write-Error 'Token request NOT successful'
    }

}

function Get-ManagedDevices ($graphToken) {
    Write-Verbose 'Getting Intune Managed Devices...'

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
        if (( (Get-Date) - $_ ).Minutes -lt 1) {
            # include this call in the updated list
            $call
        }
    }

    Write-Verbose "API Monitor: $($global:apicalls.count)"

    # if the number of api calls in the past minute exceeds 120
    if ($global:apiCalls.count -ge 120) {
        # need to throttle
        # determine number of seconds until oldest call expires
        # + 1 seconds to ensure we always stay under the throttle limit
        $secondsToSleep = 61 - ((Get-Date) - $global:apiCalls[0]).seconds
        Write-Verbose "Sleeping $secondsToSleep seconds..."
        
        # sleep determined number of seconds
        Start-Sleep -Seconds $secondsToSleep

        # resume processing normally
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
    }
    else {
        return $response.payload
    }
    
}

function New-SnipeItLocation($snipeitToken, $location) {
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
            name = $location
        } | ConvertTo-Json
    }
       
    # make api call to add location   
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        Write-Error $error[0].Exception
    }
    else {
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
        $response.messages
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
       
    # make api call to add manufacturer   
    $response = Invoke-RestMethod @irmParams
    Invoke-SnipeItApiThrottleManagement

    if ($response.Status -ne 'success') {
        $response.messages
    }
    else {
        return $response.payload
    }
}

function New-SnipeItAssetCheckin($snipeItToken, $assetId, $userId) {
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
        Write-Error 'Unable to checkout asset'
        $response.messages
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
        Write-Error 'Unable to add Model'
        $response.messages
    }
    
}

function Update-SnipeItAssets ($snipeitToken, $intuneDevices) {

    # loop through all intune devices and process them in 
    # to more usable data
    foreach ($device in $intuneDevices) {
        
        # clear re-used variables
        $snipeItUser = $null
        $snipeItDevice = $null
        $manufacturer = $null
        $snipeitModel = $null
        $snipeitLocation = $null

        
        Write-Verbose "Processing $($device.deviceName)"
        
        ##################
        ## Validate User #######################################################
        
        # match the intune user to a snipeit user for asset assgnment
        $snipeItUser = Get-SnipeItData -snipeitToken $snipeitToken -apiEndpoint 'users' -searchQuery $device.EmailAddress
        
        if (-not $snipeItUser) {
            # unable to match user, need to run an ldap sync

        }
        Write-Verbose "User Identified: $($snipeItUser.name)"
        ## END Validate User ###################################################
        ########################################################################

        ###############################
        ## Validate Device in Snipeit ###################################
        ## check if device has a serial number listed
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
            Write-Verbose "Device does NOT exist, creating..."
            # device does not exist, create it
            # verify manufacturer or create

            ##########################################
            ## BEGIN Validate or create Manufacturer ###########################
            $gsidParams = @{
                snipeitToken = $snipeitToken
                apiEndpoint  = 'manufacturers'
                searchQuery  = $device.manufacturer
            }

            # query snipeit for manufacturer
            $manufacturer = Get-SnipeItData @gsidParams

            # fix to match manufacturer by name if multiple matches exist
            if ($manufacturer.count -gt 1) {
                $manufacturer = $manufacturer | Where-Object {
                    $_.name -eq $device.manufacturer
                }
            }
            
            # if not found
            if (-not $manufacturer) {
                Write-Verbose "Manufacturer: $($device.manufacturer) not found, creating..."
                # manufacturer doesn't exist, create it
                New-SnipeItManufacturer -snipeitToken $snipeitToken -manufacturer $device.manufacturer
                $manufacturer = Get-SnipeItData @gsidParams

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
            Write-Verbose "Manufacturer: $($device.manufacturer) exists, checking model..."
                
            $gsidParams = @{
                snipeitToken = $snipeitToken
                apiENDpoint  = 'models'
                searchQuery  = $device.model
            }
            $snipeitModel = Get-SnipeItData @gsidParams

            # if model isn't found create it
            if (-not $snipeitModel) {
                Write-Verbose "Model: $($device.model) not found, creating..."
                
                # model doesn't exist, create it
                $model = @{
                    Name            = $device.model
                    Model           = $device.model
                    category_id     = '2'
                    manufacturer_id = $manufacturer.id
                    fieldset_id     = '1' 
                }
                New-SnipeItModel -snipeitToken $snipeitToken -intuneModel $model
                $snipeitModel = Get-SnipeItData @gsidParams
            }
            ## END Valide or create model ######################################
            ####################################################################
            
            ######################################
            ## BEGIN Validate or create location ###############################
            Write-Verbose "Model exists, checking location..."
            
            # get user location from Azure Graph
            $gaulParams = @{
                graphToken = $graphToken
                azureUpn   = $device.userPrincipalName
            }
            $userLocation = (Get-AzureUserLocation @gaulParams).officeLocation

            # check snipeit to see if location exists
            $gsidParams = @{
                snipeitToken = $snipeitToken
                apiEndpoint  = 'locations'
                searchQuery  = $userLocation
            }
            $snipeitLocation = Get-SnipeItData @gsidParams

            # if model isn't found create it
            if (-not $snipeitLocation) {
                Write-Verbose "Location: $userLocation not found, creating..."
                
                # model doesn't exist, create it
                New-SnipeItLocation -snipeitToken $snipeitToken -location $userLocation
                $snipeitLocation = Get-SnipeItData @gsidParams
            }
            ## END Validate or create location #################################
            ####################################################################

            ###########################
            ## BEGIN Create new asset ##########################################
            # model exists create asset
            New-SnipeItAsset -snipeitToken $snipeitToken -asset @{
                model_id        = $snipeitModel.id
                name            = $device.deviceName
                serial          = $device.serialNumber
                last_audit_date = $device.lastSyncDateTime
                location_id     = $snipeitLocation.id
}
            
            # get id for new asset
            $snipeItDevice = Get-SnipeItData -snipeitToken $snipeitToken -apiEndpoint 'hardware' -searchQuery $device.deviceName
            
            Write-Verbose "New Asset:"
            $snipeItDevice
            
            ## END Create new asset ############################################
            ####################################################################
        }
        
        Write-Verbose "Devices exists, verifying user..."
        
        # user snipeit shows currently assigned
        $snipeitUser = $snipeItDevice.assigned_to

        # user location from azure graph
        $intuneLocation = Get-AzureUserLocation -graphToken $graphToken -azureUpn $device.userPrincipalName

        # compare intune email address to snipeit user
        if ($device.emailAddress -ne $snipeitUser.username) {
            # if user doesn't match
            # need to re-assign asset in snipeit
            New-SnipeItAssetCheckin -snipeItToken $snipeitToken -assetId            

            # unsassign asset and leave note

            # assign asset and leave note
        }

        # compare intune user location to snipeit asset location
        if ($intuneLocation -ne $snipeItDevice.location.name) {
            # if location doesn't match
            # verify location exists
                
            # update asset location
        }
        ## END If Device does exist ########################################
        ####################################################################

        #  checkout asset if it isn't assigned to a user
        if (-not $snipeItDevice.assigned_to) {
            New-SnipeItAssetCheckout -snipeItToken $snipeitToken -assetId $snipeItDevice.id -userId $snipeItUser.id
            
        }
        else {
            Write-Verbose 'Device already assigned...'
        }
    }
}


<#
Properties to provide SnipeIt
 company            - Company                                      
 serial             - serial number
 model              - model (assumes Manufacturer)
 status_label       - Status (set to "In Production")
 name               - Device Name
 intuneEnroll       - Purchase Date (enrolled Date/time)
 assigned_to        - Assigned User                                
 mode/Manufacturer  - Manufacturer (only if model isn't present)
 lastSyncDateTime   - 

#>
<#
To Do:
- get jamf import working
- recurring schedule with proper credential storage
- handle override or not overriding a user assignment that doesn't match
- setup a dedicated snipeit user so it doesn't reflect my user account checking
    out all assets

    
#>

<#
Questions:
1) What should I enter for an asset location?
    pull from user logon location's or?

2) Do we want to override asset assignments in snipeit with what's in Intune?
#>

$VerbosePreference = 'Continue'
[System.DateTime[]]$global:apiCalls = @()

#################################################
# prepare tokens for Graph and Snipeit
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

###################################################


# retrieve intune devices with graph token
$intuneDevices = Get-ManagedDevices($graphToken)

# retrieve all assets
#$snipeItAssets = Get-SnipeItData -snipeitToken $snipeItToken -apiEndpoint 'hardware'

