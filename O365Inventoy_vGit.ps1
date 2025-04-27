<#

Version: POWERSHELL 7

# List all installed versions Microsoft.Graph
Get-InstalledModule -Name Microsoft.Graph -AllVersions
# Uninstall all old versions Microsoft.Graph
Get-InstalledModule -Name Microsoft.Graph -AllVersions | Uninstall-Module -Force
# Install recently version Microsoft.Graph
Install-Module -Name Microsoft.Graph -Scope CurrentUser -AllowClobber -Force

$PSVersionTable

#>

<# Parameters #>
<# Demo Tenant info #>

$tenantId = "<tenant id>"
$clientId = "<client id>"
$clientSecret = "<client secret id>"
$organizationDomain="<domain>"

$uploadSharePointSameOrAlternative="Alternative" #Same or Alternative options supported
$LocalFolderInventory="<path>"

<#To generate token from alternative tenant to upload CSV to SharePoint library#>
$tenantId_AlternativeUpload = "<alternative tenant id>"
$clientId_AlternativeUpload = "<alternative client id>"
$clientSecret_AlternativeUpload = "<alternative secret id>"

#Upload File Shp
$siteNameUploadFileShp = "<Shp site name>"
$global:siteId = "" 
$libName ="Documentos" 
$global:driveId = "" 

#filter Parameters

$totalDayAuditLog=180 #filter AuditLog
$PeriodShpAndOneDriveUsageReport='D180' #filter Usage Report SharePoint and OneDrive-  D7, D30, D90, D180
$TopMailFoldersMessages=5000 #filter to Exchange mail messages
$periodMailBoxUsageReport = 'D180'  # Período pode ser D7, D30, D90, ou D180
<#End Parameters#>
<#
    Get site id and drive id to upload files to Shp. 
    After run this function you must take note this IDs to insert at respectives parameters above
#>
function getDriveAndSiteId{
    $token=""
    if($uploadSharePointSameOrAlternative -eq "Alternative"){
        $token=getTokenGraphAlternative
    }else {
        $token=getTokenGraph
    }
    $siteResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites?search=$siteNameUploadFileShp" -Headers @{Authorization = "Bearer $token"}
    $global:siteId = $siteResponse.value[0].id.ToString().Split(',')[1] 
    $driveResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites/$global:siteId/drives" -Headers @{Authorization = "Bearer $token"}
    Write-Host "Site ID: $($global:siteId)" -ForegroundColor Green 
    foreach ($drive in $driveResponse.value) {
        Write-Host "Drive ID: $($drive.id), Drive Name: $($drive.name)" -ForegroundColor Green
        if($libName -eq $drive.name){
            $global:driveId=$drive.Id
        }
    }
    #Write-Output "Drive ID: $driveId" 
}
<#upload files inventory to SharePoint#>
function UploadFileShp{
    param (
        [string]$filePath
        <#[string]$token,
        [hashtable]$headers#>)
    if($uploadSharePointSameOrAlternative -eq "Same"){
        $token=getTokenGraph
    }else{
        $token=getTokenGraphAlternative
    }
    $headers = @{
        "Authorization" = "Bearer $token"
        "ConsistencyLevel" = "eventual"
    }
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $fileContent = [System.IO.File]::ReadAllBytes($filePath)
    $fileSize = $fileContent.Length
    $startByte = 0
    $endByte = $fileSize - 1
    $headersUpload = @{
        'Authorization' = "Bearer $token"
        'Content-Range' = "bytes $startByte-$endByte/$fileSize"
    }
    $uploadSession = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/sites/$global:siteId/drives/$global:driveId/root:/$($fileName):/createUploadSession" -Headers $headers -ContentType "application/json" -Body (@{} | ConvertTo-Json)
    $uploadUrl = $uploadSession.uploadUrl
    Invoke-RestMethod -Method Put -Uri $uploadUrl -Headers $headersUpload -Body $fileContent -ContentType "application/octet-stream"
}
function Get-ConcatenatedValue {
    param (
        [Parameter(Mandatory=$true)]
        [pscustomobject]$psObject
    )
    $values = @()
    foreach ($property in $psObject.PSObject.Properties) {
        $value = $property.Value
        if ($value -is [string]) {
            #replace line break by blank spaces
            $value = $value -replace "`n", " " -replace "`r", " "
        }
        $values += "$($property.Name)=$value"
    }
    return $values -join ", "
}
<# Get token from alternative tenant to upload CSV to SharePoint library #>
function getTokenGraphAlternative{
   
    $body=@{}
    $resource = "https://graph.microsoft.com"
    $body = @{
        client_id     = $clientId_AlternativeUpload
        scope         = "$resource/.default"
        client_secret = $clientSecret_AlternativeUpload
        grant_type    = "client_credentials"
    }
    $authUrl = "https://login.microsoftonline.com/$tenantId_AlternativeUpload/oauth2/v2.0/token"
    $response = Invoke-RestMethod -Method Post -Uri $authUrl -ContentType "application/x-www-form-urlencoded" -Body $body
    $token = $response.access_token
    if (-not $token) {
        Write-Error "Failed to obtain access token."
        return $null
    }
    return $token
}
function getTokenGraph{
    param(
        [string]$BodyType = ""
    )
    $body=@{}
    if($BodyType -eq "powerbi"){
        $body = @{
            grant_type    = "client_credentials"
            scope         = "https://analysis.windows.net/powerbi/api/.default"
            client_id     = $clientId
            client_secret = $clientSecret
        }
    }elseif($BodyType -eq "exchange"){
        $body = @{
            grant_type    = "client_credentials"
            scope         = "https://outlook.office365.com/.default"
            client_id     = $clientId
            client_secret = $clientSecret
        }
    }
    else{
        $resource = "https://graph.microsoft.com"
        $body = @{
            client_id     = $clientId
            scope         = "$resource/.default"
            client_secret = $clientSecret
            grant_type    = "client_credentials"
        }
    }
    $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $response = Invoke-RestMethod -Method Post -Uri $authUrl -ContentType "application/x-www-form-urlencoded" -Body $body
    $token = $response.access_token
    if (-not $token) {
        Write-Error "Failed to obtain access token."
        return $null
    }
    return $token
}
function getShpOneDriveInventory{
    Write-Host "Initalizing SharePoint and OneDrive Inventory" -ForegroundColor Green
    $token = getTokenGraph
    #$secureAccessToken = ConvertTo-SecureString -String $token -AsPlainText -Force
    #Connect-MgGraph -AccessToken $secureAccessToken  
    $headers = @{
        "Authorization" = "Bearer $token"
        "ConsistencyLevel" = "eventual"
    }
    #API Permission: Graph - Sites.Read.All
    $sites = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites" -Headers $headers
    #API Permission: Graph - Reports.Read.All
    $usageResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='$($PeriodShpAndOneDriveUsageReport)')" -Headers $headers
    # Periods: D7, D30, D90, ou D180
    $usageDataShp = $usageResponse | ConvertFrom-Csv
    foreach ($usage in $usageDataShp) {
        #$usage | Add-Member -MemberType NoteProperty -Name SiteId -Value $site.id
        $storageGB=[math]::Round($usage.'Storage Used (Byte)' / 1GB, 2)
        $usage | Add-Member -MemberType NoteProperty -Name 'Storage Used (GB)' -Value $storageGB
    }
      
    Write-Host "There are " $sites.value.Count " site collections present"
    function Get-SiteOwnersAndAdmins {
        param (
            [string]$siteId,
            [string]$token
        )
        $headers = @{
            "Authorization" = "Bearer $token"
            "Content-Type"  = "application/json"
        }
        #API Permission: Graph - Sites.Manage.All
        <#$permissions = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions" -Headers $headers
        $admins = $permissions.value | Where-Object { $_.roles -contains "owner" }
        $ownerDisplayNames = @()
        foreach ($permission in $permissions.value) {
            if ($permission.roles -contains "owner") {
                foreach ($grantee in $permission.grantedTo) {
                    $ownerDisplayNames += $grantee.user.displayName
                }
            }
        }#>
        $owners = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drive" -Headers $headers
        $ownerDisplayNames = @()
        if ($owners.owner.user) {
            foreach ($user in $owners.owner.user) {
                $ownerDisplayNames += $user.displayName
            }
        }
        if ($owners.owner.group) {
            foreach ($group in $owners.owner.group) {
                $ownerDisplayNames += $group.displayName
            }
        }
        return [PSCustomObject]@{
            SiteId = $siteId
            Owners = ($ownerDisplayNames -join ", ")#$owners.owner.user.displayName
            #Admins = ($admins.grantedTo.user.displayName -join ", ")
        }
    }
    <#$ownersAndAdminsData = @()

    foreach ($site in $sites.value) {
        if($site.displayName -ne ""){
            Write-Host "$($site.DisplayName) - $($site.Id)" 
            $siteInfo = Get-SiteOwnersAndAdmins -siteId $site.id -token $token
            $ownersAndAdminsData += [PSCustomObject]@{
                SiteId = "`"$($site.id)`"" 
                SiteName = $site.displayName
                WebUrl = $site.webUrl
                Owners = $siteInfo.Owners
                #Admins = "`"$($siteInfo.Admins)`"" 
            }
        }
    }
    $outputpathOwnersAndAdmins = $LocalFolderInventory + "\Graph-ShpSitesOwnersAndAdmins.csv"
    $ownersAndAdminsData | Export-Csv $outputpathOwnersAndAdmins -NoTypeInformation -Encoding Unicode
    UploadFileShp -filePath $outputpathOwnersAndAdmins#>
    $usageResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='$($PeriodShpAndOneDriveUsageReport)')" -Headers $headers
    # Periods: D7, D30, D90, ou D180
    $usageDataOneDrive = $usageResponse | ConvertFrom-Csv
    foreach ($usage in $usageDataOneDrive) {
        $storageGB=[math]::Round($usage.'Storage Used (Byte)' / 1GB, 2)
        $usage | Add-Member -MemberType NoteProperty -Name 'Storage Used (GB)' -Value $storageGB
    }
    $outputpath = $LocalFolderInventory+"\Graph-ShpSites.csv"    
    $sites.value | Select-Object * | Export-Csv $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers

    $outputpath = $LocalFolderInventory+"\Graph-ShpSitesUsage.csv"
    $usageDataShp | Export-Csv $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers

    $outputpath = $LocalFolderInventory+"\Graph-OneDriveUsage.csv"
    $usageDataOneDrive | Export-Csv $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers

    Write-Host "Finish SharePoint and OneDrive Inventory" -ForegroundColor Green
}
function getUsersAndLicensesInventory{
    Write-Host "Initalizing Users Inventory" -ForegroundColor Green
    $token = getTokenGraph
    $headers = @{
        'Authorization' = "Bearer $token"
        'Content-Type'  = 'application/json'
    }
    #API Permission: Graph - User.Read.All
    $url = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,createdDateTime,accountEnabled,assignedLicenses,businessPhones,city,companyName,country,department,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,postalCode,preferredLanguage,state,streetAddress,surname,userPrincipalName"
    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
    $users = $response.value
    $userLicenseData = @()

    foreach ($user in $users) {
        $userid=$user.id
        $url = "https://graph.microsoft.com/v1.0/users/$userid/memberOf"
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        $groups = $response.value
        $MembersOf= ($groups | ForEach-Object { $_.displayName }) -join ", "
        #API Permission: Graph - User.Read.All
        $licenseUrl = "https://graph.microsoft.com/v1.0/users/$($user.Id)/licenseDetails"
        $licenseResponse = Invoke-RestMethod -Uri $licenseUrl -Headers $headers -Method Get
        $licenses = $licenseResponse.value
        
        if ($licenses.Count -eq 0) {
            # Add user unlicensed
            $userLicenseData += [PSCustomObject]@{
                ID = $user.Id
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                AccountEnabled = $user.accountEnabled
                CreatedDateTime = $user.createdDateTime
                LicenseSkuId = $null
                ServicePlanName = $null
                ProvisioningStatus = $null
                AppliesTo = $null
                ServicePlanType = $null
                MembersOf = $MembersOf
                Department = $user.department
                JobTitle = $user.jobTitle
                MobilePhone = $user.mobilePhone
                OfficeLocation = $user.officeLocation
                PostalCode = $user.postalCode
            }
        } else {
            $licenses | ForEach-Object {
                $skuId = $_.SkuId
                $_.ServicePlans | ForEach-Object {
                    $userLicenseData += [PSCustomObject]@{
                        ID = $user.Id
                        DisplayName = $user.DisplayName
                        UserPrincipalName = $user.UserPrincipalName
                        AccountEnabled = $user.accountEnabled
                        CreatedDateTime = $user.createdDateTime
                        LicenseSkuId = $skuId
                        ServicePlanName = $_.ServicePlanName
                        ProvisioningStatus = $_.ProvisioningStatus
                        AppliesTo = $_.appliesTo
                        ServicePlanType = $_.servicePlanType
                        MembersOf = $MembersOf
                        Department = $user.department
                        JobTitle = $user.jobTitle
                        MobilePhone = $user.mobilePhone
                        OfficeLocation = $user.officeLocation
                        PostalCode = $user.postalCode
                    }
                }
            }
        }
        <#$licenses | ForEach-Object {
            $userLicenseData += [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                LicenseSkuId = $_.SkuId
                ServicePlanName= ($_.ServicePlans | ForEach-Object { $_.ServicePlanName }) -join ", "
                ProvisioningStatus= ($_.ServicePlans | ForEach-Object { $_.provisioningStatus }) -join ", "
                AppliesTo= ($_.ServicePlans | ForEach-Object { $_.appliesTo }) -join ", "
                ServicePlans = ($_.ServicePlans | ForEach-Object { $_.ServicePlanId }) -join ", "
            }
        }#>
    }
    <#$headers = @{
        "Authorization" = "Bearer $token"
        "ConsistencyLevel" = "eventual"
    }#>
    #users deleted
    $url = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?`$select=id,userPrincipalName,displayname,mail,deletedDateTime "
    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
    $usersdeleted = $response.value
    
    #API Permission: Graph - LicenseAssignment.Read.All
    $url = "https://graph.microsoft.com/v1.0/subscribedSkus"
    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
    $subscribedSkus = $response.value

    $outputpath = $LocalFolderInventory + "\Graph-SubscribedSkus.csv"
    $subscribedSkus | Export-Csv -Path $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers

    $outputpath = $LocalFolderInventory + "\Graph-UsersAndLicensesInventory.csv"
    $userLicenseData | Export-Csv -Path $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers

    $outputpath = $LocalFolderInventory + "\Graph-UsersDeletedInventory.csv"
    $usersdeleted | Export-Csv -Path $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    Write-Host "Finish Users Inventory" -ForegroundColor Green
}
function getMailboxesInventory{
    Write-Host "Initializing Mailboxes Inventory" -ForegroundColor Green
    $token=getTokenGraph
    $headers = @{
        'Authorization' = "Bearer $token"
        'Content-Type'  = 'application/json'
    }
    #API Permission: Graph - User.Read.All
    $url = "https://graph.microsoft.com/v1.0/users"
    
    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
    $users = $response.value
    $mailboxData = @()
    $mailboxMessages = @()
    $mailboxSize = @()

    foreach ($user in $users) {
        #Write-Host $user.displayName
        $isExternal = if ($user.UserPrincipalName -notlike "*@$organizationDomain") { 
            $true 
        } else {
             $false
        }
        if(-not $isExternal){
            #$totalMailboxSize = 0
            #API Permission: Graph - Mail.Read
            $foldersUrl = "https://graph.microsoft.com/v1.0/users/$($user.Id)/mailFolders"
            $foldersResponse = Invoke-RestMethod -Uri $foldersUrl -Headers $headers -Method Get
            $folders = $foldersResponse.value
        
            foreach ($folder in $folders) {
                try{
                    #API Permission: Graph - Mail.Read
                    $folderSizeUrl = "https://graph.microsoft.com/v1.0/users/$($user.Id)/mailFolders/$($folder.Id)/messages?$top=$($TopMailFoldersMessages)"
                    $folderSizeResponse = Invoke-RestMethod -Uri $folderSizeUrl -Headers $headers -Method Get
                }catch{
                    Write-Host "Mailbox problem (soft-deleted, inactive, etc): $($user.DisplayName)" -ForegroundColor DarkMagenta
                    continue
                }
                $messages = $folderSizeResponse.value                
              
                foreach ($message in $messages) {
                    $messagesUsers = [PSCustomObject]@{}
                    foreach ($property in $message.PSObject.Properties) {
                        if ($null -ne $property.Value -and $property.Name -ne "Body" -and $property.Name -ne "bodyPreview") {
                            if ($property.Value -is [System.Object[]]) {
                                $formattedValues = $property.Value | ForEach-Object {
                                    if ($_ -is [PSCustomObject]) {
                                        $jsonObject = $_ | ConvertTo-Json -Compress | ConvertFrom-Json
                                        ($jsonObject.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join "; "
                                    } else {
                                        $_
                                    }
                                }
                                $concatenatedValue = ($formattedValues -join ", ")
                                $messagesUsers | Add-Member -MemberType NoteProperty -Name $property.Name -Value $concatenatedValue
                            } 
                            elseif($property.Value -is [PSCustomObject]){
                                $jsonValue = $property.Value | ConvertTo-Json -Compress
                                $messagesUsers | Add-Member -MemberType NoteProperty -Name $property.Name -Value $jsonValue 
                            }
                            else {
                                $messagesUsers | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value 
                            }
                        }else{
                            $messagesUsers | Add-Member -MemberType NoteProperty -Name $property.Name -Value ""
                        }
                    }
                    $mailboxMessages += $messagesUsers
                }
            }        
            #$mailboxSize = $mailboxResponse.storageQuota
            $userDetails = [PSCustomObject]@{
                #TotalMailboxSize = $totalMailboxSize
            }
            # Adicionar todas as propriedades do usuário ao objeto
            foreach ($property in $user.PSObject.Properties) {
                if ($null -ne $property.Value) {
                    if ($property.Value -is [System.Object[]]) {
                        $concatenatedValue = ($property.Value -join ", ")
                        $userDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $concatenatedValue
                    } else {
                        $userDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
                    }
                }else{
                    $userDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value ""
                }
            }
            try{ 
                #API Permission: Graph - MailBoxSettings.Read
                $mailboxUrl = "https://graph.microsoft.com/v1.0/users/$($user.Id)/mailboxSettings"
                $mailboxResponse = Invoke-RestMethod -Uri $mailboxUrl -Headers $headers -Method Get
                foreach ($property in $mailboxResponse.PSObject.Properties) {
                    if ($null -ne $property.Value) {
                        if ($property.Value -is [System.Object[]]) {
                            $concatenatedValue = ($property.Value -join ", ")
                            $userDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $concatenatedValue
                        } elseif ($property.Value -is [pscustomobject]) {
                            $concatenatedValue = Get-ConcatenatedValue -psObject $property.Value
                            $userDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $concatenatedValue
                        }else {
                            $userDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
                        }
                    }else{
                        $userDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value ""
                    }
                }
            }catch{
                Write-Host "Mailbox problem (soft-deleted, inactive, etc): $($user.DisplayName)" -ForegroundColor DarkMagenta
            }        
            $mailboxData += $userDetails
        }else {
            $userDetails = [PSCustomObject]@{}
            foreach ($property in $user.PSObject.Properties) {
                if ($null -ne $property.Value) {
                    $userDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
                }
            }
            $mailboxData += $userDetails
        }
    }
    #API Permission: Graph - Mail.Read
    $url = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$periodMailBoxUsageReport')"    
    $mailboxSize = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ResponseHeadersVariable responseHeaders

    <#$headers = @{
        "Authorization" = "Bearer $token"
        "ConsistencyLevel" = "eventual"
    }#>
    $outputpath = $LocalFolderInventory + "\Graph-MailboxesStorageInventory.csv"
    $mailboxSize | ConvertFrom-Csv | Export-Csv -Path $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    
    $outputpath = $LocalFolderInventory + "\Graph-MailboxesMessagesInventory.csv"
    $mailboxMessages | Export-Csv -Path $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    $outputpath = $LocalFolderInventory + "\Graph-MailboxesInventory.csv"
    $mailboxData | Export-Csv -Path $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    Write-Host "Finish Mailboxes Inventory" -ForegroundColor Green
}
function getGroupsInventory{
    Write-Host "Initializing Groups Inventory" -ForegroundColor Green
    $token=getTokenGraph
    $headers = @{
        Authorization = "Bearer $token"
    }
    #API Permission: Graph - Group.Read.All
    $groups = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups" -Headers $headers -Method Get
    
    $allGroups = @($groups.value)
    while ($groups.'@odata.nextLink') {
        $groups = Invoke-RestMethod -Uri $groups.'@odata.nextLink' -Headers $headers -Method Get
        $allGroups += $groups.value
    }
    $groupsToExport = @()
    foreach($g in $allGroups){
        $groupsDetails = [PSCustomObject]@{}
        foreach ($property in $g.PSObject.Properties) {
            if ($null -ne $property.Value) {
                if ($property.Value -is [System.Object[]]) {
                    $concatenatedValue = ($property.Value -join ", ")
                    $groupsDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $concatenatedValue
                } else {
                    $groupsDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
                }
            }else {
                $groupsDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value ""
            }
        }
        $groupType = "Unknown"
        if ($g.groupTypes -contains "Unified") {
            $groupType = "Microsoft 365 Group"
            $groupsDetails | Add-Member -MemberType NoteProperty -Name "GroupTypeCustom" -Value $groupType
        } elseif ($g.mailEnabled -eq $true -and $g.securityEnabled -eq $true) {
            $groupType = "Mail-enabled Security Group"
            $groupsDetails | Add-Member -MemberType NoteProperty -Name "GroupTypeCustom" -Value $groupType
        } elseif ($g.mailEnabled -eq $true -and $g.securityEnabled -eq $false) {
            $groupType = "Distribution Group"
            $groupsDetails | Add-Member -MemberType NoteProperty -Name "GroupTypeCustom" -Value $groupType
        } elseif ($g.mailEnabled -eq $false -and $g.securityEnabled -eq $true) {
            $groupType = "Security Group"
            $groupsDetails | Add-Member -MemberType NoteProperty -Name "GroupTypeCustom" -Value $groupType
        }
        #Write-Output "$($g.displayName): $groupType"
        #API Permission: Graph - Group.Read.All
        $members = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/members" -Headers $headers -Method Get
        $memberNames = @()
        foreach ($member in $members.value) {
            $memberNames += $member.displayName
        }
        $groupsDetails | Add-Member -MemberType NoteProperty -Name "Members" -Value ($memberNames -join ", ")
    
        $groupsToExport += $groupsDetails
    }  
    <#$headers = @{
        "Authorization" = "Bearer $token"
        "ConsistencyLevel" = "eventual"
    }#>
    $outputpath = $LocalFolderInventory + "\Graph-GroupsInventory.csv"
    $groupsToExport | Export-Csv -Path $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    Write-Host "Finish Groups Inventory" -ForegroundColor Green
}
function getTeamsInventory{
    Write-Host "Initializing Teams Inventory" -ForegroundColor Green
    $token=getTokenGraph
    $headers = @{
        Authorization = "Bearer $token"
    }    
    #API Permission: Graph - Team.ReadBasic
    $teams = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/teams" -Headers $headers -Method Get
    $teamsDataExport = @()    
    foreach ($team in $teams.value) {
        $teamId = $team.id
        $teamName = $team.displayName
    
        #API Permission: Graph - Channel.ReadBasic.All
        $channels = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels" -Headers $headers -Method Get
        
        #API Permission: Graph - Group.Read.All
        $owners = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$teamId/owners" -Headers $headers -Method Get
        $ownerNames = ($owners.value | ForEach-Object { $_.displayName }) -join ", "
        
        #API Permission: Graph - TeamMember.Read.All
        $members = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/members" -Headers $headers -Method Get
        $memberNames = ($members.value | ForEach-Object { $_.displayName }) -join ", "
    
        foreach ($channel in $channels.value) {
            $channelId = $channel.id
            $channelName = $channel.displayName
            $channelMembers = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$channelId/members" -Headers $headers -Method Get
            $channelMemberNames = ($channelMembers.value | ForEach-Object { $_.displayName }) -join ", "
    
            $teamsData = [PSCustomObject]@{
                TeamName         = $teamName
                TeamId           = $teamId
                ChannelName      = $channelName
                ChannelId        = $channelId
                Owners           = $ownerNames
                TeamMembers      = $memberNames
                ChannelMembers   = $channelMemberNames
            }
    
            foreach ($property in $channel.PSObject.Properties) {
                if ($null -ne $property.Value) {
                    if ($property.Value -is [System.Object[]]) {
                        $concatenatedValue = ($property.Value -join ", ")
                        $teamsData | Add-Member -MemberType NoteProperty -Name $property.Name -Value $concatenatedValue
                    } else {
                        $teamsData | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
                    }
                } else {
                    $teamsData | Add-Member -MemberType NoteProperty -Name $property.Name -Value ""
                }
            }
    
            $teamsDataExport += $teamsData
        }
    }    
    <#$headers = @{
        "Authorization" = "Bearer $token"
        "ConsistencyLevel" = "eventual"
    }#>
    $outputpath = $LocalFolderInventory+"\Graph-TeamsInventory.csv"
    $teamsDataExport | Export-Csv $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    Write-Host "Finish Teams Inventory" -ForegroundColor Green
}
function getPowerBIInventory{
   <#Enable required in Power BI:
    go to Power BI portal (app.powerbi.com).
    browse to "Admin portal" > "Tenant settings".
    "Developer settings", enable "Allow service principals to use Power BI APIs"
    https://learn.microsoft.com/en-us/power-bi/developer/embedded/embed-service-principal?tabs=azure-portal
    .#>
    Write-Host "Initializing Power BI Inventory" -ForegroundColor Green
    $token=getTokenGraph -BodyType "powerbi"
    $headers = @{
        Authorization = "Bearer $token"
    }
    <#Scan#>
    #$workspaceScanRequest = @{
    #    workspaces = @()
    #}
    #$scanResponse = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo" -Headers $headers -Method Post -Body ($workspaceScanRequest | ConvertTo-Json) -ContentType "application/json"
    <#Scan#>
    
    <#
    When a user is granted permissions to a workspace, app, or Power BI item (such as a report or a dashboard), the new permissions might not be immediately available through API calls. This operation refreshes user permissions to ensure they're fully update
    https://learn.microsoft.com/en-us/rest/api/power-bi/users/refresh-user-permissions
    #>
    #$retultInvoke= Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions" -Headers $headers -Method Post
    
    <# 
    Get all workspaces that current user can access
    https://learn.microsoft.com/en-us/rest/api/power-bi/groups/get-groups
    #>
    $workspaces = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/groups" -Headers $headers -Method Get #-Debug
      
    #$workspaces=Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/groups" -Headers $headers 
    
    $powerBIData = @()
    
    foreach ($workspace in $workspaces.value) {
        $workspaceId = $workspace.id
        $workspaceName = $workspace.name
    
        $reports = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/reports" -Headers $headers -Method Get
    
        $users = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/users" -Headers $headers -Method Get
        $userNames = ($users.value | ForEach-Object { $_.displayName }) -join ", "
    
        $owners = $users.value | Where-Object { $_.groupUserAccessRight -eq "Admin" }
        $ownerNames = ($owners | ForEach-Object { $_.displayName }) -join ", "
    
        foreach ($report in $reports.value) {
            $reportId = $report.id
            $reportName = $report.name
    
            $reportUsers = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/reports/$reportId/users" -Headers $headers
            $reportUserNames = ($reportUsers.value | ForEach-Object { $_.displayName }) -join ", "
    
            $reportOwners = $reportUsers.value | Where-Object { $_.reportUserAccessRight -eq "Owner" }
            $reportOwnerNames = ($reportOwners | ForEach-Object { $_.displayName }) -join ", "
    
            $powerBIData += [PSCustomObject]@{
                WorkspaceName = $workspaceName
                WorkspaceId   = $workspaceId
                ReportName    = $reportName
                ReportId      = $reportId
                WorkspaceUsers = $userNames
                WorkspaceOwners = $ownerNames
                ReportUsers   = $reportUserNames
                ReportOwners  = $reportOwnerNames
            }
        }
    }
    
    $outputpath = $LocalFolderInventory+"\Graph-PowerBIInventory.csv"
    $powerBIData | Export-Csv $outputpath -NoTypeInformation -Encoding unicode
    $token=getTokenGraph
    <#$headers = @{
        "Authorization" = "Bearer $token"
        "ConsistencyLevel" = "eventual"
    }#>
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    Write-Host "Finish Power BI Inventory" -ForegroundColor Green
}
function getPowerBIInventoryMGMT{
    #Power Bi Inventory
    #https://learn.microsoft.com/en-us/rest/api/power-bi/admin
    #Install-Module -Name MicrosoftPowerBIMgmt
    #Connect-AzureAD -AadAccessToken getTokenGraph
    $securePassword = ConvertTo-SecureString $clientSecret -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($clientId, $securePassword)
    Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId $tenantId 
    $workspaces = Get-PowerBIWorkspace #-All
    foreach ($workspace in $workspaces) {
        Write-Output "Workspace: $($workspace.Name)"
        $reports = Get-PowerBIReport -WorkspaceId $workspace.Id -Scope Organization
        $users= $workspace.Users
        foreach($user in $users){
            Write-Host "User: $($user.Identifier), $($user.AccessRight)"
        }
        foreach ($report in $reports) {
            Write-Output "  Report: $($report.Name)"
            $url = "https://api.powerbi.com/v1.0/myorg/admin/reports/$($report.Id)/users"
            $users = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
            foreach ($user in $users.value) {
                Write-Output "Report User: $($user.displayName), Access Level: $($user.reportUserAccessRight)"
            }
        }
}
}
function getAuditLogs_SearchUnifiedAuditLog{
    <#
    Audit logging has to be enabled for your organization to successfully use the script to return audit records.
    Get-AdminAuditLogConfig | FL UnifiedAuditLogIngestionEnabled
    https://learn.microsoft.com/en-us/purview/audit-log-search-script
    #>
    Write-Host "Initializing Audit Logs Inventory" -ForegroundColor Green
    #Install-Module -Name ExchangeOnlineManagement Scope AllUsers -Force
    #Enable-OrganizationCustomization
    #Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
    #Import-module ExchangeOnlineManagement
    #Remove-Module -Name ExchangeOnlineManagement -Force
    #Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.7.1
    Import-Module -Name ExchangeOnlineManagement -RequiredVersion 3.7.1 -Force
    $token=getTokenGraph -BodyType "exchange"
    <#Permissions needed:
        App: Office 365 Exchange Online - Exchange.ManageAsApp
        Roles: Global Reader
    #>
    <#
        API Permission: Office 365 Exchange Online - Exchange.ManageAsApp
        Access from all APIS screen
    #>
    #API Permission: Office 365 Exchange Online: Exchange.ManageAsApp 
    #Azure role to app : Exchange Administrator
    Connect-ExchangeOnline -AccessToken $token -Organization $tenantId
    $startDate = (Get-Date).AddDays(-($totalDayAuditLog))
    $endDate = Get-Date
    #API Permission: Graph - AuditLog.Read.All
    #$auditLogs = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate

    $sessionId = [guid]::NewGuid().ToString()
    $auditLogs = @()
    do {
        $results = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -ResultSize 5000 -SessionId $sessionId -SessionCommand ReturnLargeSet
        $auditLogs += $results
    } while ($results.Count -eq 5000)
    $token=getTokenGraph
    <#$headers = @{
        "Authorization" = "Bearer $token"
        "ConsistencyLevel" = "eventual"
    }#>
    $outputpath = $LocalFolderInventory+"\AuditLogs.csv"
    $auditLogs | Export-Csv $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    Write-Host "Finish Audit Logs Inventory" -ForegroundColor Green
}
function getAuditLogs{
    Write-Host "Initializing Audit Logs Inventory" -ForegroundColor Green
    #Import-Module Microsoft.Graph.Reports
    #Get-MgAuditLogSignIn -Filter "(createdDateTime ge 2024-01-13T14:13:32Z and createdDateTime le 2024-01-14T17:43:26Z)" 
    $token=getTokenGraph
    $headers = @{
        Authorization = "Bearer $token"
    }
    $startDate = (Get-Date).AddDays(-($totalDayAuditLog)).ToString("yyyy-MM-ddTHH:mm:ssZ")
    #$endDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    #auditlog.read.all
    #https://learn.microsoft.com/pt-br/graph/api/resources/azure-ad-auditlog-overview?view=graph-rest-1.0
    #$auditLogs = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/auditLogs/signIns?startDateTime=$startDate&endDateTime=$endDate" -Headers $headers
    
    #API Permission: Graph - AuditLog.Read.All
    #This endpoint can a limitation to 30 days report
    $auditLogs = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?$filter=activityDateTime gt $($startDate)" -Headers $headers
    $outputpath = $LocalFolderInventory+"\Graph-AuditLog.csv"
    #$auditLogs.value | Select-Object * | Export-Csv $outputpath -NoTypeInformation -Encoding unicode
    $auditLogs.value | ForEach-Object {
        $item = $_
        $properties = @{}
        $item.PSObject.Properties | ForEach-Object {
            $propertyName = $_.Name
            $propertyValue = $_.Value
            if ($propertyValue -is [System.Object[]]) {
                $properties[$propertyName] = ($propertyValue | ForEach-Object {
                    if ($_ -is [PSCustomObject]) {
                        ($_.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join "; "
                    } else {
                        $_
                    }
                }) -join ", "
            } elseif ($propertyValue -is [PSCustomObject]) {
                $properties[$propertyName] = ($propertyValue.PSObject.Properties | ForEach-Object { 
                    "$($_.Name)=$($_.Value)" 
                }) -join "; "
            } else {
                $properties[$propertyName] = $propertyValue
            }
        }
        New-Object PSObject -Property $properties
    } | Export-Csv $outputpath -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpath #-token $token -headers $headers
    Write-Host "Finish Audit Logs Inventory" -ForegroundColor Green
}
function getHealthOverviews{
    Write-Host "Initializing Health Overviews Inventory" -ForegroundColor Green
    #Import-Module Microsoft.Graph.Reports
    #Get-MgAuditLogSignIn -Filter "(createdDateTime ge 2024-01-13T14:13:32Z and createdDateTime le 2024-01-14T17:43:26Z)" 
    $token=getTokenGraph
    $headers = @{
        Authorization = "Bearer $token"
    }
    $outputpathHealthOverviews = $LocalFolderInventory+"\Graph-HealthOverviews.csv"
    $outputpathIssues = $LocalFolderInventory+"\Graph-Issues.csv"
    #API Permission: Graph - ServiceHealth.Read.All
    $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews" -Headers $headers -Method Get
    $dataHealthOverviews=@()
    if ($response.value) {
        $dataHealthOverviews=$response.value     
    } else {
        Write-Output "No data found"
    }
     #API Permission: Graph - ServiceHealth.Read.All
     $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues" -Headers $headers -Method Get
     $dataIssues=@()
     if ($response.value) {
         $dataIssues=$response.value     
     } else {
         Write-Output "No data found"
     }
    $dataHealthOverviews | Export-Csv $outputpathHealthOverviews -NoTypeInformation -Encoding unicode
    $dataIssues | Export-Csv $outputpathIssues -NoTypeInformation -Encoding unicode
    UploadFileShp -filePath $outputpathHealthOverviews
    UploadFileShp -filePath $outputpathIssues
}
getDriveAndSiteId
getShpOneDriveInventory
getUsersAndLicensesInventory
getMailboxesInventory
getGroupsInventory
getTeamsInventory
getAuditLogs_SearchUnifiedAuditLog
getPowerBIInventory
getHealthOverviews

#another tests#
#getAuditLogs
#getPowerBIInventoryMGMT





