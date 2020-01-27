Import-Module 'MSAL.PS'

# Slack text width with the formatting we use maxes out ~80 characters...
$Width = 80
$CommandsToExport = @()

function Find-AADUser {
    <#
    .SYNOPSIS
        Find Azure AD User 
    .PARAMETER Identity displayName
        Azure AD displayName
    .USAGE
        !FindAADUser displayName
    .EXAMPLE
        !FindAADUser 'Darren Robinson'
        !FindAADUser 'darren.robinson@myazureaddomain.com'        
    .EXAMPLE
        !SearchAADUser 'Darren Robinson'
        !SearchAADUser 'darren.robinson@myazuredomain.com'
    .DESCRIPTION !FindUser
        Search AzureAD user objects using displayName
    .LINK
        https://blog.darrenjrobinson.com
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'FindAADUser',
        Aliases = ('SearchAADUser'),
        Permissions = 'read'
    )]
    param(      
        [parameter(position = 1,
            parametersetname = 'id', mandatory)]
        [string]$Identity,

        [PoshBot.FromConfig('AADCreds')]
        [parameter(mandatory)]
        [PSCredential]$AADCreds,

        [PoshBot.FromConfig('AADTenant')]
        [parameter(mandatory)]
        [string]$aadTenant,

        [PoshBot.FromConfig('FindAADUserAttributes')]
        [parameter(Mandatory)]
        [Array]$aadProperties #= @('givenName', 'surname', 'displayName', 'jobTitle', 'officeLocation', 'userPrincipalName', 'accountEnabled', 'employeeId', 'mail')
    )
    
    $aadID = [string]$AADCreds.UserName
    $aadSecret = [SecureString]$AADCreds.Password
    
    try {
        $myToken = Get-MsalToken -clientID $aadID -clientSecret $aadSecret -tenantID $aadTenant 
    }
    catch {
        Write-Verbose "Authentication to Azure AD failed. $($_)"   
        $o = $_ | Format-Table -AutoSize | Out-String -Width $Width 
    }

    $apiBeta = 'beta'
    $baseURI = 'https://graph.microsoft.com/'
    $usersURI = '/users'

    $resultUsers = $null
    
    $resultUsersURI = "$($baseURI)$($apiBeta)$($usersURI)?filter=startswith(displayName,`'$($Identity)`')"
    $resultUsers = Invoke-RestMethod -Headers @{Authorization = "Bearer $($myToken.AccessToken)" } `
        -Uri  $resultUsersURI `
        -Method Get

    if ($resultUsers.value.Count -ge 1) {
        $userObj = $resultUsers.value
        $o = $userObj | Select-Object -Property $aadProperties | Format-Table -AutoSize | Out-String -Width $Width
    }   
    else {
        $o = "Azure Active Directory User with displayName containing `'$($Identity)`' not found!"
    }
    New-PoshBotCardResponse -Type Normal -Text $o  
}
$CommandsToExport += 'Find-AADUser'

function Get-AADUser {
    <#
    .SYNOPSIS
        Get Azure AD User
    .PARAMETER Identity <UPN>
        Azure AD UPN
    .USAGE
        !GetAADUser <UPN>
    .EXAMPLE
        !GetAADUser 'Darren.Robinson@myaadtenant.com.au' 
        
        Return MFA Info for the user
        !GetAADUser 'Darren.Robinson@myaadtenant.com.au' -mfa 

        Return Group Memberships for the user
        !GetAADUser 'Darren.Robinson@myaadtenant.com.au' -groups

        Return Group and MFA info for the user
        !GetAADUser 'Darren.Robinson@myaadtenant.com.au' -groups -mfa 

    .DESCRIPTION !GetAADUser
        Get an Azure AD User using UPN
    .LINK
        https://blog.darrenjrobinson.com
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'GetAADUser',
        Aliases = ('AADUser'),
        Permissions = 'read'
    )]
    param(
        [parameter(position = 1,
            parametersetname = 'id')]
        [string]$Identity,

        [PoshBot.FromConfig('AADCreds')]
        [parameter(mandatory)]
        [PSCredential]$AADCreds,

        [PoshBot.FromConfig('AADTenant')]
        [parameter(mandatory)]
        [string]$aadTenant,

        [PoshBot.FromConfig('GetAADUserAttributes')]
        [parameter(Mandatory)]
        [Array]$aadProperties, #= @('givenName', 'surname', 'displayName', 'jobTitle', 'officeLocation', 'userPrincipalName', 'accountEnabled', 'employeeId', 'mail'),

        [switch]$mfa,

        [switch]$groups                
    )

    $aadID = [string]$AADCreds.UserName
    $aadSecret = [SecureString]$AADCreds.Password

    try {
        $myToken = Get-MsalToken -clientID $aadID -clientSecret $aadSecret -tenantID $aadTenant 
    }
    catch {
        Write-Verbose "Authentication to Azure AD failed. $($_)"   
        $o = $_ | Format-Table -AutoSize | Out-String -Width $Width 
    }

    $apiBeta = 'beta'
    $baseURI = 'https://graph.microsoft.com/'
    $mfaURI = '/reports/credentialUserRegistrationDetails'
    $usersURI = '/users'

    $resultUsers = $null
    $resultMFA = $null 
    $resultGroups = $null 
    
    $resultUsersURI = "$($baseURI)$($apiBeta)$($usersURI)?filter=userPrincipalName eq `'$($Identity)`'"
    $resultUsers = Invoke-RestMethod -Headers @{Authorization = "Bearer $($myToken.AccessToken)" } `
        -Uri  $resultUsersURI `
        -Method Get

    if ($resultUsers.value.Count -eq 1) {
        $userObj = $resultUsers.value
        
        if ($mfa) {
            $aadProperties = $aadProperties + @('ssprIsRegistered', 'ssprIsEnabled', 'mfaSsprIsCapable', 'mfaIsRegistered', 'mfaAuthMethods')
            $resultMFAURI = "$($baseURI)$($apiBeta)$($mfaURI)?filter=userPrincipalName eq `'$($Identity)`'"
            $resultMFA = Invoke-RestMethod -Headers @{Authorization = "Bearer $($myToken.AccessToken)" } `
                -Uri  $resultMFAURI `
                -Method Get
        
            if ($resultMFA.value) {
                $userObj | Add-Member -Type NoteProperty -Name "ssprIsRegistered" -Value $resultMFA.value.isRegistered
                $userObj | Add-Member -Type NoteProperty -Name "ssprIsEnabled" -Value $resultMFA.value.isEnabled
                $userObj | Add-Member -Type NoteProperty -Name "mfaSsprIsCapable" -Value $resultMFA.value.isCapable
                $userObj | Add-Member -Type NoteProperty -Name "mfaIsRegistered" -Value $resultMFA.value.isMfaRegistered
                $userObj | Add-Member -Type NoteProperty -Name "mfaAuthMethods" -Value $resultMFA.value.authMethods
            }
        }
        
        if ($groups) {
            $resultGroupURI = "$($baseURI)$($apiBeta)/users/$($Identity)/memberOf"
            $resultGroups = Invoke-RestMethod -Headers @{Authorization = "Bearer $($myToken.AccessToken)" } `
                -Uri  $resultGroupURI `
                -Method Get

            if ($resultGroups.value) {    
                $aadProperties = $aadProperties + @('groups')        
                $userObj | Add-Member -Type NoteProperty -Name "groups" -Value @($resultGroups.value.displayName -join ", ") 
            }
        }
        $o = $userObj | Select-Object -Property $aadProperties | Format-List | Out-String -Width $Width
    }
    else {
        $o = "`'$($Identity)`' not found. Try searching for the Azure AD user using !FindAADUser displayName"
    }
    
    New-PoshBotCardResponse -Type Normal -Text $o 
}
$CommandsToExport += 'Get-AADUser'

Export-ModuleMember -Function $CommandsToExport