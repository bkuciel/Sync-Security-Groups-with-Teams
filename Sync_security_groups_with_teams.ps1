[CmdletBinding(SupportsShouldProcess=$true)]
Param()
# Logs Location
$Logfile = "$PSScriptRoot\Logs\Sync-$(get-date -f dd-MM-yyyy).log"

# Logs folder
if(!(Test-Path $PSScriptRoot\Logs -PathType Container))
{
    New-Item -ItemType Directory -Force -Path $PSScriptRoot\Logs
}
#region Functions
Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
}
Function Get-RecursiveAzureAdGroupMemberUsers{
[cmdletbinding()]
param(
   [parameter(Mandatory=$True,ValueFromPipeline=$true)]
   $AzureGroup
)
    Begin{
        If(-not(Get-AzureADCurrentSessionInfo)){Connect-AzureAD $credentials}
    }
    Process {
        [psobject[]] $UserMembers = @()
        Write-Verbose -Message "Enumerating $($AzureGroup.DisplayName)"
        $Members = Get-AzureADGroupMember -ObjectId $AzureGroup.ObjectId -All $true
        
        $UserMembers = $Members | Where-Object{$_.ObjectType -eq 'User'}
        If($Members | Where-Object{$_.ObjectType -eq 'Group'}){
            $UserMembers += $Members | Where-Object{$_.ObjectType -eq 'Group'} | ForEach-Object{ Get-RecursiveAzureAdGroupMemberUsers -AzureGroup $_}
        }
        
    }
    end {
        Return $UserMembers
    }
}
#endregion

$Time = get-date -f "dd-MM-yyyy HH:mm:ss"
LogWrite "Starting script $Time"

#region Authentication
$credentials = import-clixml -path "$PSScriptRoot\cred.clixml"
try {
    Connect-MicrosoftTeams -Credential $credentials
    Connect-AzureAD -Credential $credentials
}
catch {
    Write-Host "Problem with credentials occured:"
    Write-Host $_
    LogWrite "Problem with credentials occured:"
    LogWrite "$($_)"
    exit
}
#endregion

#region config_file
try{
    $list = import-csv "$PSScriptRoot\Security_groups_Teams_sync_example.csv"
}
catch{
    Write-host "CSV file not found"
    Write-Host $_
    LogWrite "$($_)"
    exit
}
#endregion

$AllADSecGroups = Get-AzureADGroup -All $true
#Add users from Security Group to Teams and Teams Channel
Write-Host "Looking up for users that should be added to Teams"
LogWrite "Looking up for users that should be added to Teams"
Foreach($item in $list){
    $ADSecurityGroupName = $item.AD_Security_Group
    $TeamName = $item.Team
    $ChannelName = $item.Channel
    $Team = Get-Team -DisplayName $teamname
    try{
        $Channel = (Get-TeamChannel -GroupId $($Team.GroupId)) | Where-Object {$_.DisplayName -eq $ChannelName}
    }
    catch{
        write-information $_
    }
    Write-Verbose "Processing Team $TeamName, $ChannelName, AD: $ADSecurityGroupName"

    #Get Members of Security group
    $ADSecGroup = $AllADSecGroups | Where-Object {$_.DisplayName -eq $ADSecurityGroupName}
    $SecurityGroupUsers = $ADSecGroup | Get-RecursiveAzureAdGroupMemberUsers | select -Unique

    #Adding user to TeamChannel if not present
    Foreach($SecurityGroupUser in $SecurityGroupUsers)
    {
        Write-Verbose "Processing user $($SecurityGroupUser.UserPrincipalName)"
        if($ChannelName -ne "General"){
            try{
                $ChannelUser = Get-TeamChannelUser -GroupId $Team.GroupID -DisplayName $ChannelName | Where-Object {$_.User -eq $SecurityGroupUser.UserPrincipalName}
            }
            catch{
                write-information $_
            }
            if(!$ChannelUser){   
                # Check if user is a part of Team
                $TeamUser = Get-TeamUser -GroupId $Team.GroupID | Where-Object {$_.User -eq $SecurityGroupUser.UserPrincipalName}
                if(!$TeamUser){
                    Write-Host Adding User $SecurityGroupUser.UserPrincipalName to Team $Team.DisplayName
                    LogWrite "Adding User $($SecurityGroupUser.UserPrincipalName) to Team $($Team.DisplayName)"
                    try{
                        Add-TeamUser -GroupId $Team.GroupID -User $SecurityGroupUser.UserPrincipalName
                    }
                    catch{
                        Write-Information $_
                    }
                }

                Write-Host Creating user: $SecurityGroupUser.UserPrincipalName
                Write-Host in Team: $Team.DisplayName
                Write-Host in Channel: $ChannelName
                LogWrite "Creating user: $($SecurityGroupUser.UserPrincipalName)"
                LogWrite "in Team: $($Team.DisplayName)"
                LogWrite "in Channel: $ChannelName"
                # Add User from Security Group to Team Channel
                if($ChannelName -ne "General"){
                    if($Channel.MembershipType -eq "Private"){
                        try{
                            Add-TeamChannelUser -GroupId $Team.GroupID -DisplayName $ChannelName -User $SecurityGroupUser.UserPrincipalName
                        }
                        catch {
                            Write-Information $_
                        }
                    }
                }
            }
        }
        else{
            $ChannelUser = Get-TeamUser -GroupId $Team.GroupID | Where-Object {$_.User -eq $SecurityGroupUser.UserPrincipalName}
            if(!$ChannelUser){
                Write-Host Adding User $SecurityGroupUser.UserPrincipalName to Team $Team.DisplayName
                LogWrite "Adding User $($SecurityGroupUser.UserPrincipalName) to Team $($Team.DisplayName) and to Channel $ChannelName"
                try{
                    Add-TeamUser -GroupId $Team.GroupID -User $SecurityGroupUser.UserPrincipalName 
                }
                catch{
                    Write-Information $_
                }
            }
        }
    }
Write-Host ""
}
#Remove users from Teams Channel not listed in Security Group
Write-host "Looking up for users that should be removed from Teams"
LogWrite "Looking up for users that should be removed from Teams"
ForEach($item in $list)
{
    $ADSecurityGroupName = $item.AD_Security_Group
    $TeamName = $item.Team
    $ChannelName = $item.Channel
    $Team = Get-Team -DisplayName $teamname
    $Channel = (Get-TeamChannel -GroupId $($Team.GroupId)) | Where-Object {$_.DisplayName -eq $ChannelName}
    Write-Verbose "Processing Team $TeamName, $ChannelName, AD: $ADSecurityGroupName"

    $ChannelUsers = Get-TeamChannelUser -GroupId $($Team.GroupID) -DisplayName $ChannelName #Users in Teams Channel
    $ADSecGroup = $AllADSecGroups | Where-Object {$_.DisplayName -eq $ADSecurityGroupName}

    ForEach($ChannelUser in $ChannelUsers)
    {
        Write-Verbose "Processing ChannelUser: $($ChannelUser.User)"
        if($ChannelUser.User -ne "admin@contoso.onmicrosoft.com"){
            $SecurityGroupUser = $ADSecGroup | Get-RecursiveAzureAdGroupMemberUsers | Where-Object {$_.UserPrincipalName -eq $ChannelUser.User}
            if(!$SecurityGroupUser){
                Write-Host Removing user: $ChannelUser.User
                Write-Host From $Team.DisplayName, Channel $ChannelName
                LogWrite "Removing user: $($ChannelUser.User)"
                LogWrite "From $($Team.DisplayName), Channel $ChannelName"
                if(($ChannelName -ne "General") -and ($Channel.MembershipType -eq "Private")){
                    Remove-TeamChannelUser -GroupId $Team.GroupID -DisplayName $ChannelName -User $ChannelUser.User
                }
                elseif($ChannelName -eq "General"){
                    Remove-TeamUser -GroupId $Team.GroupID -User $ChannelUser.User
                }
            }
        }
    }
Write-Host ""
}
Disconnect-AzureAD
Disconnect-MicrosoftTeams
$Time = get-date -f "dd-MM-yyyy HH:mm:ss"
LogWrite "Script Ended at $Time"