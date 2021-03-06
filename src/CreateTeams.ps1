[CmdletBinding()]
param(
    $guestAccess,
    $teamsName,
    $teamsOwner,
    $SPListItemID
)

#Automatic Teams creation starts here
#Global variables
$SPSite = "https://codesalot.sharepoint.com/sites/TeamsDemo/"
$SPList = "Nytt Team"

$login = Get-AutomationPSCredential -Name 'AzureAdmin'

Login-AzureRmAccount -Credential $login
$tenantId = (Get-AzureRmAutomationVariable -Name 'TenantId' -AutomationAccountName 'TeamsDemo' -ResourceGroupName "TeamsDemo").Value

Import-Module SharePointPnPPowerShellOnline
Import-Module MicrosoftTeams

Function Invoke-FileUpload{
    Param(   
    #[Parameter(Mandatory=$true,ValueFromPipeline=$true)]        
    #[String]$UserName,
            
    #[Parameter(Mandatory=$true,ValueFromPipeline=$true)]        
    #[String]$GroupName
    )

    $spoconn = Connect-PnPOnline –Url "https://codesalot.sharepoint.com/sites/$teamsName" –Credentials $login -ReturnConnection
    Add-PnPFolder -Name "General" -Folder "/Shared Documents"
    Add-PnPFolder -Name "01 Planning" -Folder "/Shared Documents"
    Add-PnPFolder -Name "02 Execution" -Folder "/Shared Documents"
    Add-PnPFolder -Name "03 Final" -Folder "/Shared Documents"

    #copy template to new team channel
    #Copy-PnPFile -SourceUrl /sites/JHKontraktHndtering/Shared%20Documents/Avtale%20Mal.docx -TargetUrl /sites/$TeamsName/Shared%20Documents/General -Force -Confirm
         
    
} #End Upload-Files 


#Connecting to O365
Connect-MicrosoftTeams -TenantId $tenantId -Credential $login

#Create new Team
$team = New-Team -displayname $teamsName -Visibility Private
Add-TeamUser -GroupId $team.GroupId -User $TeamsOwner.Email -Role Owner

#Add channels
New-TeamChannel -GroupId $team.GroupId -DisplayName "01 Planning"
New-TeamChannel -GroupId $team.GroupId -DisplayName "02 Execution"
New-TeamChannel -GroupId $team.GroupId -DisplayName "03 Final"

#Teams created
Write-Output 'Teams created'

#call upload file function
Invoke-FileUpload

#Disabling Guest Access to Teams
Write-Output "GuestAccess allowed:"
Write-Output $guestAccess


if($guestAccess -eq "No")
{
        #importing AzureADPreview modules
        Import-Module AzureADPreview
        Connect-AzureAD -TenantId $tenantId -Credential $login

        #Turn OFF guest access
        $template = Get-AzureADDirectorySettingTemplate | ? {$_.displayname -eq "group.unified.guest"}
        $settingsCopy = $template.CreateDirectorySetting()
        $settingsCopy["AllowToAddGuests"]=$False

        New-AzureADObjectSetting -TargetType Groups -TargetObjectId $team.GroupId -DirectorySetting $settingsCopy

        #Verify settings
        Get-AzureADObjectSetting -TargetObjectId $team.GroupId -TargetType Groups | fl Values
        
        #reset $guestaccess flag
        $guestAccess = "NA"

}

#Updating SharePoint list item status
$spoconn = Connect-PnPOnline –Url $SPSite –Credentials (Get-AutomationPSCredential -Name 'AzureAdmin') -ReturnConnection -Verbose
$itemupdate = Set-PnPListItem -List $SPList -Identity $SPListItemID -Values @{"TeamsCreated" = "Success"} -Connection $spoconn

Write-Output "All done"
