# ----------------------------------------------------------------------------------------------------------
# PURPOSE:    WSUS - Module to manage Update Approvals for WSUS Groups
#
# VERSION     DATE         USER                DETAILS
# 1           03/05/2017   Craig Tolley        First Version
# 1.1         30/05/2017   Craig Tolley        Added 'Superseded' information to update details
# ----------------------------------------------------------------------------------------------------------
$Global:WsusServer = $null

function Connect-WsusServer {
[CmdletBinding()]
    Param (
        [ValidateNotNullOrEmpty()]
        [String]$WsusServerFqdn = "localhost",
        
        [Int]$WsusServerPort = 8530,
        
        [Boolean]$WsusServerSecureConnect = $false
    )

    # Load the assembly required
    try {
        [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
    }
    catch {
        Write-Error "Unable to load the Microsoft.UpdateServices.Administration assembly: $($_.Exception.Message). Is the WSUS RSAT installed on this system?"
        $Global:WsusServer = $null
        break
    }

    # Attempt the connection to the WSUS Server
    try {
        $Global:WsusServer = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WsusServerFqdn, $WsusServerSecureConnect, $WsusServerPort)
    }
    catch {
        $Global:WsusServer = $null
        Write-Error "Unable to connect to the WSUS Server: $($_.Exception.Message)"
    }
}

function Get-WsusUpdateClassifications {

if ($Global:WsusServer -eq $null) { Write-Error "WSUS Connection not initialized" }

$UpdateClassifications = $WsusServer.GetUpdateClassifications()
$UpdateClassifications

}

function Get-WsusUpdateDetails {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$UpdateGuid
    )

    if ($Global:WsusServer -eq $null) { Write-Error "WSUS Connection not initialized" }
        Write-Verbose "Update GUID: $UpdateGuid"
    try {
        $Update = $WsusServer.GetUpdate([Guid]$UpdateGuid) 
    } 
    catch { 
        Write-Error "Update could not be retrieved from the database. Error: $_"
        return
    }

    # Add in a collection of all of the approvals for the update so that we return an entire object containing all of the information about the update. 
    Add-Member -InputObject $Update -MemberType NoteProperty -Name "Approvals" -Value NotSet
    $Update.Approvals = @()
    $Groups = $WsusServer.GetComputerTargetGroups()
    ForEach ($Approval in $Update.GetUpdateApprovals()) {
        $UpdateApproval = New-Object -TypeName PSCustomObject
        Add-Member -InputObject $UpdateApproval -MemberType NoteProperty -Name Action -Value $Approval.Action
        Add-Member -InputObject $UpdateApproval -MemberType NoteProperty -Name CreationDate -Value $Approval.CreationDate
        Add-Member -InputObject $UpdateApproval -MemberType NoteProperty -Name GoLiveTime -Value $Approval.GoLiveTime
        Add-Member -InputObject $UpdateApproval -MemberType NoteProperty -Name Deadline -Value $Approval.Deadline
        Add-Member -InputObject $UpdateApproval -MemberType NoteProperty -Name State -Value $Approval.State
        Add-Member -InputObject $UpdateApproval -MemberType NoteProperty -Name AdministratorName -Value $Approval.AdministratorName


        Add-Member -InputObject $UpdateApproval -MemberType NoteProperty -Name ComputerGroupName -Value NotSet
        $UpdateApproval.ComputerGroupName = $Groups | Where { $_.Id -eq $Approval.ComputerTargetGroupId } | Select -ExpandProperty Name

        $Update.Approvals += $UpdateApproval
    }

    return $Update
}

function Get-WsusUpdateCategories {

if ($Global:WsusServer -eq $null) { Write-Error "WSUS Connection not initialized" }

$UpdateCategories  = $WsusServer.GetUpdateCategories()
$UpdateCategories 

}

function Get-WsusComputerGroups {

if ($Global:WsusServer -eq $null) { Write-Error "WSUS Connection not initialized" }

$ComputerGroups  = $WsusServer.GetComputerTargetGroups()
$ComputerGroups 

}

function Get-WsusComputerGroupParents {
    Param (
        [Parameter(Mandatory=$true)]
        $ComputerGroupGuid,

        # If specified, all parent groups will be returned. If not, only the direct parent of the specified group will be returned
        [switch]$Recurse,

        # If specified, the group specified in the $ComputerGroupGuid parameter will be included in the results to give a full tree
        [switch]$IncludeBaseGroupInResults
    )

    $Groups = $WsusServer.GetComputerTargetGroups()
    $ParentGroups = @()
    
    if ($IncludeBaseGroupInResults) { $ParentGroups += $Groups | Where { $_.Id -eq $ComputerGroupGuid } }
    
    ForEach( $Group in $Groups) {
        $Children = $Group.GetChildTargetGroups()
        if ($Children.Id -contains $ComputerGroupGuid) { 
            $ParentGroups += $Group 
            if ($Recurse) { $ParentGroups += Get-WsusComputerGroupParents -ComputerGroupGuid $Group.Id.ToString() }
        }
    }
    return $ParentGroups    
}

function Get-WsusUpdatesAndApprovals {
    Param (
        [string]$ClassificationGuid = "*",
        
        [string]$UpdateCategoryGuid = "*",
        
        [string]$ComputerGroupGuid,

        [ValidateSet("Any", "Approved", "Unapproved", "Declined")]
        [string]$ApprovalStatus = "Any"
    )

    if ($Global:WsusServer -eq $null) { Write-Error "WSUS Connection not initialized" }

    # Get the details of the specified Group
    $Groups = $WsusServer.GetComputerTargetGroups()
    If ($Groups.Id -notcontains $ComputerGroupGuid) {
        Write-Error "Group cannot be found in the list of groups on the WSUS Server"
        return
    }
    $PossibleGroups = @()
    $PossibleGroups += $Groups | Where {$_.Id -eq $ComputerGroupGuid}
    $PossibleGroups += Get-WsusComputerGroupParents -ComputerGroupGuid $ComputerGroupGuid -Recurse

    # Build an Update Scope with the filters that have been defined and retrieve the updates
    $UpdateSearchScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    Write-Verbose "Classification GUID: $ClassificationGuid"
    Write-Verbose "Update Category GUID: $UpdateCategoryGuid"
    If ($ClassificationGuid -ne "*") {
        $UpdateSearchScope.Classifications.AddRange(($WsusServer.GetUpdateClassifications() | Where { $_.Id -eq $ClassificationGuid}))
    }
    If ($UpdateCategoryGuid -ne "*") {
        $UpdateSearchScope.Categories.AddRange(($WsusServer.GetUpdateCategories() | Where { $_.Id -eq $UpdateCategoryGuid}))
    }
    $Updates = $WsusServer.GetUpdates($UpdateSearchScope)
    
    # Go through each of the updates. If the update is approved/unapproved for the group then add to the collection
    $i = 0
    $UpdateDetails = @()
    ForEach ($Update in $Updates)
    {
        $i ++
        Write-Progress -Activity "Getting update approvals" -PercentComplete (($i/$($Updates.Count))*100) -Status "$i of $($Updates.Count)"

        $UpdateObj = New-Object -TypeName PSCustomObject
        Add-Member -InputObject $UpdateObj -MemberType NoteProperty -Name Id -Value $Update.Id.UpdateId.Guid
        Add-Member -InputObject $UpdateObj -MemberType NoteProperty -Name Title -Value $Update.Title
        Add-Member -InputObject $UpdateObj -MemberType NoteProperty -Name ReleaseDate -Value $Update.ArrivalDate
        Add-Member -InputObject $UpdateObj -MemberType NoteProperty -Name UpdateClassification -Value $Update.UpdateClassificationTitle
        Add-Member -InputObject $UpdateObj -MemberType NoteProperty -Name ApprovalState -Value NotSet
        Add-Member -InputObject $UpdateObj -MemberType NoteProperty -Name Checked -Value $false
        
        $UpdateApprovals = $Update.GetUpdateApprovals() | Where { $PossibleGroups.Id -contains $_.ComputerTargetGroupId } 
        
        # This method was dropped as it takes around 30% longer in testing.
        #$UpdateApprovals = $PossibleGroups | ForEach { $Update.GetUpdateApprovals($_) } 

        if ($Update.IsDeclined -eq $true -and $Update.IsSuperseded -eq $true) { 
            $UpdateObj.ApprovalState = "Declined (Superseded)" 
        } 

        elseif ($Update.IsDeclined -eq $true -and $Update.IsSuperseded -eq $false) { 
            $UpdateObj.ApprovalState = "Declined" 
        } 

        elseif ($UpdateApprovals.Count -eq 0) {
            $UpdateObj.ApprovalState = "Unapproved"
        }
        
        elseif (($UpdateApprovals | Measure).Count -eq 1) {
            $Group = $PossibleGroups | Where {$_.Id -eq $UpdateApprovals.ComputerTargetGroupId}
            $UpdateObj.ApprovalState = "Approved - $($UpdateApprovals.Action) ($($Group.Name))"
        }

        # We have multiple approvals, so we need to work up the tree to find the one that actually applies.
        else {
            $ChildGroupId = $ComputerGroupGuid
            Do {
                $ParentGroup = Get-WsusComputerGroupParents -ComputerGroupGuid $ChildGroupId
                $ParentApprovals = $UpdateApprovals | Where {$_.ComputerTargetGroupId -eq $ParentGroup.Id}
                
                # If an approval is found, then we return that record. Set ChildGroupId to Null to break out
                if ($ParentApprovals -ne $null) {
                    $UpdateObj.ApprovalState = "Approved - $($ParentApprovals.Action) ($($ParentGroup.Name))"
                    $ChildGroupId = $null
                }

                # If no approvals are found for the parent, then go up another level
                else {
                    $ChildGroupId = $ParentGroup.Id
                }
            }
            Until ($ChildGroupId -eq $null)
        }
        
        $UpdateDetails += $UpdateObj
    }
    Write-Progress -Completed $true
    # Output Results based on the ApprovalStatus Flag
    If ($ApprovalStatus -eq "Any") { Write-Output $UpdateDetails }
    Else { Write-Output $UpdateDetails | Where {$_.ApprovalState -eq $ApprovalStatus } }
    
}

function Set-WsusSetUpdateApproval {
    Param (
        [ValidateNotNullOrEmpty()]
        [string]$UpdateGuid,

        [ValidateNotNullOrEmpty()]
        [string]$ComputerGroupGuid,

        [ValidateSet("Install", "Uninstall", "NotApproved")]
        [string]$ApprovalAction
    )

    if ($Global:WsusServer -eq $null) { Write-Error "WSUS Connection not initialized" }

    # Get the update. If this fails then the update does not exist. 
    Write-Verbose "Update GUID: $UpdateGuid"
    try {
        $Update = $WsusServer.GetUpdate([Guid]$UpdateGuid) 
    } 
    catch { 
        Write-Error "Unable to change approvals for update. Update could not be retrieved from the database. Error: $_"
        return
    }

    # Get the details of the specified Group
    Write-Verbose "Computer Group Guid: $ComputerGroupGuid"
    $Groups = $WsusServer.GetComputerTargetGroups()
    If ($Groups.Id -notcontains $ComputerGroupGuid) {
        Write-Error "Group cannot be found in the list of groups on the WSUS Server"
        return
    }
    $GroupObj = $Groups | Where {$_.Id -eq $ComputerGroupGuid}

    Write-Verbose "Approval Action: $ApprovalAction"
    try {
        $Update.Approve($ApprovalAction, $GroupObj)
        Write-Verbose "Approved: $($Update.Title). Group: $($GroupObj.Name)"
    }
    catch {
        Write-Error "Unable to approve update $($Update.Title) for action $ApprovalAction and Target Group $($GroupObj.Name). Error $_"
        return
    }
}

function Decline-WsusUpdate {
    Param (
        [ValidateNotNullOrEmpty()]
        [string]$UpdateGuid
    )

    if ($Global:WsusServer -eq $null) { Write-Error "WSUS Connection not initialized" }

    # Get the update. If this fails then the update does not exist. 
    Write-Verbose "Update GUID: $UpdateGuid"
    try {
        $Update = $WsusServer.GetUpdate([Guid]$UpdateGuid) 
    } 
    catch { 
        Write-Error "Unable to decline update. Update could not be retrieved from the database. Error: $_"
        return
    }

    try {
        $Update.Decline()
        Write-Verbose "Declined: $($Update.Title). Group: $($GroupObj.Name)"
    }
    catch {
        Write-Error "Unable to Decline update $($Update.Name). Error: $_"
        return
    }
}