$VerbosePreference = "continue"

[xml]$XAMLMain = Get-Content -Path (Join-Path -Path (Split-Path $MyInvocation.MyCommand.Path) -ChildPath "Wsus-UpdateManager.UI.Main.xaml")
[xml]$XAMLDetails = Get-Content -Path (Join-Path -Path (Split-Path $MyInvocation.MyCommand.Path) -ChildPath "Wsus-UpdateManager.UI.Details.xaml")
[xml]$XAMLHelpAbout = Get-Content -Path (Join-Path -Path (Split-Path $MyInvocation.MyCommand.Path) -ChildPath "Wsus-UpdateManager.UI.HelpAbout.xaml")

function Show-WsusUpdateManagerUi {
    # Load the WPF Assemblys
    Add-type -AssemblyName PresentationCore
    Add-type -AssemblyName PresentationFramework
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    # XAML Window code
    $Reader= New-Object System.Xml.XmlNodeReader $global:XAMLMain
    $Window=[Windows.Markup.XamlReader]::Load($Reader)  

    #region BindControls
    $txt_CurrentAction = $Window.FindName("CurrentAction")
    $txt_ItemCounts = $Window.FindName("ItemCounts")

    $grp_connect = $Window.FindName("group_connect")
    $txt_connect_servername = $Window.FindName("connect_servername")
    $txt_connect_serverport = $Window.FindName("connect_serverport")
    $chk_connect_usesecureconnect = $Window.FindName("connect_usesecureconnect")
    $btn_connect_go = $Window.FindName("connect_go")

    $grp_filters = $Window.FindName("group_filters")
    $ddl_filters_computergroup = $Window.FindName("filters_computergroup")
    $ddl_filters_updatecategory = $Window.FindName("filters_updatecategory")
    $ddl_filters_updateclassification = $Window.FindName("filters_updateclassification")
    $ddl_filters_approvalstatus = $Window.FindName("filters_approvalstatus")
    $chk_filters_includesuperseded = $Window.FindName("filters_includesuperseded")
    $btn_filters_getupdatedetails = $Window.FindName("filters_getupdatedetails")

    $grp_updates = $Window.FindName("group_updates")
    $btn_manage_selectall = $Window.FindName("manage_selectall")
    $btn_manage_unselectall = $Window.FindName("manage_unselectall")
    $btn_manage_approveselected = $Window.FindName("manage_approveselected")
    $btn_manage_declineselected = $Window.FindName("manage_declineselected")
    $btn_manage_exportcsv = $Window.FindName("manage_exportcsv")
    $dgv_manage_updates = $Window.FindName("manage_updates")
    $ddl_manage_approvalsaction = $Window.FindName("ddl_manage_approvalsaction")
    $ddl_manage_approvalstarget = $Window.FindName("ddl_manage_approvalstarget")

    $btn_show_help = $Window.FindName("show_help")

    #endregion

    #region ConnectEvents
        $btn_connect_go.Add_Click( {
            try {
                $txt_CurrentAction.Content = "Connecting..."
                $grp_filters.IsEnabled = $false
                $grp_updates.IsEnabled = $false
                Connect-WsusServer -WsusServerFqdn $txt_connect_servername.Text -WsusServerPort $txt_connect_serverport.Text -WsusServerSecureConnect ([bool]($chk_connect_usesecureconnect.Checked)) -ErrorAction Stop
                $grp_filters.IsEnabled = $true
                $txt_CurrentAction.Content = "Connected"
            }
            catch {
                [System.Windows.MessageBox]::Show("Could not connect to the WSUS Server. `r`n`r`nError: $($_)", "Unable to Connect",[System.Windows.MessageBoxButton]::OK ,[System.Windows.MessageBoxImage]::Error)
                $grp_filters.IsEnabled = $false
                $grp_updates.IsEnabled = $false
                $txt_CurrentAction.Content = "Failed to connect."
                return
            }

            # Now we are connected, retrieve details to populate all of the filters
            try {
                # We build the collection this way rather than directly defining in the XAML code as it make it easier when evaluating the currently selected value later on in the code. 
                # The actual options don't change though
                $txt_CurrentAction.Content = "Retrieving Approval Status'"
                $ApprovalStatus = @()
                $ApprovalStatus += "" | Select @{N="Text";E={"(Any)"}}, @{N="Value";E={"Any"}}
                $ApprovalStatus += "" | Select @{N="Text";E={"Approved"}}, @{N="Value";E={"Approved"}}
                $ApprovalStatus += "" | Select @{N="Text";E={"Declined"}}, @{N="Value";E={"Declined"}}
                $ApprovalStatus += "" | Select @{N="Text";E={"Unapproved"}}, @{N="Value";E={"Unapproved"}}
                $ddl_filters_approvalstatus.ItemsSource = $ApprovalStatus
                $ddl_filters_approvalstatus.DisplayMemberPath = "Text"
                $ddl_filters_approvalstatus.SelectedValuePath = "Value"
                $ddl_filters_approvalstatus.SelectedIndex = 0

                $txt_CurrentAction.Content = "Retrieving Update Categories"
                $UpdateCategories = @()
                $UpdateCategories += "" | Select @{N="Id";E={"*"}}, @{N="Title";E={"(Any)"}}
                $UpdateCategories += Get-WsusUpdateCategories | Select Id, Title | Sort Title
                $ddl_filters_updatecategory.ItemsSource = $UpdateCategories
                $ddl_filters_updatecategory.DisplayMemberPath = "Title"
                $ddl_filters_updatecategory.SelectedValuePath = "Id"
                $ddl_filters_updatecategory.SelectedIndex = 0

                $txt_CurrentAction.Content = "Retrieving Update Classifications"
                $UpdateClassifications = @()
                $UpdateClassifications += "" | Select @{N="Id";E={"*"}}, @{N="Title";E={"(Any)"}}
                $UpdateClassifications += Get-WsusUpdateClassifications
                $ddl_filters_updateclassification.ItemsSource = $UpdateClassifications | Select Id, Title | Sort Title
                $ddl_filters_updateclassification.DisplayMemberPath = "Title"
                $ddl_filters_updateclassification.SelectedValuePath = "Id"
                $ddl_filters_updateclassification.SelectedIndex = 0

                $txt_CurrentAction.Content = "Retrieving Computer Groups"
                $ComputerGroups = @(Get-WsusComputerGroups | Select Id, Name | Sort Name)
                $ddl_filters_computergroup.ItemsSource = $ComputerGroups
                $ddl_filters_computergroup.DisplayMemberPath = "Name"
                $ddl_filters_computergroup.SelectedValuePath = "Id"
                $ddl_filters_computergroup.SelectedIndex = 0

                $txt_CurrentAction.Content = "Ready to search for Updates"
                $txt_ItemCounts.Content = ""

            }
            catch {
                [System.Windows.MessageBox]::Show("Could not retrieve details of Categories, Classifications and Computer Groups. `r`n`r`nError: $($_)", "Failed to Get Current Configuration",[System.Windows.MessageBoxButton]::OK ,[System.Windows.MessageBoxImage]::Error)
                $grp_filters.IsEnabled = $false
                $grp_updates.IsEnabled = $false
                $txt_CurrentAction.Content = "Failed to load all options."
                return
            }
        } )
    #endregion

    #region FilterEvents
        $btn_filters_getupdatedetails.Add_Click( {
            # Retrieve details of all of the updates that meet the supplied criteria and display in the Grid View
            try {
                $txt_CurrentAction.Content = "Retrieving details of updates (this can take some time depending on your chosen filters)"
                $dgv_manage_updates.ItemsSource = $null
                $Global:Updates = Get-WsusUpdatesAndApprovals -ClassificationGuid $ddl_filters_updateclassification.SelectedValue -UpdateCategoryGuid $ddl_filters_updatecategory.SelectedValue -ComputerGroupGuid $ddl_filters_computergroup.SelectedItem.Id -ApprovalStatus $ddl_filters_approvalstatus.SelectedValue -IncludeSupersededUpdates ([bool]($chk_filters_includesuperseded.IsChecked))
                $dgv_manage_updates.ItemsSource = $Global:Updates
                $txt_ItemCounts.Content = "Retrieved $($Global:Updates.Count) updates."
                
                # Get details of all of the parent groups
                $txt_CurrentAction.Content = "Retrieving details of parent groups"
                $ComputerGroups = @(Get-WsusComputerGroupParents -ComputerGroupGuid $ddl_filters_computergroup.SelectedItem.Id -Recurse -IncludeBaseGroupInResults | Select Id, Name)
                $ddl_manage_approvalstarget.ItemsSource = $ComputerGroups
                $ddl_manage_approvalstarget.DisplayMemberPath = "Name"
                $ddl_manage_approvalstarget.SelectedValuePath = "Id"
                $ddl_manage_approvalstarget.SelectedIndex = 0

                $txt_CurrentAction.Content = "Update details retrieved"
                $grp_updates.IsEnabled = $true
            }
            catch {
                [System.Windows.MessageBox]::Show("Unable to retrieve details of updates for the chosen criteria. `r`n`r`nError: $($_)", "Error Retrieving Updates",[System.Windows.MessageBoxButton]::OK ,[System.Windows.MessageBoxImage]::Error)
                $txt_CurrentAction.Content = "Failed to retrieve details of updates"
                $grp_updates.IsEnabled = $false
                return
            }
        } )
    #endregion

    #region ManageEvents
        $dgv_manage_updates.Add_MouseDoubleClick( {
            Show-WsusUpdateDetailsUi -UpdateGuid $dgv_manage_updates.SelectedItem.Id
        } )

        $btn_manage_selectall.Add_Click( {
            ForEach ($Update in $Global:Updates) { $Update.Checked = $true }
            $dgv_manage_updates.ItemsSource = $null
            $dgv_manage_updates.ItemsSource = $Global:Updates
        } )
        
        $btn_manage_unselectall.Add_Click( {
            ForEach ($Update in $Global:Updates) { $Update.Checked = $false }
            $dgv_manage_updates.ItemsSource = $null
            $dgv_manage_updates.ItemsSource = $Global:Updates
        } )
        
        $btn_manage_approveselected.Add_Click( {
            
            $CheckedUpdates = @($Global:Updates | Where {$_.Checked -eq $true})
            if (@($CheckedUpdates).Count -eq 0) { 
                $txt_CurrentAction.Content = "No updates selected"
                [System.Windows.MessageBox]::Show("Please check the updates that you would like to approve.","No Updates Selected", [System.Windows.MessageBoxButton]::OK ,[System.Windows.MessageBoxImage]::Warning)
                return
            }

            # Confirm with the user that they want to take the action that they have chosen
            $PromptTitle = "Approve Updates"
            $PromptMessage = "Are you sure you want to approve the $($CheckedUpdates.Count) selected updates for $($ddl_manage_approvalsaction.Text) to the $($ddl_manage_approvalstarget.SelectedItem.Name) group?"
            $Result = [System.Windows.MessageBox]::Show($PromptMessage,$PromptTitle,[System.Windows.MessageBoxButton]::YesNo ,[System.Windows.MessageBoxImage]::Question)
            if ($Result -ne [System.Windows.MessageBoxResult]::Yes) { return }
       
            $ApprovalErrors = @()
            ForEach ($Update in $Global:Updates | Where {$_.Checked -eq $true} ) {
                try {
                    $txt_CurrentAction.Content = "Setting Approvals for $($Update.Id)"
                    Set-WsusSetUpdateApproval -UpdateGuid $Update.Id -ComputerGroupGuid $ddl_manage_approvalstarget.SelectedItem.Id -ApprovalAction $ddl_manage_approvalsaction.Text
                    $Update.ApprovalState = "Approved - $($ddl_manage_approvalsaction.Text) ($($ddl_manage_approvalstarget.SelectedItem.Name))"
                } catch {
                    $ApprovalErrors += "Failed to set approval for '$($Update.Title). Error: $_"
                }
            }

            $dgv_manage_updates.ItemsSource = $null
            $dgv_manage_updates.ItemsSource = $Global:Updates

            # Notify the user of errors
            if ($ApprovalErrors.Count -gt 0) { [System.Windows.MessageBox]::Show([String]::Join("`r`n",$ApprovalErrors),'Error Approving Updates','Ok','Error') }
            
            $txt_CurrentAction.Content = "Completed setting approvals on $($CheckedUpdates.Count) with $($ApprovalErrors.Count) failures"
        } )
        
        $btn_manage_declineselected.Add_Click( {
            
            $CheckedUpdates = @($Global:Updates | Where { $_.Checked -eq $true })
            if (@($CheckedUpdates).Count -eq 0) { 
                $txt_CurrentAction.Content = "No updates selected" 
                [System.Windows.MessageBox]::Show("Please check the updates that you would like to decline.","No Updates Selected", [System.Windows.MessageBoxButton]::OK ,[System.Windows.MessageBoxImage]::Warning)
                return
            }

            # Confirm with the user that they want to take the action that they have chosen
            $PromptTitle = "Decline Updates"
            $PromptMessage = "Are you sure you want to decline the $($CheckedUpdates.Count) selected updates? Declined updates are not available to *ANY* target computer groups."
            $Result = [System.Windows.MessageBox]::Show($PromptMessage,$PromptTitle,[System.Windows.MessageBoxButton]::YesNo ,[System.Windows.MessageBoxImage]::Question)
            if ($Result -ne [System.Windows.MessageBoxResult]::Yes) { return }

            $DeclineErrors = @()
            ForEach ($Update in $CheckedUpdates) {
                try {
                    $txt_CurrentAction.Content = "Declining Update $($Update.Title)"
                    Decline-WsusUpdate -UpdateGuid $Update.Id
                    $Update.ApprovalState = "Declined"
                } catch {
                    $DeclineErrors += "Failed to set approval for '$($Update.Title). Error: $_"
                }
            }

            $dgv_manage_updates.ItemsSource = $null
            $dgv_manage_updates.ItemsSource = $Global:Updates

            # Notify the user of errors
            if ($DeclineErrors.Count -gt 0) { [System.Windows.MessageBox]::Show([String]::Join("`r`n",$DeclineErrors),'Error Declining Updates','Ok','Error') }
            $txt_CurrentAction.Content = "Completed declining $($CheckedUpdates.Count) updates with $($ApprovalErrors.Count) failures"
        } )
                
        $btn_manage_exportcsv.Add_Click( {
            if ($dgv_manage_updates.Items -eq $null) {
                [System.Windows.MessageBox]::Show("There is no data to export.", "No Data",[System.Windows.MessageBoxButton]::OK ,[System.Windows.MessageBoxImage]::Exclamation)
                return
            }

            try {
                $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                $SaveFileDialog.OverwritePrompt = $true
                $SaveFileDialog.Title = "Save Current View to CSV"
                $SaveFileDialog.Filename = "UpdatesExport.csv"
                $SaveFileDialog.Filter = "CSV (*.csv) | *.csv"
                $SaveDialogResult = $SaveFileDialog.ShowDialog()
                if ($SaveDialogResult -eq "OK") {
                    $dgv_manage_updates.Items | Export-Csv -Path $SaveFileDialog.FileName -NoTypeInformation
                }
                $txt_CurrentAction.Content = "CSV Export Completed to $($SaveFileDialog.FileName)"
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to export current updates details to CSV. `r`n`r`nError: $($_)", "Error Exporting Update Details",[System.Windows.MessageBoxButton]::OK ,[System.Windows.MessageBoxImage]::Error)
                $txt_CurrentAction.Content = "CSV Export failed"
                return
            }
        } )

    #endregion

    $btn_show_help.Add_Click( {
        Show-WsusUpdateHelpAboutUi
        } )

    $Window.ShowDialog() | Out-Null
} 

function Show-WsusUpdateDetailsUi {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$UpdateGuid
    )

    # Load the WPF Assemblys
    Add-type -AssemblyName PresentationCore
    Add-type -AssemblyName PresentationFramework
    
    # XAML Window code
    $Reader=(New-Object System.Xml.XmlNodeReader $global:XAMLDetails)  
    $DetailsWindow=[Windows.Markup.XamlReader]::Load( $Reader )  
   
    #region BindControls
        $txt_UpdateTitle = $DetailsWindow.FindName("txt_UpdateTitle")
        $txt_UpdateDescription = $DetailsWindow.FindName("txt_UpdateDescription")
        $txt_UpdateUrl = $DetailsWindow.FindName("txt_UpdateUrl")
        $txt_UpdateGuid = $DetailsWindow.FindName("txt_UpdateGuid")
        $txt_UpdateClassification = $DetailsWindow.FindName("txt_UpdateClassification")
        $txt_UpdateProduct = $DetailsWindow.FindName("txt_UpdateProduct")
        $txt_UpdateArrivalDate = $DetailsWindow.FindName("txt_UpdateArrivalDate")
        $txt_UpdateIsApproved = $DetailsWindow.FindName("txt_UpdateIsApproved")
        $txt_UpdateIsDeclined = $DetailsWindow.FindName("txt_UpdateIsDeclined")
        $txt_UpdateSupersedes = $DetailsWindow.FindName("txt_UpdateSupersedes")
        $txt_UpdateIsSuperseded = $DetailsWindow.FindName("txt_UpdateIsSuperseded")
        $txt_ApprovalsCount = $DetailsWindow.FindName("txt_ApprovalsCount")
        $dg_Approvals = $DetailsWindow.FindName("dg_Approvals")
        $btn_CloseDetails = $DetailsWindow.FindName("btn_CloseDetails")
    # endregion

    # Retrieve the details of the update from the GUID
    try {
        $Update = Get-WsusUpdateDetails -UpdateGuid $UpdateGuid
    }
    catch {
        [System.Windows.MessageBox]::Show("Could not retrieve details about the update. `r`n`r`nError: $($_)", "Unable to retrieve Update Details",[System.Windows.MessageBoxButton]::OK ,[System.Windows.MessageBoxImage]::Error)
        $DetailsWindow.Close()
    }

    # Populate all of the fields
    $txt_UpdateTitle.Text = $Update.Title
    $txt_UpdateDescription.Text = $Update.Description
    
    
    # Add in a Hyperlink for each URL
    $txt_UpdateUrl.Text = ""
    foreach ($Url in $Update.AdditionalInformationUrls) {
        $NewLink = New-Object System.Windows.Documents.Hyperlink
        $NewLink.NavigateUri = $Url
        $NewLink.Inlines.Add($Url)
        $NewLink.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        $NewLink.add_Click({ Start-Process ($this.NavigateUri) }) 
        $txt_UpdateUrl.AddChild($NewLink)
    }
    
    $txt_UpdateGuid.Text = $UpdateGuid
    $txt_UpdateClassification.Text = $Update.UpdateClassificationTitle
    $txt_UpdateProduct.Text = [String]::Join(", ",@($Update.ProductFamilyTitles)) + " - " + [String]::Join(", ",@($Update.ProductTitles))
    $txt_UpdateArrivalDate.Text = $Update.ArrivalDate
    $txt_UpdateIsApproved.Text = $Update.IsApproved
    $txt_UpdateIsDeclined.Text = $Update.IsDeclined
    $txt_UpdateSupersedes.Text = $Update.HasSupersededUpdates
    $txt_UpdateIsSuperseded.Text = $Update.IsSuperseded
    $txt_ApprovalsCount.Text = @($Update.Approvals).Count
    $dg_Approvals.ItemsSource = @($Update.Approvals)

    $btn_CloseDetails.Add_Click( { $DetailsWindow.Close() } )

    $DetailsWindow.ShowDialog() | Out-Null
}

function Show-WsusUpdateHelpAboutUi {

    # Load the WPF Assemblys
    Add-type -AssemblyName PresentationCore
    Add-type -AssemblyName PresentationFramework
    
    # XAML Window code
    $Reader=(New-Object System.Xml.XmlNodeReader $global:XAMLHelpAbout)  
    $HelpWindow=[Windows.Markup.XamlReader]::Load( $Reader )  
   
    #region BindControls
        $btn_Close = $HelpWindow.FindName("btn_Close")
        $lnk_Author = $HelpWindow.FindName("AuthorLink")
        $lnk_GitHub = $HelpWindow.FindName("GitHubLink")
    # endregion

    $btn_Close.Add_Click( { $HelpWindow.Close() } )

    $lnk_Author.Add_Click({
        Start-Process ($lnk_Author.NavigateUri)
    })

    $lnk_GitHub.Add_Click({
        Start-Process ($lnk_GitHub.NavigateUri)
    })

    $HelpWindow.ShowDialog() | Out-Null
}