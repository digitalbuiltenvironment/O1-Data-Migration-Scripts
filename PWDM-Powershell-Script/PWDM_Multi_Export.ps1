Install-Module -Name PWPS_PWDM
Install-Module -Name PWPS_DAB

#region Configuration

$currentDirectory = $PSScriptRoot

# Read the .txt file line by line
$projectIdsFile = "$currentDirectory\project_ids.txt"

$export_path = "D:\OMC2\PWDM" # The top level folder path to export PWDM content for the target connected project
$script_name = "PWDMExportScript" # Name of script, used for logging

#endregion

# Check if the file exists
if (Test-Path -Path $projectIdsFile) {
    $projectIds = Get-Content -Path $projectIdsFile

    # Loop through each project ID from the file
    foreach ($connected_project_id in $projectIds) {

        #region Tokens

        $progress_path = "$export_path\progress\$connected_project_id" # The folder path to store the master progress sheet
        $log_path = "$export_path\logs\$connected_project_id\pwdm_export.log" # The log file path

        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Generating Token" -Verbose

        #$token = Get-IMSNonFederatedAccountToken -UserName bentleyuser03@optimus-pw.com -Password (Read-Host -AsSecureString)
        $token = Get-IMSTokenFromConnectionClient

        #endregion

        #region Build Data Table

        if (!(Test-Path "$progress_path\pwdm_export_progress.xlsx"))
        {
    
            try
            {
                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Building master export list" -Verbose

                #region Files table

                [System.Data.DataTable]$export_progress_table = New-Object -TypeName System.Data.DataTable
                $export_progress_table.TableName = "PWDM_Packages"

                $export_progress_table.Columns.Add("PackageId", "System.Guid")
                $export_progress_table.Columns.Add("PackageType", "System.String")
                $export_progress_table.Columns.Add("Exported", "System.Boolean")

                #region Add Submittals

                $submittal_ids = Get-DMSubmittalIds -ConnectedProjectId $connected_project_id -IMSToken $token -BatchSize 50

                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Adding $(($submittal_ids | Measure-Object).Count) submittals" -Verbose

                foreach ($package_id in $submittal_ids)
                {
                    $new_row = $export_progress_table.NewRow()

                    $new_row.PackageId = $package_id
                    $new_row.PackageType = "Submittal"
                    $new_row.Exported = $false

                    $export_progress_table.Rows.Add($new_row)
                }

                #endregion

                #region Add Transmittals

                $transmittal_ids = Get-DMTransmittalIds -ConnectedProjectId $connected_project_id -IMSToken $token -BatchSize 50

                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Adding $(($transmittal_ids | Measure-Object).Count) transmittals" -Verbose

                foreach ($package_id in $transmittal_ids)
                {
                    $new_row = $export_progress_table.NewRow()

                    $new_row.PackageId = $package_id
                    $new_row.PackageType = "Transmittal"
                    $new_row.Exported = $false

                    $export_progress_table.Rows.Add($new_row)
                }

                #endregion

                #region Add IncomingRFIs

                $incoming_rfi_ids = Get-DMIncomingRFIIds -ConnectedProjectId $connected_project_id -IMSToken $token -BatchSize 50

                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Adding $(($incoming_rfi_ids | Measure-Object).Count) incoming rfis" -Verbose

                foreach ($package_id in $incoming_rfi_ids)
                {
                    $new_row = $export_progress_table.NewRow()

                    $new_row.PackageId = $package_id
                    $new_row.PackageType = "IncomingRFI"
                    $new_row.Exported = $false

                    $export_progress_table.Rows.Add($new_row)
                }

                #endregion

                #region Add OutgoingRFIs

                $outgoing_rfi_ids = Get-DMOutgoingRFIIds -ConnectedProjectId $connected_project_id -IMSToken $token -BatchSize 50

                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Adding $(($outgoing_rfi_ids | Measure-Object).Count) outgoing rfis" -Verbose

                foreach ($package_id in $outgoing_rfi_ids)
                {
                    $new_row = $export_progress_table.NewRow()

                    $new_row.PackageId = $package_id
                    $new_row.PackageType = "OutgoingRFI"
                    $new_row.Exported = $false

                    $export_progress_table.Rows.Add($new_row)
                }

                #endregion

                #region Add IncomingCorrespondences

                $incoming_correspondence_ids = Get-DMIncomingCorrespondenceIds -ConnectedProjectId $connected_project_id -IMSToken $token -BatchSize 50

                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Adding $(($incoming_correspondence_ids | Measure-Object).Count) incoming correspondence" -Verbose

                foreach ($package_id in $incoming_correspondence_ids)
                {
                    $new_row = $export_progress_table.NewRow()

                    $new_row.PackageId = $package_id
                    $new_row.PackageType = "IncomingCorrespondence"
                    $new_row.Exported = $false

                    $export_progress_table.Rows.Add($new_row)
                }

                #endregion

                #region Add OutgoingCorrespondences

                $outgoing_correspondence_ids = Get-DMOutgoingCorrespondenceIds -ConnectedProjectId $connected_project_id -IMSToken $token -BatchSize 50

                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Adding $(($outgoing_correspondence_ids | Measure-Object).Count) outgoing correspondence" -Verbose

                foreach ($package_id in $outgoing_correspondence_ids)
                {
                    $new_row = $export_progress_table.NewRow()

                    $new_row.PackageId = $package_id
                    $new_row.PackageType = "OutgoingCorrespondence"
                    $new_row.Exported = $false

                    $export_progress_table.Rows.Add($new_row)
                }

                #endregion

                #endregion

                #region Metadata table

                [System.Data.DataTable]$metadata_progress_table = New-Object -TypeName System.Data.DataTable
                $metadata_progress_table.TableName = "PWDM_Metadata"

                $metadata_progress_table.Columns.Add("PackageType", "System.String")
                $metadata_progress_table.Columns.Add("Exported", "System.Boolean")

                foreach ($package_type in $($export_progress_table.Rows.PackageType | Sort-Object -Unique))
                {
                    $new_row = $metadata_progress_table.NewRow()

                    $new_row.PackageType = $package_type
                    $new_row.Exported = $false

                    $metadata_progress_table.Rows.Add($new_row)
                }

                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Total $($export_progress_table.Rows.Count) packages" -Verbose

                #endregion

                New-XLSXWorkbook -InputTables $export_progress_table, $metadata_progress_table -OutputFileName "$progress_path\pwdm_export_progress.xlsx"
            }
            catch
            {
                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Error -Message "Failed to build tables" -Verbose
                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Error -Message $_ -Verbose
                Break
            }

        }

        #endregion

        #region Load Latest Progress

        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Loading latest progress" -Verbose

        $progress_files = Get-ChildItem -LiteralPath $progress_path
        $latest_progress_file = $progress_files | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

        $latest_progress = Get-TablesFromXLSXWorkbook $latest_progress_file.FullName

        $latest_files = $latest_progress | Where-Object {$PSItem.TableName -eq "PWDM_Packages"}
        $latest_files_count = ($latest_files.Rows | Measure-Object).Count
        $latest_metadata = $latest_progress | Where-Object {$PSItem.TableName -eq "PWDM_Metadata"}
        $latest_metadata_count = ($latest_metadata.Rows | Measure-Object).Count

        $latest_remaining_files = $latest_files.Rows | Where-Object {$PSItem.Exported -eq $false}
        $latest_remaining_files_count = ($latest_remaining_files | Measure-Object).Count
        $latest_complete_files = $latest_files.Rows | Where-Object {$PSItem.Exported -eq $true}
        $latest_complete_files_count = ($latest_complete_files | Measure-Object).Count

        $latest_remaining_metadata = $latest_metadata.Rows | Where-Object {$PSItem.Exported -eq $false}
        $latest_remaining_metadata_count = ($latest_remaining_metadata | Measure-Object).Count
        $latest_complete_metadata = $latest_metadata.Rows | Where-Object {$PSItem.Exported -eq $true}
        $latest_complete_metadata_count = ($latest_complete_metadata | Measure-Object).Count

        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "$latest_files_count total packages" -Verbose
        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "$latest_complete_files_count completed packages" -Verbose
        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "$latest_remaining_files_count remaining packages" -Verbose
        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "$latest_metadata_count total metadata files" -Verbose
        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "$latest_complete_metadata_count completed metadata files" -Verbose
        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "$latest_remaining_metadata_count remaining metadata files" -Verbose

        #endregion

        #region Export Packages

        $updated_files_count = 0
        $i = 1

        foreach ($package_id in $latest_remaining_files)
        {
            Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Package Id: $($package_id.PackageId) ($i/$latest_remaining_files_count)" -Verbose
            Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Package Type: $($package_id.PackageType)" -Verbose
            $i++

            try
            {
     
                switch ($package_id.PackageType)
                {
                    'Submittal'
                    {
                        $package = Get-DMSubmittalByPackageID -ConnectedProjectId $connected_project_id -IMSToken $token -PackageID $package_id.PackageId
                        Export-DMSubmittalsFiles -SubmittalPackages $package -ExportPath $export_path -ExportInProjectTree
                    }
                    'Transmittal'
                    {
                        $package = Get-DMTransmittalByPackageID -ConnectedProjectId $connected_project_id -IMSToken $token -PackageID $package_id.PackageId
                        Export-DMTransmittalsFiles -TransmittalPackages $package -ExportPath $export_path -ExportInProjectTree
       
                    }
                    'IncomingRFI'
                    {
                        $package = Get-DMIncomingRFIByPackageID -ConnectedProjectId $connected_project_id -IMSToken $token -PackageID $package_id.PackageId
                        Export-DMIncomingRFIsFiles -IncomingRFIPackages $package -ExportPath $export_path -ExportInProjectTree
       
                    }
                    'OutgoingRFI'
                    {
                        $package = Get-DMOutgoingRFIByPackageID -ConnectedProjectId $connected_project_id -IMSToken $token -PackageID $package_id.PackageId
                        Export-DMOutgoingRFIsFiles -OutgoingRFIPackages $package -ExportPath $export_path -ExportInProjectTree
       
                    }
                    'IncomingCorrespondence'
                    {
                        $package = Get-DMIncomingCorrespondenceByPackageID -ConnectedProjectId $connected_project_id -IMSToken $token -PackageID $package_id.PackageId
                        Export-DMIncomingCorrespondenceFiles -IncomingCorrespondencePackages $package -ExportPath $export_path -ExportInProjectTree
       
                    }
                    'OutgoingCorrespondence'
                    {
                        $package = Get-DMOutgoingCorrespondenceByPackageID -ConnectedProjectId $connected_project_id -IMSToken $token -PackageID $package_id.PackageId
                        Export-DMOutgoingCorrespondenceFiles -OutgoingCorrespondencePackages $package -ExportPath $export_path -ExportInProjectTree
       
                    }
                    Default {
                        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Warn -Message "Invalid package type" -Verbose
                    }

                }

                ($latest_files.Rows | Where-Object {$PSItem.PackageId -eq $package_id.PackageId}).Exported = $true

                $updated_files_count++

            }
            catch
            {
                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Error -Message "Failed to export $($package_id.PackageId)" -Verbose
                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Error -Message $_ -Verbose
                Continue
            }
    
        } 

        #endregion

        #region Export Metadata

        $updated_metadata_count = 0
        $i = 1

        foreach ($package_id in $latest_remaining_metadata)
        {
            Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Metadata Type: $($package_id.PackageType) ($i/$latest_remaining_metadata_count)" -Verbose
            $i++
            
            try
            {
     
                switch ($package_id.PackageType)
                {
                    'Submittal'
                    {
                        Export-DMSubmittalsReport -ConnectedProjectId $connected_project_id -IMSToken $token -SubmittalIds $(($latest_files.Rows | Where-Object {$PSItem.PackageType -eq $package_id.PackageType}).PackageId) -ExportPath $export_path -ExportInProjectTree
                    }
                    'Transmittal'
                    {
                        Export-DMTransmittalsReport -ConnectedProjectId $connected_project_id -IMSToken $token -TransmittalIds $(($latest_files.Rows | Where-Object {$PSItem.PackageType -eq $package_id.PackageType}).PackageId) -ExportPath $export_path -ExportInProjectTree
       
                    }
                    'IncomingRFI'
                    {
                        Export-DMIncomingRFIsReport -ConnectedProjectId $connected_project_id -IMSToken $token -IncomingRFIIds $(($latest_files.Rows | Where-Object {$PSItem.PackageType -eq $package_id.PackageType}).PackageId) -ExportPath $export_path -ExportInProjectTree
       
                    }
                    'OutgoingRFI'
                    {
                        Export-DMOutgoingRFIsReport -ConnectedProjectId $connected_project_id -IMSToken $token -OutgoingRFIIds $(($latest_files.Rows | Where-Object {$PSItem.PackageType -eq $package_id.PackageType}).PackageId) -ExportPath $export_path -ExportInProjectTree
       
                    }
                    'IncomingCorrespondence'
                    {
                        Export-DMIncomingCorrespondenceReport -ConnectedProjectId $connected_project_id -IMSToken $token -IncomingCorrespondenceIds $(($latest_files.Rows | Where-Object {$PSItem.PackageType -eq $package_id.PackageType}).PackageId) -ExportPath $export_path -ExportInProjectTree
       
                    }
                    'OutgoingCorrespondence'
                    {
                        Export-DMOutgoingCorrespondenceReport -ConnectedProjectId $connected_project_id -IMSToken $token -OutgoingCorrespondenceIds $(($latest_files.Rows | Where-Object {$PSItem.PackageType -eq $package_id.PackageType}).PackageId) -ExportPath $export_path -ExportInProjectTree
       
                    }
                    Default {
                        Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Warn -Message "Invalid package type" -Verbose
                    }

                }

                ($latest_metadata.Rows | Where-Object {$PSItem.PackageType -eq $package_id.PackageType}).Exported = $true

                $updated_metadata_count++

            }
            catch
            {
                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Error -Message "Failed to export $($package_id.PackageId)" -Verbose
                Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Error -Message $_ -Verbose
                Continue
            }
    
        } 

        #endregion

        #region Update Progress

        if (($updated_files_count -gt 0) -or ($updated_metadata_count -gt 0))
        {
            Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Updating Progress" -Verbose

            New-XLSXWorkbook -InputTables $latest_files,$latest_metadata -OutputFileName "$progress_path\pwdm_export_progress_$(Get-Date -Format "yyyy-MM-dd-HH-mm-ss").xlsx"
        }
        else
        {
            Write-PWPSLog -Path $log_path -Cmdlet $script_name -Level Info -Message "Export for $connected_project_id is complete" -Verbose  
        }

        #endregion

        # Check if it's the last item
        if ($connected_project_id -eq $projectIds[-1]) {
            Write-Host "All projects exported."
        }

    }
} 
else {
    Write-Host "Project IDs file not found."
}
