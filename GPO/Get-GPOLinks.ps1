function Get-GPOLinks {
    <#
    .SYNOPSIS
    Retrieves the Group Policy Objects (GPOs) from a specified domain, and lists their links and status.

    .DESCRIPTION
    The `Get-GPOLinks` function retrieves GPOs from a specified Active Directory domain. 
    It can filter the GPOs by partial name or retrieve all GPOs. For each GPO, 
    the function outputs its current status, whether it is linked or not, and the action to take 
    (either "To Keep" or "To Delete"). If the `-ExportToXlsx` switch is used, the results are exported 
    to an Excel file using the `ImportExcel` module.

    .PARAMETER Domain
    Specifies the domain from which to retrieve the GPOs. This parameter is mandatory.

    .PARAMETER PartialName
    A partial name to filter the GPOs by. This parameter is optional and is only used 
    if the `-All` switch is not provided.

    .PARAMETER All
    If specified, the function retrieves all GPOs in the domain, regardless of their name. 
    If this switch is omitted, the `-PartialName` parameter must be provided.

    .PARAMETER ExportToXlsx
    If this switch is provided, the results will be exported to an Excel file in the temporary folder 
    of the current user.

    .EXAMPLE
    Get-GPOLinks -Domain "contoso.com" -PartialName "Finance"

    This command retrieves all GPOs in the "contoso.com" domain that have "Finance" in their name.

    .EXAMPLE
    Get-GPOLinks -Domain "contoso.com" -All

    This command retrieves all GPOs in the "contoso.com" domain.

    .EXAMPLE
    Get-GPOLinks -Domain "contoso.com" -PartialName "HR" -ExportToXlsx

    This command retrieves all GPOs in the "contoso.com" domain with "HR" in their name and exports 
    the results to an Excel file.

    .NOTES
    The function requires the `GroupPolicy`, `ActiveDirectory`, and optionally the `ImportExcel` module 
    if the `-ExportToXlsx` switch is used.

    .REQUIRES
    Modules: GroupPolicy, ActiveDirectory, ImportExcel (for export functionality).

    .OUTPUTS
    The function outputs a custom object for each GPO, with properties:
      - Domain: The domain the GPO belongs to.
      - GPOName: The name of the GPO.
      - Status: The current status of the GPO (e.g., "Still Used", "Unused/Unlinked", "Set to Disabled").
      - Action: Recommended action (e.g., "To Keep", "To Delete", "Potentially Delete").
      - LinkPath: The paths to the links where the GPO is applied.

    .COMPONENT
    Active Directory, Group Policy, Excel (optional).

    .LINK
    For more information on Group Policy PowerShell commands, refer to:
    https://docs.microsoft.com/en-us/powershell/module/grouppolicy/

    #>
    
    param (
        [Parameter(Mandatory = $false)]
        [string]$PartialName,

        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [switch]$All,

        [switch]$ExportToXlsx
    )

    # Check if the GroupPolicy module is loaded, if not, load it
    Get-ImportedModule -ModuleName GroupPolicy, ActiveDirectory

    # If -All switch is not provided, PartialName must be specified
    if (-not $All -and -not $PartialName) {
        Write-Error "You must provide either the -All switch or specify a -PartialName."
        return
    }

    # Get the list of GPOs, either all or filtered by partial name
    try {
        if ($All) {
            $gpos = Get-GPO -All -Domain $Domain
        } else {
            $gpos = Get-GPO -All -Domain $Domain | Where-Object { $_.DisplayName -like "*$PartialName*" }
        }
    } catch {
        Write-Error "Failed to retrieve GPOs from the domain."
        return
    }

    # If no GPOs match the partial name
    if (-not $gpos) {
        Write-Host "No GPOs found matching the specified criteria."
        return
    }

    # Array to hold results
    $results = @()

    # Iterate over each matching GPO and get its links
    foreach ($gpo in $gpos) {
        Write-Host "GPO Name: $($gpo.DisplayName)" -ForegroundColor Cyan
        try {
            $gpoReport = Get-GPOReport -id $gpo.Id -ReportType Xml -Domain $Domain -ErrorAction Stop
            $xml = [Xml]$gpoReport
            $GpoLinks = $xml.GPO.LinksTo.SOMPath
        } catch {
            Write-Error "Failed to retrieve report for the GPO: $($gpo.DisplayName)"
            continue
        }

        # Determine GPO status and action
        if (-not $gpoLinks) {
            $status = "Unused/Unlinked"
            $action = "To Delete"
            $linkPaths = "N/A"
        } else {
            # Collect all link paths
            $linkPaths = ($xml | ForEach-Object { $_.GPO.LinksTo.SOMPath }) -join "`n"

            if ($xml.GPO.LinksTo.Enabled -eq "true") {
                $status = "Still Used"
                $action = "To Keep"
            } else {
                $status = "Set to Disabled"
                $action = "Potentially Delete"
            }
        }

        # Create a PSCustomObject for each GPO
        $result = [PSCustomObject]@{
            Domain     = $Domain
            GPOName    = $gpo.DisplayName
            Status     = $status
            Action     = $action
            LinkPath   = $linkPaths
        }

        $result | Format-List

        $results += $result

        Write-Host ""  # Empty line for better readability between GPOs

    }

    # If the ExportToXlsx switch is used, export results to an Excel file
    if ($ExportToXlsx) {
        # Ensure the ImportExcel module is installed
        Get-ImportedModule ImportExcel
        $filePath = "$env:TEMP\GPO_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
        $results | Export-Excel -Path $filePath -AutoSize
        Write-Host "Report exported to $filePath"
    } else {
        # Otherwise, display the results in the console
        $results | Format-Table -AutoSize
    }
}
