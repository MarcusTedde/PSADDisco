function Get-ImportedModule {
    <#
    .SYNOPSIS
    Ensures that specified PowerShell modules are imported into the current session.

    .DESCRIPTION
    The `Get-ImportedModule` function checks if a specified list of modules is already imported into the 
    current PowerShell session. If a module is not yet imported, the function will attempt to import it. 
    If the module cannot be imported, an error is thrown. This is useful for ensuring required modules 
    are available before executing any dependent scripts or commands.

    .PARAMETER ModuleName
    Specifies one or more module names that need to be imported into the session. This parameter is mandatory 
    and accepts an array of module names.

    .EXAMPLE
    Get-ImportedModule -ModuleName "ActiveDirectory", "GroupPolicy"

    This command checks if the `ActiveDirectory` and `GroupPolicy` modules are imported into the current session.
    If any of the modules are not loaded, it attempts to import them.

    .EXAMPLE
    Get-ImportedModule -ModuleName "ImportExcel"

    This command checks if the `ImportExcel` module is imported into the current session. 
    If it is not, the function will attempt to import it.

    .NOTES
    If the module is already imported, the function outputs a message indicating that. 
    If it needs to be imported, and the import is successful, a success message is displayed. 
    If the module cannot be found or imported, an error is thrown.

    .OUTPUTS
    The function outputs success or error messages depending on whether the modules were already imported, 
    successfully imported, or failed to import.

    .LINK
    For more information on `Import-Module`, refer to:
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/import-module
    #>

    param (
        [Parameter(Mandatory = $true)]
        [string[]]$ModuleName
    )

    foreach ($module in $ModuleName) {
        # Check if the specified module is already imported
        if (-not (Get-Module -Name $module -ErrorAction SilentlyContinue)) {
            try {
                # Attempt to import the module
                Import-Module -Name $module -ErrorAction Stop
                Write-Host "Successfully imported the $module module." -ForegroundColor Green
            } catch {
                Write-Error "Failed to load the $module module. Please ensure it is installed."
            }
        } else {
            Write-Host "The $module module is already imported." -ForegroundColor Yellow
        }
    }
}
