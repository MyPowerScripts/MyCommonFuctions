<#
  .SYNOPSIS
    My Common PowerShell Functions
  .DESCRIPTION
    My Common PowerShell Functions
  .EXAMPLE
    Import-Module -Name "MyCommonFunctions"
  .EXAMPLE
    Import-Module -Name "D:\MyTest\MyCommonFunctions\MyCommonFunctions.psm1"
  .NOTES
    Original Script By Ken Sweet on 10/15/2017 at 06:53 AM
  .LINK
#>

$PSModule = $ExecutionContext.SessionState.Module
$PSModuleRoot = $PSModule.ModuleBase

Write-Verbose -Message "Loading Module = $($PSModule.Name)"
Write-Verbose -Message "Module Base = $($PSModule.ModuleBase)"
Write-Verbose -Message ""

$Params = @{ }

#region ******** Update Format & Type Data ********
if ([System.IO.Directory]::Exists("$PSModuleRoot\TypeData"))
{
  Write-Verbose -Message "Updating Format and Type Data"
  # Update Format Data
  ForEach ($FormatData in [System.IO.Directory]::EnumerateFiles("$PSModuleRoot\TypeData", "*.Format.ps1xml"))
  {
    Write-Verbose -Message $FormatData
    Update-FormatData -AppendPath $FormatData -ErrorAction "SilentlyContinue"
  }
  
  # Update Type Data
  ForEach ($TypeData in [System.IO.Directory]::EnumerateFiles("$PSModuleRoot\TypeData", "*.Types.ps1xml"))
  {
    Write-Verbose -Message $TypeData
    Update-TypeData -AppendPath $TypeData -ErrorAction "SilentlyContinue"
  }
  Write-Verbose -Message ""
}
#endregion

#region ******** Load Private Commands ********
if ([System.IO.Directory]::Exists("$PSModuleRoot\Private"))
{
  Write-Verbose -Message "Loading Private Functions"
  ForEach ($Command in [System.IO.Directory]::EnumerateFiles("$PSModuleRoot\Private", "*.ps1"))
  {
    Write-Verbose -Message $Command
    . $Command
  }
  Write-Verbose -Message ""
}
#endregion

#region ******** Load Public Commands ********
$Functions = @()
if ([System.IO.Directory]::Exists("$PSModuleRoot\Public"))
{
  Write-Verbose -Message "Loading Public Functions"
  ForEach ($Command in [System.IO.Directory]::EnumerateFiles("$PSModuleRoot\Public", "*.ps1"))
  {
    Write-Verbose -Message $Command
    $Functions += [System.IO.Path]::GetFileNameWithoutExtension($Command)
    . $Command
  }
  Write-Verbose -Message ""
}
if ($Functions.Count)
{
  $Params.Function = $Functions
}
#endregion

#region ******** Load Public Variables ********
Write-Verbose -Message "Loading Public Variables"
$Variables = @()
Write-Verbose -Message ""
if ($Variables.Count)
{
  $Params.Variable = $Variables
}
#endregion

Export-ModuleMember @Params

Write-Verbose -Message ""
Write-Verbose -Message "Finished Loading Module = $($PSModule.Name)"
Write-Verbose -Message ""
