
#region function Test-MyClassLoaded
function Test-MyClassLoaded()
{
  <#
    .SYNOPSIS
      Test if Custom Class is Loaded
    .DESCRIPTION
      Test if Custom Class is Loaded
    .PARAMETER Name
      Name of Custom Class
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      $IsLoaded = Test-MyClassLoaded -Name "CustomClass"
    .NOTES
      Original Function By Ken Sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "Default")]
    [String]$Name
  )
  Write-Verbose -Message "Enter Function Test-MyClassLoaded"
  
  (-not [String]::IsNullOrEmpty(([Management.Automation.PSTypeName]$Name).Type))
  
  Write-Verbose -Message "Exit Function Test-MyClassLoaded"
}
#endregion
