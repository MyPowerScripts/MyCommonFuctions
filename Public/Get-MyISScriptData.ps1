
#region function Get-MyISScriptData
function Get-MyISScriptData()
{
  <#
    .SYNOPSIS
      Reads Script Data from the Registry
    .DESCRIPTION
      Reads Script Data from the Registry
    .PARAMETER Script
     Name of the Regsitry Key to read the values from. Defaults to the name of the script.
    .PARAMETER Name
     Name of the Value to read
    .PARAMETER User
      Read from the HKCU Registry Hive
    .PARAMETER Computer
      Read from the HKLM Registry Hive
    .PARAMETER Bitness
      Specify 32/64 bit HKLM Registry Hive
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      $RegValues = Get-MyISScriptData -Name "Name"
  
      Read the value from the HKCU Registry Hive under the Default Script Name registry key
    .EXAMPLE
      $RegValues = Get-MyISScriptData -Name "Name" -User -Script "ScriptName"
  
      Read the value from the HKCU Registry Hive under the Specified Script Name registry key
    .EXAMPLE
      $RegValues = Get-MyISScriptData -Name "Name" -Computer
  
      Read the value from the 64 bit HKLM Registry Hive under the Default Script Name registry key
    .EXAMPLE
      $RegValues = Get-MyISScriptData -Name "Name" -Bitness "32" -Script "ScriptName" -Computer
  
      Read the value from the 32 bit HKLM Registry Hive under the Specified Script Name registry key
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "User")]
  param (
    [String]$Script = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName),
    [String[]]$Name = "*",
    [parameter(Mandatory = $False, ParameterSetName = "User")]
    [Switch]$User,
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [Switch]$Computer,
    [parameter(Mandatory = $False, ParameterSetName = "Comp")]
    [ValidateSet("32", "64")]
    [String]$Bitness = "64"
  )
  Write-Verbose -Message "Enter Function Get-MyISScriptData"
  
  # Get Default Registry Path
  if ($Computer.IsPresent)
  {
    if ($Bitness -eq "64")
    {
      $RegPath = "Registry::HKEY_LOCAL_MACHINE\Software"
    }
    else
    {
      $RegPath = "Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node"
    }
  }
  else
  {
    $RegPath = "Registry::HKEY_CURRENT_USER\Software"
  }
  
  # Get the values from the registry.
  Get-ItemProperty -Path "$RegPath\MyISScriptData\$Script" -ErrorAction "SilentlyContinue" | Select-Object -Property $Name
  
  Write-Verbose -Message "Exit Function Get-MyISScriptData"
}
#endregion
