
#region function Remove-MyISScriptData
function Remove-MyISScriptData()
{
  <#
    .SYNOPSIS
      Removes Script Data from the Registry
    .DESCRIPTION
      Removes Script Data from the Registry
    .PARAMETER Script
     Name of the Regsitry Key to remove. Defaults to the name of the script.
    .PARAMETER User
      Remove from the HKCU Registry Hive
    .PARAMETER Computer
      Remove from the HKLM Registry Hive
    .PARAMETER Bitness
      Specify 32/64 bit HKLM Registry Hive
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      Remove-MyISScriptData
  
      Removes the default script registry key from the HKCU Registry Hive
    .EXAMPLE
      Remove-MyISScriptData -User -Script "ScriptName"
  
      Removes the Specified Script Name registry key from the HKCU Registry Hive
    .EXAMPLE
      Remove-MyISScriptData -Computer
  
      Removes the default script registry key from the 32/64 bit HKLM Registry Hive
    .EXAMPLE
      Remove-MyISScriptData -Computer -Script "ScriptName" -Bitness "32"
  
      Removes the Specified Script Name registry key from the 32 bit HKLM Registry Hive
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "User")]
  param (
    [String]$Script = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName),
    [parameter(Mandatory = $False, ParameterSetName = "User")]
    [Switch]$User,
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [Switch]$Computer,
    [parameter(Mandatory = $False, ParameterSetName = "Comp")]
    [ValidateSet("32", "64", "All")]
    [String]$Bitness = "All"
  )
  Write-Verbose -Message "Enter Function Remove-MyISScriptData"
  
  # Get Default Registry Paths
  $RegPaths = New-Object -TypeName System.Collections.ArrayList
  if ($Computer.IsPresent)
  {
    if ($Bitness -match "All|32")
    {
      [Void]$RegPaths.Add("Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node")
    }
    if ($Bitness -match "All|64")
    {
      [Void]$RegPaths.Add("Registry::HKEY_LOCAL_MACHINE\Software")
    }
  }
  else
  {
    [Void]$RegPaths.Add("Registry::HKEY_CURRENT_USER\Software")
  }
  
  # Remove the values from the registry.
  ForEach ($RegPath in $RegPaths)
  {
    [Void](Remove-Item -Path "$RegPath\MyISScriptData\$Script")
  }
  
  Write-Verbose -Message "Exit Function Remove-MyISScriptData"
}
#endregion
