
#region function Set-MyISScriptData
function Set-MyISScriptData()
{
  <#
    .SYNOPSIS
      Writes Script Data to the Registry
    .DESCRIPTION
      Writes Script Data to the Registry
    .PARAMETER Script
     Name of the Regsitry Key to write the values under. Defaults to the name of the script.
    .PARAMETER Name
     Name of the Value to write
    .PARAMETER Value
      The Data to write
    .PARAMETER MultiValue
      Write Multiple values to the Registry
    .PARAMETER User
      Write to the HKCU Registry Hive
    .PARAMETER Computer
      Write to the HKLM Registry Hive
    .PARAMETER Bitness
      Specify 32/64 bit HKLM Registry Hive
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      Set-MyISScriptData -Name "Name" -Value "Value"
  
      Write REG_SZ value to the HKCU Registry Hive under the Default Script Name registry key
    .EXAMPLE
      Set-MyISScriptData -Name "Name" -Value @("This", "That") -User -Script "ScriptName"
  
      Write REG_MULTI_SZ value to the HKCU Registry Hive under the Specified Script Name registry key
  
      Single element arrays will be written as REG_SZ. To ensure they are written as REG_MULTI_SZ
      Use @() or (,) when specifing the Value paramter value
    .EXAMPLE
      Set-MyISScriptData -Name "Name" -Value (,8) -Bitness "64" -Computer
  
      Write REG_MULTI_SZ value to the 64 bit HKLM Registry Hive under the Default Script Name registry key
  
      Number arrays are written to the registry as strings.
    .EXAMPLE
      Set-MyISScriptData -Name "Name" -Value 0 -Computer
  
      Write REG_DWORD value to the HKLM Registry Hive under the Default Script Name registry key
    .EXAMPLE
      Set-MyISScriptData -MultiValue @{"Name" = "MyName"; "Number" = 4; "Array" = @("First", 2, "3rd", 4)} -Computer -Bitness "32"
  
      Write multiple values to the 32 bit HKLM Registry Hive under the Default Script Name registry key
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "User")]
  param (
    [String]$Script = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName),
    [parameter(Mandatory = $True, ParameterSetName = "User")]
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [String]$Name,
    [parameter(Mandatory = $True, ParameterSetName = "User")]
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [Object]$Value,
    [parameter(Mandatory = $True, ParameterSetName = "UserMulti")]
    [parameter(Mandatory = $True, ParameterSetName = "CompMulti")]
    [HashTable]$MultiValue,
    [parameter(Mandatory = $False, ParameterSetName = "User")]
    [parameter(Mandatory = $False, ParameterSetName = "UserMulti")]
    [Switch]$User,
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [parameter(Mandatory = $True, ParameterSetName = "CompMulti")]
    [Switch]$Computer,
    [parameter(Mandatory = $False, ParameterSetName = "Comp")]
    [parameter(Mandatory = $False, ParameterSetName = "CompMulti")]
    [ValidateSet("32", "64", "All")]
    [String]$Bitness = "All"
  )
  Write-Verbose -Message "Enter Function Set-MyISScriptData"
  
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
  
  # Create the Registry Keys if Needed.
  ForEach ($RegPath in $RegPaths)
  {
    if ([String]::IsNullOrEmpty((Get-Item -Path "$RegPath\MyISScriptData" -ErrorAction "SilentlyContinue")))
    {
      Try
      {
        [Void](New-Item -Path $RegPath -Name "MyISScriptData")
      }
      Catch
      {
        Throw "Error Creating Registry Key : MyISScriptData"
      }
    }
    if ([String]::IsNullOrEmpty((Get-Item -Path "$RegPath\MyISScriptData\$Script" -ErrorAction "SilentlyContinue")))
    {
      Try
      {
        [Void](New-Item -Path "$RegPath\MyISScriptData" -Name $Script)
      }
      Catch
      {
        Throw "Error Creating Registry Key : $Script"
      }
    }
  }
  
  # Write the values to the registry.
  Switch -regex ($PSCmdlet.ParameterSetName)
  {
    "Multi"
    {
      ForEach ($Key in $MultiValue.Keys)
      {
        if ($MultiValue[$Key] -is [Array])
        {
          $Data = [String[]]$MultiValue[$Key]
        }
        else
        {
          $Data = $MultiValue[$Key]
        }
        ForEach ($RegPath in $RegPaths)
        {
          [Void](Set-ItemProperty -Path "$RegPath\MyISScriptData\$Script" -Name $Key -Value $Data)
        }
      }
    }
    Default
    {
      if ($Value -is [Array])
      {
        $Data = [String[]]$Value
      }
      else
      {
        $Data = $Value
      }
      ForEach ($RegPath in $RegPaths)
      {
        [Void](Set-ItemProperty -Path "$RegPath\MyISScriptData\$Script" -Name $Name -Value $Data)
      }
    }
  }
  
  Write-Verbose -Message "Exit Function Set-MyISScriptData"
}
#endregion
