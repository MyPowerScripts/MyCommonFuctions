
#region function Encrypt-MyPassword
function Encrypt-MyPassword()
{
  <#
    .SYNOPSIS
      Encrypts a Password for use in a Script
    .DESCRIPTION
      Encrypts a Password for use in a Script
    .PARAMETER Password
      Password to be Encrypted
    .PARAMETER ProtectionScope
      Who can Decrypt
        Currentuser = = Specific User
        LocalMachine = = Any User
    .PARAMETER EncryptKey
      Option Extra Encryption Security
    .EXAMPLE
      Encrypt-MyPassword -Password "Password"
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [String]$Password,
    [ValidateSet("LocalMachine", "CurrentUser")]
    [System.Security.Cryptography.DataProtectionScope]$ProtectionScope = "CurrentUser",
    [String]$EncryptKey = $Null
  )
  Write-Verbose -Message "Enter Function Encrypt-MyPassword"
  
  if ([String]::IsNullOrEmpty(([Management.Automation.PSTypeName]"System.Security.Cryptography.ProtectedData").Type))
  {
    Add-Type -AssemblyName System.Security -Debug:$False
  }
  
  if ($PSBoundParameters.ContainsKey("EncryptKey"))
  {
    $OptionalEntropy = [System.Text.Encoding]::ASCII.GetBytes($EncryptKey)
  }
  else
  {
    $OptionalEntropy = $Null
  }
  
  $TempData = [System.Text.Encoding]::ASCII.GetBytes($Password)
  $EncryptedData = [System.Security.Cryptography.ProtectedData]::Protect($TempData, $OptionalEntropy, $ProtectionScope)
  [System.Convert]::ToBase64String($EncryptedData)
  
  Write-Verbose -Message "Exit Function Encrypt-MyPassword"
}
#endregion
