
#region function Decrypt-MyPassword
function Decrypt-MyPassword()
{
  <#
    .SYNOPSIS
      Decrypts a Password for use in a Script
    .DESCRIPTION
      Decrypts a Password for use in a Script
    .PARAMETER Password
      Password to be Decrypted
    .PARAMETER ProtectionScope
      Who can Decrypt
        Currentuser = = Specific User
        LocalMachine = = Any User
    .PARAMETER DecryptKey
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
    [System.Security.Cryptography.DataProtectionScope]$ProtectionScope = "CurrentUser",
    [String]$DecryptKey = $Null
  )
  
  if ([String]::IsNullOrEmpty(([Management.Automation.PSTypeName]"System.Security.Cryptography.ProtectedData").Type))
  {
    Add-Type -AssemblyName System.Security -Debug:$False
  }
  
  Write-Verbose -Message "Enter Function Decrypt-MyPassword"
  
  if ($PSBoundParameters.ContainsKey("DecryptKey"))
  {
    $OptionalEntropy = [System.Text.Encoding]::ASCII.GetBytes($DecryptKey)
  }
  else
  {
    $OptionalEntropy = $Null
  }
  
  $EncryptedData = [System.Convert]::FromBase64String($Password)
  $DecryptedData = [System.Security.Cryptography.ProtectedData]::Unprotect($EncryptedData, $OptionalEntropy, $ProtectionScope)
  [System.Text.Encoding]::ASCII.GetString($DecryptedData)
  
  Write-Verbose -Message "Exit Function Decrypt-MyPassword"
}
#endregion
