
#region function Decrypt-WithCertificate
function Decrypt-WithCertificate()
{
  <#
    .SYNOPSIS
      Decrypt Encrypted Text using a Certificate
    .DESCRIPTION
      Decrypt Encrypted Text using a Certificate
    .PARAMETER EncryptedText
    .PARAMETER Location
    .PARAMETER Store
    .PARAMETER Subject
    .EXAMPLE
      $DecryptedText = Decrypt-WithCertificate -EncryptedText "String" -Subject "CN=PowerShell Scripts"
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [String]$EncryptedText,
    [ValidateSet("CurrentUser", "LocalMachine")]
    [String]$Location = "LocalMachine",
    [String]$Store = "My",
    [parameter(Mandatory = $True)]
    [String]$Subject = "CN=PowerShell Scripts"
  )
  Write-Verbose -Message "Enter Function Decrypt-WithCertificate"
  
  if (($Cert = @(Get-ChildItem -Path "Cert:\$Location\$Store" | Where-Object -FilterScript { $PSItem.Subject -eq $Subject })).Count)
  {
    [system.text.encoding]::UTF8.GetString($Cert.PrivateKey.Decrypt(([System.Convert]::FromBase64String($EncryptedText)), $True))
  }
  
  Write-Verbose -Message "Exit Function Decrypt-WithCertificate"
}
#endregion

