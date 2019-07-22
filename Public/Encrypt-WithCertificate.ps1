
#region function Encrypt-WithCertificate
function Encrypt-WithCertificate()
{
  <#
    .SYNOPSIS
      Encrypt Text using a Certificate
    .DESCRIPTION
      Encrypt Text using a Certificate
    .PARAMETER Text
    .PARAMETER Location
    .PARAMETER Store
    .PARAMETER Subject
    .EXAMPLE
      $EncryptedText = Encrypt-WithCertificate -Text "String" -Subject "CN=PowerShell Scripts"
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [String]$Text,
    [ValidateSet("CurrentUser", "LocalMachine")]
    [String]$Location = "LocalMachine",
    [String]$Store = "My",
    [parameter(Mandatory = $True)]
    [String]$Subject = "CN=PowerShell Scripts"
  )
  Write-Verbose -Message "Enter Function Encrypt-WithCertificate"
  
  if (($Cert = @(Get-ChildItem -Path "Cert:\$Location\$Store" | Where-Object -FilterScript { $PSItem.Subject -eq $Subject })).Count)
  {
    [System.Convert]::ToBase64String($Cert.PublicKey.Key.Encrypt(([system.text.encoding]::UTF8.GetBytes($Text)), $True))
  }

  Write-Verbose -Message "Exit Function Encrypt-WithCertificate"
}
#endregion

