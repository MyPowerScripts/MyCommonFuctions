
#region function Write-MyLogFile
function Write-MyLogFile()
{
  <#
    .SYNOPSIS
    .DESCRIPTION
    .PARAMETER LogFile
    .PARAMETER Severity
    .PARAMETER Message
    .PARAMETER Context
    .PARAMETER Thread
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      Write-MyLogFile -LogFile $LogFile -Message "This is My Info Log File Message"
      Write-MyLogFile -LogFile $LogFile -Severity "Info" -Message "This is My Info Log File Message"
    .EXAMPLE
      Write-MyLogFile -LogFile $LogFile -Severity "Warning" -Message "This is My Warning Log File Message"
    .EXAMPLE
      Write-MyLogFile -LogFile $LogFile -Severity "Error" -Message "This is My Error Log File Message"
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [ValidateScript({ [System.IO.Directory]::Exists([System.IO.Path]::GetDirectoryName($PSItem)) })]
    [String]$LogFile,
    [ValidateSet("Info", "Warning", "Error")]
    [String]$Severity = "Info",
    [parameter(Mandatory = $True)]
    [String]$Message,
    [String]$Context = "",
    [Int]$Thread = $PID
  )
  Write-Verbose -Message "Enter Function Write-MyLogFile"
  
  $TempDate = [DateTime]::Now
  $TempStack = @(Get-PSCallStack)
  Switch ($Severity)
  {
    "Info" { $TempSeverity = 1; Break }
    "Warning" { $TempSeverity = 2; Break }
    "Error" { $TempSeverity = 3; Break }
  }
  Add-Content -Path $LogFile -Value ("<![LOG[{0}]LOG]!><time=`"{1}`" date=`"{2}`" component=`"{3}`" context=`"{4}`" type=`"{5}`" thread=`"{6}`" file=`"{7}`">" -f $Message, $($TempDate.ToString("HH:mm:ss.fff+000")), $($TempDate.ToString("MM-dd-yyyy")), $TempStack[1].Command, $Context, $TempSeverity, $Thread, "$([System.IO.Path]::GetFileName($TempStack[1].ScriptName)):$($TempStack[1].ScriptLineNumber)-$($TempStack.Count - 3)") -Encoding "Ascii"
  
  $TempDate = $Null
  $TempStack = $Null
  $TempSeverity = $Null
  
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  
Write-Verbose -Message "Exit Function Write-MyLogFile"
}
#endregion
