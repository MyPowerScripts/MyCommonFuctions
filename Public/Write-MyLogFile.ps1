#region function Write-MyLogFile
function Write-MyLogFile()
{
  <#
    .SYNOPSIS
    .DESCRIPTION
    .PARAMETER LogPath
    .PARAMETER LogFolder
    .PARAMETER LogName
    .PARAMETER Severity
    .PARAMETER Message
    .PARAMETER Context
    .PARAMETER Thread
    .PARAMETER StackInfo
    .PARAMETER MaxSize
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      Write-MyLogFile -LogFolder "MyLogFolder" -Message "This is My Info Log File Message"
      Write-MyLogFile -LogFolder "MyLogFolder" -Severity "Info" -Message "This is My Info Log File Message"
    .EXAMPLE
      Write-MyLogFile -LogFolder "MyLogFolder" -Severity "Warning" -Message "This is My Warning Log File Message"
    .EXAMPLE
      Write-MyLogFile -LogFolder "MyLogFolder" -Severity "Error" -Message "This is My Error Log File Message"
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "LogFolder")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "LogPath")]
    [ValidateScript({ [System.IO.Directory]::Exists($PSItem) })]
    [String]$LogPath = "$($ENV:SystemRoot)\Logs",
    [parameter(Mandatory = $False, ParameterSetName = "LogFolder")]
    [String]$LogFolder = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName),
    [String]$LogName = "$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName)).log",
    [ValidateSet("Info", "Warning", "Error")]
    [String]$Severity = "Info",
    [parameter(Mandatory = $True)]
    [String]$Message,
    [String]$Context = "",
    [Int]$Thread = $PID,
    [Switch]$StackInfo,
    [ValidateRange(0, 16777216)]
    [Int]$MaxSize = 5242880
  )
  Write-Verbose -Message "Enter Function Write-MyLogFile"
  
  if ($PSCmdlet.ParameterSetName -eq "LogPath")
  {
    $TempFile = "$LogPath\$LogName"
  }
  else
  {
    if (-not [System.IO.Directory]::Exists("$LogPath\$LogFolder"))
    {
      [Void][System.IO.Directory]::CreateDirectory("$LogPath\$LogFolder")
    }
    $TempFile = "$LogPath\$LogFolder\$LogName"
  }
  
  Switch ($Severity)
  {
    "Info" { $TempSeverity = 1; Break }
    "Warning" { $TempSeverity = 2; Break }
    "Error" { $TempSeverity = 3; Break }
  }
  
  $TempDate = [DateTime]::Now
  $TempSource = [System.IO.Path]::GetFileName($MyInvocation.ScriptName)
  if ($StackInfo.IsPresent)
  {
    $TempStack = @(Get-PSCallStack)
    $TempCommand = $TempCommand = [System.IO.Path]::GetFileNameWithoutExtension($TempStack[1].Command)
    $TempSource = "Line: $($TempStack[1].ScriptLineNumber) - Scope: $($TempStack.Count - 3)"
  }
  else
  {
    $TempCommand = [System.IO.Path]::GetFileNameWithoutExtension($TempSource)
  }
  
  if ([System.IO.File]::Exists($TempFile) -and $MaxSize -gt 0)
  {
    if (([System.IO.FileInfo]$TempFile).Length -gt $MaxSize)
    {
      $TempBackup = [System.IO.Path]::ChangeExtension($TempFile, "lo_")
      if ([System.IO.File]::Exists($TempBackup))
      {
        Remove-Item -Force -Path $TempBackup
      }
      Rename-Item -Force -Path $TempFile -NewName ([System.IO.Path]::GetFileName($TempBackup))
    }
  }
  Add-Content -Path $TempFile -Value ("<![LOG[{0}]LOG]!><time=`"{1}`" date=`"{2}`" component=`"{3}`" context=`"{4}`" type=`"{5}`" thread=`"{6}`" file=`"{7}`">" -f $Message, $($TempDate.ToString("HH:mm:ss.fff+000")), $($TempDate.ToString("MM-dd-yyyy")), $TempCommand, $Context, $TempSeverity, $Thread, $TempSource)
  
  $TempFile = $Null
  $TempBackup = $Null
  $TempSeverity = $Null
  $TempDate = $Null
  $TempSource = $Null
  $TempCommand = $Null
  $TempStack = $Null
  
  Write-Verbose -Message "Exit Function Write-MyLogFile"
}
#endregion
