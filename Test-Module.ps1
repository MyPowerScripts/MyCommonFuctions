<#
  .SYNOPSIS
    Test Script for MyCommonFunctions Module
  .DESCRIPTION
    Test Script for MyCommonFunctions Module
  .EXAMPLE
    Test-Module.ps1
  .NOTES
    Original Script By Ken Sweet on 10/15/2017 at 06:53 AM
  .LINK
#>

$VerbosePreference = "Continue"

$ErrorActionPreference = "Stop"

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#Explicitly import the module for testing
Import-Module -Name "$PWD\MyCommonFunctions.psm1" | Out-Null

$VerbosePreference = "SilentlyContinue"

#$ProcessList = Get-Process -Name "Chrome"
#CopyTo-ClipBoard -MyData $ProcessList -Title "My List of Processes" -Columns ([Ordered]@{ "Name" = "Left"; "ID" = "Right"; "StartTime" = "Center" }) -Verbose

$Host.EnterNestedPrompt()
