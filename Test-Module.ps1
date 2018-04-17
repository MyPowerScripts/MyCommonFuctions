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

Get-Command -Module MyCommonFunctions

$VerbosePreference = "SilentlyContinue"

#$ProcessList = Get-Process -Name "Chrome"
#CopyTo-ClipBoard -MyData $ProcessList -Title "My List of Processes" -Columns ([Ordered]@{ "Name" = "Left"; "ID" = "Right"; "StartTime" = "Center" }) -Verbose

Set-MyISScriptData -Name "A1" -Value 1
Set-MyISScriptData -Name "A2" -Value ( ,"One")
Set-MyISScriptData -MultiValue @{ "B1" = 2; "C1" = 3 }
Set-MyISScriptData -MultiValue @{ "B2" = "Two"; "C2" = "Three" }
Set-MyISScriptData -Name "D1" -Value @(1, 2, 3)
Set-MyISScriptData -Name "D2" -Value @("One", "Two", "Three")

Get-MyISScriptData -Name "A1", "B1", "C2", "D2", "D1"

Remove-MyISScriptData
Remove-MyISScriptData
#Remove-MyISScriptData


#$Host.EnterNestedPrompt()

Remove-Module -Name "MyCommonFunctions"
