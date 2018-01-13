
#region function Copy-MyScreen
function Copy-MyScreen ()
{
  <#
    .SYNOPSIS
      Copies Screen
    .DESCRIPTION
      Copies Screen
    .EXAMPLE
      Copy-MyScreen
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [Switch]$Full
  )
  Write-Verbose -Message "Enter Function Copy-MyScreen"
  
  $LineStringBuilder = New-Object -TypeName System.Text.StringBuilder
  $ScreenStringBuilder = New-Object -TypeName System.Text.StringBuilder
  
  $Width = [System.Console]::WindowWidth - 1
  $End = [System.Console]::CursorTop
  if ($Full.IsPresent)
  {
    $Height = $End
    $Start = 0
  }
  else
  {
    $Height = [System.Console]::WindowHeight - 1
    $Start = $End - $Height
  }
  
  $Screen = $Host.UI.RawUI.GetBufferContents([System.Management.Automation.Host.Rectangle]::new(0, $Start, $Width, $End))
  
  For ($Row = 0; $Row -le $Height; $Row++)
  {
    For ($Column = 0; $Column -le $Width; $Column++)
    {
      [Void]$LineStringBuilder.Append(($Screen[$Row, $Column].Character))
    }
    [Void]$ScreenStringBuilder.AppendLine(($LineStringBuilder.ToString()).Trim())
    [Void]$LineStringBuilder.Clear()
  }
  
  [System.Windows.Forms.Clipboard]::Clear()
  $DataObject = New-Object -TypeName System.Windows.Forms.DataObject
  $DataObject.SetData("Text", ($ScreenStringBuilder.ToString()).Trim())
  [System.Windows.Forms.Clipboard]::SetDataObject($DataObject)
  
  Write-Verbose -Message "Exit Function Copy-MyScreen"
}
#EndRegion 
