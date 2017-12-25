
#region function Set-MyClipBoard
function Set-MyClipBoard()
{
  <#
    .SYNOPSIS
      Copies Object Data to the ClipBoard
    .DESCRIPTION
      Copies Object Data to the ClipBoard
    .PARAMETER MyData
    .PARAMETER Title
    .PARAMETER TitleColor
    .PARAMETER Columns
    .PARAMETER ColumnColor
    .PARAMETER Row0
    .PARAMETER Row1
    .EXAMPLE
      Set-MyClipBoard -MyData $MyData -Title "This is My Title"
    .EXAMPLE
      Set-MyClipBoard -MyData $MyData -Title "This is My Title" -Columns ([Ordered@{"Property1" = "Left"; "Property2" = "Right"; "Property3" = "Center"})
    .NOTES
      Original Function By ken.sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "NoTitle")]
  param (
    [parameter(Mandatory = $True)]
    [Object[]]$MyData,
    [parameter(Mandatory = $True, ParameterSetName = "Title")]
    [String]$Title,
    [parameter(Mandatory = $False, ParameterSetName = "Title")]
    [String]$TitleColor = "DodgerBlue",
    [System.Collections.Specialized.OrderedDictionary]$Columns,
    [String]$ColumnColor = "SkyBlue",
    [String]$Row0 = "White",
    [String]$Row1 = "Gainsboro",
    [Switch]$OfficeFix
  )
  Write-Verbose -Message "Enter Function Set-MyClipBoard"
  
  if (-not $PSBoundParameters.ContainsKey("Columns"))
  {
    $Columns = [Ordered]@{ }
    $MyData[0].PSObject.Properties | Where-Object -FilterScript { $PSItem.MemberType -match "Property|NoteProperty" } | ForEach-Object -Process { $Columns.Add($PSItem.Name, "Left") }
  }
  
  $HTMLStringBuilder = New-Object -TypeName System.Text.StringBuilder
  
  [Void]$HTMLStringBuilder.Append("Version:1.0`r`nStartHTML:000START`r`nEndHTML:00000END`r`nStartFragment:00FSTART`r`nEndFragment:0000FEND`r`n")
  [Void]$HTMLStringBuilder.Replace("000START", ("{0:X8}" -f $HTMLStringBuilder.Length))
  [Void]$HTMLStringBuilder.Append("<html><head><style>`r`nth { text-align: center; color: black; font: bold; background:$ColumnColor; }`r`ntd:nth-child(1) { text-align:right; }`r`ntable, th, td { border: 1px solid black; border-collapse: collapse; }`r`ntr:nth-child(odd) {background: $Row0;}`r`ntr:nth-child($Row1) {background: gainsboro;}`r`n</style><title>$Title</title></head><body><!--StartFragment-->")
  [Void]$HTMLStringBuilder.Replace("00FSTART", ("{0:X8}" -f $HTMLStringBuilder.Length))
  
  if ($OfficeFix.IsPresent)
  {
    if ($PSBoundParameters.ContainsKey("Title"))
    {
      [Void]$HTMLStringBuilder.Append("<table><tr><th style='background:$TitleColor;' colspan='$($Columns.Keys.Count)'>&nbsp;$($Title)&nbsp;</th></tr>")
    }
    else
    {
      [Void]$HTMLStringBuilder.Append("<table>")
    }
    [Void]$HTMLStringBuilder.Append(("<tr style='background:$ColumnColor;'><th>&nbsp;" + ($Columns.Keys -join "&nbsp;</th><th>&nbsp;") + "&nbsp;</th></tr>"))
    $Row = -1
    $RowColor = @($Row0, $Row1)
    ForEach ($Item in $MyData)
    {
      [Void]$HTMLStringBuilder.Append("<tr style='background: $($RowColor[($Row = ($Row + 1) % 2)])'>")
      ForEach ($Key in $Columns.Keys)
      {
        [Void]$HTMLStringBuilder.Append("<td style='text-align:$($Columns[$Key])'>&nbsp;$($Item.$Key)&nbsp;</td>")
      }
      [Void]$HTMLStringBuilder.Append("&nbsp;</td></tr>")
    }
    [Void]$HTMLStringBuilder.Append("</table><br><br>")
  }
  else
  {
    [Void]$HTMLStringBuilder.Append(($MyData | ConvertTo-Html -Property $($Columns.Keys) -Fragment | Out-String))
  }
  [Void]$HTMLStringBuilder.Replace("0000FEND", ("{0:X8}" -f $HTMLStringBuilder.Length))
  [Void]$HTMLStringBuilder.Append("<!--EndFragment--></body></html>")
  [Void]$HTMLStringBuilder.Replace("00000END", ("{0:X8}" -f $HTMLStringBuilder.Length))
  
  [System.Windows.Forms.Clipboard]::Clear()
  $DataObject = New-Object -TypeName System.Windows.Forms.DataObject
  $DataObject.SetData("Text", ($MyData | Select-Object -Property $($Columns.Keys) | ConvertTo-Csv -NoTypeInformation | Out-String))
  $DataObject.SetData("HTML Format", $HTMLStringBuilder.ToString())
  [System.Windows.Forms.Clipboard]::SetDataObject($DataObject)
  
  Write-Verbose -Message "Exit Function Set-MyClipBoard"
}
#endregion
