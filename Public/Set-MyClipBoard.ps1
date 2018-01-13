
#region function Set-MyClipboard
function Set-MyClipboard()
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
    .PARAMETER RowEven
    .PARAMETER RowOdd
    .PARAMETER OfficeFix
    .PARAMETER PassThru
    .EXAMPLE
      Set-MyClipBoard -MyData $MyData -Title "This is My Title"
    .EXAMPLE
      $MyData | Set-MyClipBoard -Title "This is My Title"
    .EXAMPLE
      Set-MyClipBoard -MyData $MyData -Title "This is My Title" -Property "Property1", "Property2", "Property3"
    .EXAMPLE
      Set-MyClipBoard -MyData $MyData -Title "This is My Title" -Columns ([Ordered@{"Property1" = "Left"; "Property2" = "Right"; "Property3" = "Center"})
    .NOTES
      Original Function By Ken Sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [parameter(Mandatory = $True, ValueFromPipeline = $True)]
    [Object[]]$MyData,
    [String]$Title,
    [String]$TitleColor = "DodgerBlue",
    [parameter(Mandatory = $True, ParameterSetName = "Columns")]
    [System.Collections.Specialized.OrderedDictionary]$Columns,
    [parameter(Mandatory = $True, ParameterSetName = "Names")]
    [String[]]$Property,
    [String]$ColumnColor = "SkyBlue",
    [String]$RowEven = "White",
    [String]$RowOdd = "Gainsboro",
    [Switch]$OfficeFix,
    [Switch]$PassThru
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Set-MyClipboard Begin Block"
    
    $HTMLStringBuilder = New-Object -TypeName System.Text.StringBuilder
    
    [Void]$HTMLStringBuilder.Append("Version:1.0`r`nStartHTML:000START`r`nEndHTML:00000END`r`nStartFragment:00FSTART`r`nEndFragment:0000FEND`r`n")
    [Void]$HTMLStringBuilder.Replace("000START", ("{0:X8}" -f $HTMLStringBuilder.Length))
    [Void]$HTMLStringBuilder.Append("<html><head><style>`r`nth { text-align: center; color: black; font: bold; background:$ColumnColor; }`r`ntd:nth-child(1) { text-align:right; }`r`ntable, th, td { border: 1px solid black; border-collapse: collapse; }`r`ntr:nth-child(odd) {background: $RowEven;}`r`ntr:nth-child($RowOdd) {background: gainsboro;}`r`n</style><title>$Title</title></head><body><!--StartFragment-->")
    [Void]$HTMLStringBuilder.Replace("00FSTART", ("{0:X8}" -f $HTMLStringBuilder.Length))
    
    $ObjectList = New-Object -TypeName System.Collections.ArrayList
    
    $GetColumns = $True
    
    Write-Verbose -Message "Exit Function Set-MyClipboard Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Set-MyClipboard Process Block"
    
    if ($GetColumns)
    {
      $Cols = [Ordered]@{ }
      Switch ($PSCmdlet.ParameterSetName)
      {
        "Columns"
        {
          $Cols = $Columns
          Break
        }
        "Names"
        {
          $Property | ForEach-Object -Process { $Cols.Add($PSItem, "Left") }
          Break
        }
        Default
        {
          $MyData[0].PSObject.Properties | Where-Object -FilterScript { $PSItem.MemberType -match "Property|NoteProperty" } | ForEach-Object -Process { $Cols.Add($PSItem.Name, "Left") }
          Break
        }
      }
      $ColNames = @($Cols.Keys)
      $GetColumns = $False
    }
    
    $ObjectList.AddRange(@($MyData | Select-Object -Property $ColNames))
    
    if ($PassThru.IsPresent)
    {
      $MyData
    }
    
    Write-Verbose -Message "Exit Function Set-MyClipboard Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Set-MyClipboard End Block"
    
    if ($OfficeFix.IsPresent)
    {
      if ($PSBoundParameters.ContainsKey("Title"))
      {
        [Void]$HTMLStringBuilder.Append("<table><tr><th style='background:$TitleColor;' colspan='$($Cols.Keys.Count)'>&nbsp;$($Title)&nbsp;</th></tr>")
      }
      else
      {
        [Void]$HTMLStringBuilder.Append("<table>")
      }
      [Void]$HTMLStringBuilder.Append(("<tr style='background:$ColumnColor;'><th>&nbsp;" + ($Cols.Keys -join "&nbsp;</th><th>&nbsp;") + "&nbsp;</th></tr>"))
      $Row = -1
      $RowColor = @($RowEven, $RowOdd)
      ForEach ($Item in $ObjectList)
      {
        [Void]$HTMLStringBuilder.Append("<tr style='background: $($RowColor[($Row = ($Row + 1) % 2)])'>")
        ForEach ($Key in $Cols.Keys)
        {
          [Void]$HTMLStringBuilder.Append("<td style='text-align:$($Cols[$Key])'>&nbsp;$($Item.$Key)&nbsp;</td>")
        }
        [Void]$HTMLStringBuilder.Append("</tr>")
      }
      [Void]$HTMLStringBuilder.Append("</table><br><br>")
    }
    else
    {
      [Void]$HTMLStringBuilder.Append(($ObjectList | ConvertTo-Html -Property $ColNames -Fragment | Out-String))
    }
    [Void]$HTMLStringBuilder.Replace("0000FEND", ("{0:X8}" -f $HTMLStringBuilder.Length))
    [Void]$HTMLStringBuilder.Append("<!--EndFragment--></body></html>")
    [Void]$HTMLStringBuilder.Replace("00000END", ("{0:X8}" -f $HTMLStringBuilder.Length))
    
    [System.Windows.Forms.Clipboard]::Clear()
    $DataObject = New-Object -TypeName System.Windows.Forms.DataObject
    $DataObject.SetData("Text", ($ObjectList | Select-Object -Property $ColNames | ConvertTo-Csv -NoTypeInformation | Out-String))
    $DataObject.SetData("HTML Format", $HTMLStringBuilder.ToString())
    [System.Windows.Forms.Clipboard]::SetDataObject($DataObject)
    
    $ObjectList = $Null
    $HTMLStringBuilder = $Null
    $DataObject = $Null
    $Cols = $Null
    $ColNames = $Null
    $Row = $Null
    $RowColor = $Null
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Verbose -Message "Exit Function Set-MyClipboard End Block"
  }
}
#endregion
