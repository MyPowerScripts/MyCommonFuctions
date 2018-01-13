
#region function Export-MyExcel
function Export-MyExcel()
{
  <#
    .SYNOPSIS
      Copies Object Data to the ClipBoard
    .DESCRIPTION
      Copies Object Data to the ClipBoard
    .PARAMETER MyData
    .PARAMETER Columns
    .PARAMETER PassThru
    .EXAMPLE
      Export-MyExcel -MyData $MyData
    .EXAMPLE
      $MyData | Export-MyExcel
    .EXAMPLE
      Export-MyExcel -MyData $MyData -Columns "Property1", "Property2", "Property3"
    .EXAMPLE
      $MyData | Export-MyExcel -Columns "Property1", "Property2", "Property3"
    .NOTES
      Original Function By Ken Sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [parameter(Mandatory = $True, ValueFromPipeline = $True)]
    [Object[]]$MyData,
    [String[]]$Property,
    [Switch]$PassThru
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Export-MyExcel Begin Block"
    
    $HTMLStringBuilder = New-Object -TypeName System.Text.StringBuilder
    
    [Void]$HTMLStringBuilder.Append("Version:1.0`r`nStartHTML:000START`r`nEndHTML:00000END`r`nStartFragment:00FSTART`r`nEndFragment:0000FEND`r`n")
    [Void]$HTMLStringBuilder.Replace("000START", ("{0:X8}" -f $HTMLStringBuilder.Length))
    [Void]$HTMLStringBuilder.Append("<html><head><title>$Title</title></head><body><!--StartFragment-->")
    [Void]$HTMLStringBuilder.Replace("00FSTART", ("{0:X8}" -f $HTMLStringBuilder.Length))
    
    $ObjectList = New-Object -TypeName System.Collections.ArrayList
    
    $GetColumns = $True
    
    Write-Verbose -Message "Exit Function Export-MyExcel Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Export-MyExcel Process Block"
    
    if ($GetColumns)
    {
      if ($PSBoundParameters.ContainsKey("Property"))
      {
        $Columns = $Property
      }
      else
      {
        $Columns = $MyData[0].PSObject.Properties | Where-Object -FilterScript { $PSItem.MemberType -match "Property|NoteProperty" } | Select-Object -ExpandProperty "Name"
      }
      $GetColumns = $False
    }
    
    $ObjectList.AddRange(@($MyData | Select-Object -Property $Columns))
    
    if ($PassThru.IsPresent)
    {
      $MyData
    }
    
    Write-Verbose -Message "Exit Function Export-MyExcel Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Export-MyExcel End Block"
    
    [Void]$HTMLStringBuilder.Append(($ObjectList | ConvertTo-Html -Property $Columns -Fragment | Out-String))
    
    [Void]$HTMLStringBuilder.Replace("0000FEND", ("{0:X8}" -f $HTMLStringBuilder.Length))
    [Void]$HTMLStringBuilder.Append("<!--EndFragment--></body></html>")
    [Void]$HTMLStringBuilder.Replace("00000END", ("{0:X8}" -f $HTMLStringBuilder.Length))
    
    [System.Windows.Forms.Clipboard]::Clear()
    $DataObject = New-Object -TypeName System.Windows.Forms.DataObject
    $DataObject.SetData("HTML Format", $HTMLStringBuilder.ToString())
    [System.Windows.Forms.Clipboard]::SetDataObject($DataObject)
    
    
    $ExcelApp = New-Object -comobject Excel.Application
    $ExcelWB = $ExcelApp.Workbooks.Add()
    while ($ExcelWB.Worksheets.Count -gt 1)
    {
      $($ExcelWB.Worksheets.Item($ExcelWB.Worksheets.Count)).Delete()
    }
    $ExcelWS = $ExcelWB.Worksheets.Item(1)
    $ExcelWS.Name = "Export-MyExcel"
    $RowNumber = 1
    $ColumnNumber = 1
    $ExcelWS.Rows.Item($RowNumber).Font.Bold = $True
    $ExcelWS.Rows.Item($RowNumber).NumberFormat = "@"
    $ExcelWS.Rows.Item($RowNumber).NumberFormatLocal = "@"
    $ExcelApp.Visible = $True
    
    $ExcelWS.Paste()
    
    [System.Windows.Forms.Clipboard]::Clear()
    
    $ObjectList = $Null
    $HTMLStringBuilder = $Null
    $DataObject = $Null
    $Columns = $Null
    $RowNumber = $Null
    $ColumnNumber = $Null
    
    [Void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWB)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    $ExcelWS = $Null
    $ExcelWB = $Null
    
    [Void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApp)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    $ExcelApp = $Null
    
    Write-Verbose -Message "Exit Function Export-MyExcel End Block"
  }
}
#endregion
