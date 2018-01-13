
#region function Export-ToExcel
function Export-ToExcel ()
{
  <#
    .SYNOPSIS
      Exports piped output to an excel worksheet
    .DESCRIPTION
      Exports piped output to an excel worksheet
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      Export-ToExcel -Object <Object> -ShowWB
    .EXAMPLE
      <Object>| Export-ToExcel -ShowWB
    .EXAMPLE
      <Object>| Export-ToExcel –ShowWB –FileName "C:\Test.xlsx"
    .EXAMPLE
      <Object>| Export-ToExcel –ShowWB –FileName "C:\Test.xlsx" -Close
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "ShowIt")]
  param (
    [parameter(Mandatory = $True, ValueFromPipeline = $True, HelpMessage = "Enter Object to Export")]
    [Object[]]$Object,
    [string[]]$Property,
    [parameter(Mandatory = $True, ParameterSetName = "SaveIt")]
    [string]$FileName,
    [parameter(ParameterSetName = "SaveIt")]
    [Switch]$NoClose
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Export-ToExcel Begin Block"
    
    $ExcelApp = New-Object -ComObject Excel.Application
    $ExcelWB = $ExcelApp.Workbooks.Add()
    while ($ExcelWB.Worksheets.Count -gt 1)
    {
      $($ExcelWB.Worksheets.Item($ExcelWB.Worksheets.Count)).Delete()
    }
    $ExcelWS = $ExcelWB.Worksheets.Item(1)
    $ExcelWS.Name = "Export-ToExcel"
    $RowNumber = 1
    $ColumnNumber = 0
    $ExcelWS.Rows.Item($RowNumber).Font.Bold = $True
    $ExcelWS.Rows.Item($RowNumber).NumberFormat = "@"
    $ExcelWS.Rows.Item($RowNumber).NumberFormatLocal = "@"
    $PropertyList = @()
    if (($PSCmdlet.ParameterSetName -eq "Showit") -or $NoClose.IsPresent)
    {
      $ExcelApp.Visible = $True
    }
    
    Write-Verbose -Message "Exit Function Export-ToExcel Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Export-ToExcel Process Block"
    
    if ($RowNumber -eq 1)
    {
      if ($PSBoundParameters.ContainsKey("Property"))
      {
        $PropertyList = @($Property)
      }
      else
      {
        $PropertyList = @(@($Object)[0].PSObject.Properties | Where-Object -FilterScript { $PSItem.MemberType -match "Property|NoteProperty" } | Select-Object -ExpandProperty "Name")
      }
      foreach ($Prop in $PropertyList)
      {
        $ExcelWS.Cells.Item($RowNumber, ($ColumnNumber += 1)).Value = $Prop
      }
      $RowNumber += 1
    }
    
    ForEach ($Item in $Object)
    {
      $ColumnNumber = 0
      ForEach ($Prop in $PropertyList)
      {
        $ExcelWS.Cells.Item($RowNumber, ($ColumnNumber += 1)).Value = "$($Item.psobject.Properties[$Prop].Value)"
      }
      $RowNumber += 1
    }
    
    Write-Verbose -Message "Exit Function Export-ToExcel Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Export-ToExcel End Block"
    
    For ($Count = 1; $Count -le $PropertyList.Count; $Count += 1)
    {
      [Void]$ExcelWS.Columns.Item($Count).AutoFit()
    }
    if ($PSBoundParameters.ContainsKey("FileName"))
    {
      if ([System.IO.File]::Exists($FileName))
      {
        [System.IO.File]::Delete($FileName)
      }
      $ExcelWB.SaveAs($FileName)
      if (-not $NoClose.IsPresent)
      {
        $ExcelWB.Close()
        $ExcelApp.Quit()
      }
    }
    
    [Void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWB)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    $ExcelWB = $Null
    $ExcelWS = $Null
    
    [Void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApp)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    $ExcelApp = $Null
    
    Write-Verbose -Message "Exit Function Export-ToExcel End Block"
  }
}
#endregion 
