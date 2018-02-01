
#region function Get-MyADObject
function Get-MyADObject()
{
  <#
    .SYNOPSIS
      Searches AD and returns an AD SearchResultCollection 
    .DESCRIPTION
      Searches AD and returns an AD SearchResultCollection 
    .PARAMETER LDAPFilter
      AD Search LDAP Filter
    .PARAMETER PageSize
      Search Page Size
    .PARAMETER SizeLimit
      Search Search Size
    .PARAMETER SearchRoot
      Starting Search OU
    .PARAMETER SearchScope
      Search Scope
    .PARAMETER Sort
      Sort Direction
    .PARAMETER SortProperty
      Property to Sort By
    .PARAMETER PropertiesToLoad
      Properties to Load
    .PARAMETER UserName
      User Name to use when searching active directory
    .PARAMETER Password
      Password to use when searching active directory
    .EXAMPLE
      Get-MyADObject [<String>]
    .EXAMPLE
      Get-MyADObject -filter [<String>]
    .NOTES
      Original Function By Ken Sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [String]$LDAPFilter = "(objectClass=*)",
    [Long]$PageSize = 1000,
    [Long]$SizeLimit = 1000,
    [String]$SearchRoot = "LDAP://$($([ADSI]'').distinguishedName)",
    [ValidateSet("Base", "OneLevel", "Subtree")]
    [System.DirectoryServices.SearchScope]$SearchScope = "SubTree",
    [ValidateSet("Ascending", "Descending")]
    [System.DirectoryServices.SortDirection]$Sort = "Ascending",
    [String]$SortProperty,
    [String[]]$PropertiesToLoad,
    [String]$UserName,
    [String]$Password
  )
  Write-Verbose -Message "Enter Function Get-MyADObject"
  Try
  {
    $MySearcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
    $MySearcher.PageSize = $PageSize
    $MySearcher.SizeLimit = $SizeLimit
    $MySearcher.Filter = $LDAPFilter
    Switch -regex ($SearchRoot)
    {
      "GC://*"
      {
        $MySearchRoot = $SearchRoot.ToUpper()
        $MySearcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
        break
      }
      "LDAP://*"
      {
        $MySearchRoot = $SearchRoot.ToUpper()
        $MySearcher.SearchScope = $SearchScope
        break
      }
      Default
      {
        $MySearchRoot = "LDAP://$($SearchRoot.ToUpper())"
        $MySearcher.SearchScope = $SearchScope
        break
      }
    }
    if ($PSBoundParameters.ContainsKey("UserName") -and $PSBoundParameters.ContainsKey("Password"))
    {
      $MySearcher.SearchRoot = New-Object -TypeName System.DirectoryServices.DirectoryEntry($MySearchRoot, $UserName, $Password)
    }
    else
    {
      $MySearcher.SearchRoot = New-Object -TypeName System.DirectoryServices.DirectoryEntry($MySearchRoot)
    }
    if ($PSBoundParameters.ContainsKey("SortProperty"))
    {
      $MySearcher.Sort.PropertyName = $SortProperty
      $MySearcher.Sort.Direction = $Sort
    }
    if ($PSBoundParameters.ContainsKey("PropertiesToLoad"))
    {
      $MySearcher.PropertiesToLoad.AddRange($PropertiesToLoad)
    }
    $MySearcher.FindAll()
    $MySearcher = $Null
    $MySearchRoot = $Null
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
  }
  Catch
  {
    Write-Debug -Message "ErrMsg: $($Error[0].Exception.Message)"
    Write-Debug -Message "Line: $($Error[0].InvocationInfo.ScriptLineNumber)"
    Write-Debug -Message "Code: $(($Error[0].InvocationInfo.Line).Trim())"
  }
  Write-Verbose -Message "Exit Function Get-MyADObject"
}
#endregion

