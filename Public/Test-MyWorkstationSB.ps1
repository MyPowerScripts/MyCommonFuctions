
#region function Test-MyWorkstationSB
function Test-MyWorkstationSB()
{
  <#
    .SYNOPSIS
      Verify Remote Workstation is the Correct One
    .DESCRIPTION
      Verify Remote Workstation is the Correct One
    .PARAMETER ComputerName
      Name of the Computer to Verify
    .PARAMETER Credential
      Credentials to use when connecting to the Remote Computer
    .PARAMETER Wait
      How Long to Wait for Job to be Completed
    .PARAMETER Serial
      Return Serial Number
    .PARAMETER Mobile
      Check if System is Desktop / Laptop
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      Test-MyWorkstationSB -ComputerName "MyWorkstation"
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $False, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
    [String[]]$ComputerName = [System.Environment]::MachineName,
    [PSCredential]$Credential,
    [ValidateRange(30, 300)]
    [Int]$Wait = 120,
    [Switch]$Serial,
    [Switch]$Mobile
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Test-MyWorkstationSB"
    
    # Primay Job Script to Verify Computer Name and WMI Functionality and Return Commonly used Properties
    # Accepts ComputerName and Credentials as Parameters to make connection to remote computer
    #region ******** $MyCompDataJob ScriptBlock ********
    $MyCompDataJob = {
      param (
        [HashTable]$Params
      )
      $ErrorActionPreference = "Stop"
      Try
      {
        Write-Output -InputObject (Get-WmiObject @Params -Class Win32_ComputerSystem | Select-Object -Property "UserName", "Name", "PartOfDomain", "Domain", "Manufacturer", "Model", "TotalPhysicalMemory", @{ "Name" = "Error"; "Expression" = { $False } }, @{ "Name" = "ErrorMessage"; "Expression" = { "" } })
      }
      Catch
      {
        Write-Output -InputObject ([PSCustomObject]@{ "Error" = $True; "ErrorMessage" = $($Error[0]) })
      }
    }
    #endregion
    
    # Secondary Job Script to Verify WMI Functionality and Return Commonly used Properties
    # Accepts ComputerName and Credentials as Parameters to make connection to remote computer
    #region ******** $MyOSDataJob ScriptBlock ********
    $MyOSDataJob = {
      param (
        [HashTable]$Params
      )
      $ErrorActionPreference = "Stop"
      Try
      {
        Write-Output -InputObject (Get-WmiObject @Params -Class Win32_OperatingSystem | Select-Object -Property "Caption", "CSDVersion", "OSArchitecture", "ProductType", "LocalDateTime", "InstallDate", "LastBootUpTime", @{ "Name" = "Error"; "Expression" = { $False } }, @{ "Name" = "ErrorMessage"; "Expression" = { "" } })
      }
      Catch
      {
        Write-Output -InputObject ([PSCustomObject]@{ "Error" = $True; "ErrorMessage" = $($Error[0]) })
      }
    }
    #endregion
    
    # Optional SerialNumber Job Script to Return Computer Serial Number
    # Accepts ComputerName and Credentials as Parameters to make connection to remote computer
    #region ******** $MyBIOSDataJob ScriptBlock ********
    $MyBIOSDataJob = {
      param (
        [HashTable]$Params
      )
      $ErrorActionPreference = "Stop"
      Try
      {
        Write-Output -InputObject (Get-WmiObject @Params -Class Win32_Bios | Select-Object -Property "SerialNumber", @{ "Name" = "Error"; "Expression" = { $False } }, @{ "Name" = "ErrorMessage"; "Expression" = { "" } })
      }
      Catch
      {
        Write-Output -InputObject ([PSCustomObject]@{ "Error" = $True; "ErrorMessage" = $($Error[0]) })
      }
    }
    #endregion
    
    # Optional Mobile / ChassiType Job Script to Return Computer ChassisType (Desktop / Laptop)
    # Accepts ComputerName and Credentials as Parameters to make connection to remote computer
    #region ******** $MyChassisDataJob ScriptBlock ********
    $MyChassisDataJob = {
      param (
        [HashTable]$Params
      )
      $ErrorActionPreference = "Stop"
      Try
      {
        Write-Output -InputObject ([PSCustomObject]@{ "IsMobile" = $((@(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32) -contains (((Get-WmiObject @Params -Class Win32_SystemEnclosure).ChassisTypes)[0]))); "Error" = $False; "ErrorMessage" = "" })
      }
      Catch
      {
        Write-Output -InputObject ([PSCustomObject]@{ "Error" = $True; "ErrorMessage" = $($Error[0]) })
      }
    }
    #endregion
  }
  Process
  {
    Write-Verbose -Message "Enter Function Test-MyWorkstationSB - Process"
    
    ForEach ($Computer in $ComputerName)
    {
      # Used to Calculate Verify Time
      $StartTime = [DateTime]::Now
      
      # Default Custom Object for the Verify Function to Return, Since it will always return a value I create the Object with the default error / failure values and update the poperties as needed
      #region ******** Custom Return Object $VerifyObject ********
      $VerifyObject = [PSCustomObject]@{
        "ComputerName" = $Computer.ToUpper();
        "Found" = $False;
        "UserName" = "";
        "Domain" = "";
        "DomainMember" = "";
        "ProductType" = 0;
        "Manufacturer" = "";
        "Model" = "";
        "IsMobile" = $False;
        "SerialNumber" = "";
        "Memory" = "";
        "OperatingSystem" = "";
        "ServicePack" = "";
        "Architecture" = "";
        "LocalDateTime" = "";
        "InstallDate" = "";
        "LastBootUpTime" = "";
        "IPAddress" = "";
        "Status" = "Off-Line";
        "Time" = [TimeSpan]::Zero;
        "ErrorMessage" = ""
      }
      #endregion
      
      if ($Computer -match "^(([a-zA-Z]|[a-zA-Z][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z]|[A-Za-z][A-Za-z0-9\-]*[A-Za-z0-9])$")
      {
        Try
        {
          # Get IP Address from DNS, you want to do all remote checks using IP rather than ComputerName.  If you connect to a computer using the wrong name Get-WmiObject will fail and using the IP Address will not
          $IPAddresses = @([System.Net.Dns]::GetHostAddresses($Computer) | Where-Object -FilterScript { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | Select-Object -ExpandProperty IPAddressToString)
          ForEach ($IPAddress in $IPAddresses)
          {
            # I think this is Faster than using Test-Connection
            if (((New-Object -TypeName System.Net.NetworkInformation.Ping).Send($IPAddress)).Status -eq [System.Net.NetworkInformation.IPStatus]::Success)
            {
              # Default Common Get-WmiObject Options
              if ($PSBoundParameters.ContainsKey("Credential"))
              {
                $Params = @{
                  "ComputerName" = $IPAddress;
                  "Credential" = $Credential
                }
              }
              else
              {
                $Params = @{
                  "ComputerName" = $IPAddress
                }
              }
              
              # Start Setting Return Values as they are Found
              $VerifyObject.Status = "On-Line"
              $VerifyObject.IPAddress = $IPAddress
              
              # Start Primary Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
              [Void]($MyJob = Start-Job -ArgumentList $Params -ScriptBlock $MyCompDataJob)
              # Wait for Job to Finish or Wait Time has Elasped
              [Void](Wait-Job -Job $MyJob -Timeout $Wait)
              
              # Check if Job is Complete and has Data
              if ($MyJob.State -eq "Completed" -and $MyJob.HasMoreData)
              {
                # Get Job Data
                $MyCompData = Get-Job -ID $MyJob.ID | Receive-Job -AutoRemoveJob -Wait -Force
                if ($MyCompData.Error)
                {
                  $VerifyObject.Status = "Verify Comp Error"
                  $VerifyObject.ErrorMessage = $MyCompData.ErrorMessage
                }
                else
                {
                  # Set Found Properties
                  $VerifyObject.ComputerName = "$($MyCompData.Name)"
                  $VerifyObject.UserName = "$($MyCompData.UserName)"
                  $VerifyObject.Domain = "$($MyCompData.Domain)"
                  $VerifyObject.DomainMember = "$($MyCompData.PartOfDomain)"
                  $VerifyObject.Manufacturer = "$($MyCompData.Manufacturer)"
                  $VerifyObject.Model = "$($MyCompData.Model)"
                  $VerifyObject.Memory = "$($MyCompData.TotalPhysicalMemory)"
                  
                  # Verify Remote Computer is the Connect Computer, No need to get any more information
                  if ($MyCompData.Name -eq @($Computer.Split(".", [System.StringSplitOptions]::RemoveEmptyEntries))[0])
                  {
                    # Found Corrct Workstation
                    $VerifyObject.Found = $True
                    
                    # Start Secondary Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                    [Void]($MyJob = Start-Job -ArgumentList $Params -ScriptBlock $MyOSDataJob)
                    # Wait for Job to Finish or Wait Time has Elasped
                    [Void](Wait-Job -Job $MyJob -Timeout $Wait)
                    
                    # Check if Job is Complete and has Data
                    if ($MyJob.State -eq "Completed" -and $MyJob.HasMoreData)
                    {
                      # Get Job Data
                      $MyOSData = Get-Job -ID $MyJob.ID | Receive-Job -AutoRemoveJob -Wait -Force
                      if ($MyOSData.Error)
                      {
                        $VerifyObject.Status = "Verify Operating System Error"
                        $VerifyObject.ErrorMessage = $MyOSData.ErrorMessage
                      }
                      else
                      {
                        # Set Found Properties
                        $VerifyObject.ProductType = $MyOSData.ProductType
                        $VerifyObject.OperatingSystem = "$($MyOSData.Caption)"
                        $VerifyObject.ServicePack = "$($MyOSData.CSDVersion)"
                        $VerifyObject.Architecture = $(if ([String]::IsNullOrEmpty($MyOSData.OSArchitecture)) { "32-bit" } else { "$($MyOSData.OSArchitecture)" })
                        $VerifyObject.LocalDateTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.LocalDateTime)
                        $VerifyObject.InstallDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.InstallDate)
                        $VerifyObject.LastBootUpTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.LastBootUpTime)
                        
                        # Optional SerialNumber Job
                        if ($Serial)
                        {
                          # Start Optional Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                          [Void]($MyJob = Start-Job -ArgumentList $Params -ScriptBlock $MyBIOSDataJob)
                          # Wait for Job to Finish or Wait Time has Elasped
                          [Void](Wait-Job -Job $MyJob -Timeout $Wait)
                          
                          # Check if Job is Complete and has Data
                          if ($MyJob.State -eq "Completed" -and $MyJob.HasMoreData)
                          {
                            # Get Job Data
                            $MyBIOSData = Get-Job -ID $MyJob.ID | Receive-Job -AutoRemoveJob -Wait -Force
                            if ($MyBIOSData.Error)
                            {
                              $VerifyObject.Status = "Verify SerialNumber Error"
                              $VerifyObject.ErrorMessage = $MyBIOSData.ErrorMessage
                            }
                            else
                            {
                              # Set Found Property
                              $VerifyObject.SerialNumber = "$($MyBIOSData.SerialNumber)"
                            }
                          }
                          else
                          {
                            $VerifyObject.Status = "Verify SerialNumber Error"
                            [Void](Remove-Job -Job $MyJob -Force)
                          }
                        }
                        
                        # Optional Mobile / ChassisType Job
                        if ($Mobile)
                        {
                          # Start Optional Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                          [Void]($MyJob = Start-Job -ArgumentList $Params -ScriptBlock $MyChassisDataJob)
                          # Wait for Job to Finish or Wait Time has Elasped
                          [Void](Wait-Job -Job $MyJob -Timeout $Wait)
                          
                          # Check if Job is Complete and has Data
                          if ($MyJob.State -eq "Completed" -and $MyJob.HasMoreData)
                          {
                            # Get Job Data
                            $MyChassisData = Get-Job -ID $MyJob.ID | Receive-Job -AutoRemoveJob -Wait -Force
                            if ($MyChassisData.Error)
                            {
                              $VerifyObject.Status = "Verify is Mobile Error"
                              $VerifyObject.ErrorMessage = $MyChassisData.ErrorMessage
                            }
                            else
                            {
                              # Set Found Property
                              $VerifyObject.IsMobile = $($MyChassisData.IsMobile)
                            }
                          }
                          else
                          {
                            $VerifyObject.Status = "Verify is Mobile Error"
                            [Void](Remove-Job -Job $MyJob -Force)
                          }
                        }
                      }
                    }
                    else
                    {
                      $VerifyObject.Status = "Verify Operating System Error"
                      [Void](Remove-Job -Job $MyJob -Force)
                    }
                  }
                  else
                  {
                    $VerifyObject.Status = "Wrong Workstation Name"
                  }
                }
              }
              else
              {
                $VerifyObject.Status = "Verify Workstation Error"
                [Void](Remove-Job -Job $MyJob -Force)
              }
              # Beak out of Loop, Verify was a Success no need to try other IP Address is any
              Break
            }
          }
        }
        Catch
        {
          # Workstation Not in DNS
          $VerifyObject.Status = "Workstation Not in DNS"
        }
      }
      else
      {
        $VerifyObject.Status = "Invalid Computer Name"
      }
      
      # Calculate Verify Time
      $VerifyObject.Time = ([DateTime]::Now - $StartTime)
      
      # Return Custom Object with Collected Verify Information
      Write-Output -InputObject $VerifyObject
      
      $VerifyObject = $Null
      $Params = $Null
      $MyJob = $Null
      $MyCompData = $Null
      $MyOSData = $Null
      $MyBIOSData = $Null
      $MyChassisData = $Null
      $StartTime = $Null
      
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
    }
    Write-Verbose -Message "Exit Function Test-MyWorkstationSB - Process"
  }
  End
  {
    $MyCompDataJob = $Null
    $MyOSDataJob = $Null
    $MyBIOSDataJob = $Null
    $MyChassisDataJob = $Null
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Verbose -Message "Exit Function Test-MyWorkstationSB"
  }
}
#endregion
