Get-ADObject -SearchBase "OU=L30_PCN,OU=Assets,DC=wmgpcn,DC=local" -LDAPFilter "(objectClass=computer)" | 
Where-Object { $_.Name -notlike "PCNVS*" -and $_.Name -notlike "DEVVS*" -and $_.Name -notlike "PCNVC*" <#-and $_.Name -notlike "PCNLAP*" #> } | 
Select-Object -ExpandProperty Name | Set-Variable -Name computers

$daysToCheck = -30
$checkDate = (Get-Date).AddDays($daysToCheck)

$results = New-Object System.Collections.Generic.List[System.Object]
$errors = New-Object System.Collections.Generic.List[System.Object]

Write-Host "`nRunning... Please wait..."

foreach ($computer in $computers) { 
    Try {
        Test-Connection $computer -Count 1 -ErrorAction Stop > $null
        $UnexpectedRebootCount = (Get-WinEvent -ComputerName $computer -FilterHashtable @{ LogName='System'; StartTime=$checkDate; Id='6008' } -ErrorAction Stop | Measure-Object).Count
        if ($unexpectedRebootCount -gt 0) {
            $results.add([PSCustomObject]@{'Hostname'=$computer ; 
                                           'Unexpected Reboot Count (Event ID 6008)' = $UnexpectedRebootCount
                                          }
                        )
        }
    }
    Catch { 
        if ($_.Exception.Message -notlike "*No events were found that match the specified selection criteria*") {
            $errors.add([PSCustomObject]@{'Hostname'=$computer ; 
                                          'Error' = $_.Exception.Message
                                         }
                        )
                }
          }
}
 
#Export the data to file.
$outputFile = ".\UnexpectedShutdownCheck Output-$(Get-Date -Format MMddyyyy_HHmmss).csv"
$results | Sort-Object 'Unexpected Reboot Count (Event ID 6008)', Hostname | ConvertTo-CSV -NoTypeInformation | Add-Content -Path $outputFile
if($results.count -gt 0) { $results | Sort-Object WUServer, Hostname }
else { Write-Host "`nNo nodes found with unexpected shutdown (Event ID 6008) within the last $daysToCheck days." }

$outputString = "`r`n** Errors **"
Add-Content -Path $outputFile -Value $outputString
$errors | Select-Object | ConvertTo-CSV -NoTypeInformation | Add-Content -Path $outputFile
$errors | Sort-Object Hostname 