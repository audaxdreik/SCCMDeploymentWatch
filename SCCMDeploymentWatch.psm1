﻿<#
.NOTES
=======================================================================================================================
    Created on:    5/22/2017 10:06 AM
    Created by:    audax dreik
    Filename:      SCCMDeploymentWatch.psm1
=======================================================================================================================
.DESCRIPTION
Easy monitoring of ongoing SCCM deployments.
#>

class SCCMDeploymentStats {

    [string]$CollectionName
    [int]$Active
    [int]$Success
    [int]$Pending
    [int]$Errors
    [int]$Unknown
    [double]$Compliance
    [int]$LowerThreshold
    [int]$UpperThreshold

    # Constructors
    SCCMDeploymentStats () { }

    SCCMDeploymentStats ([string]$CollectionName, [int]$Active, [int]$Success, [int]$Pending, [int]$Errors, [int]$Unknown) {
        $this.CollectionName = $CollectionName
        $this.Active         = $Active
        $this.Success        = $Success
        $this.Pending        = $Pending
        $this.Error          = $Errors
        $this.Unknown        = $Unknown
        $this.Compliance     = [System.Math]::Round(($Success / $Active) * 100, 2)

        $this.LowerThreshold = 85
        $this.UpperThreshold = 95
    }

    # for printing pretty like to format file
    [string] showCompliance() {

        switch ($this.Compliance){
            { $_ -ge $this.UpperThreshold } {
                # if greater or equal to the upper threshold, yay! highlight in green ANSI
                return "$([char](0x1B))[48;5;2m$($this.Compliance.ToString("#.00"))$([char](0x1B))[0m"
            }
            { $_ -lt $this.LowerThreshold } {
                # if less than lower threshold, boo! highlight in red ANSI
                return "$([char](0x1B))[48;5;9m$($this.Compliance.ToString("#.00"))$([char](0x1B))[0m"
            }
        }

        # default cast to string and return
        return $this.Compliance.ToString("#.00")

    }

}

<#
.SYNOPSIS
Returns the string or array of strings for currently watched SCCM deployments.
.DESCRIPTION
Gets the contents of C:\Users\[USERNAME]\Documents\WindowsPowerShell\SCCMWatch.dat indicating all of the Applications
or Software Update Groups you are currently watching.
.EXAMPLE
PS C:\> Get-SCCMApplicationDeploymentWatch

Returns a string or array of strings generated from SCCMWatch.dat which can be fed into some of the other functions.
.NOTES
Just a simple wrapper/helper function.
#>
function Get-SCCMDeploymentWatchList {
    [CmdletBinding()]
    param ()

    Get-Content -Path $env:USERPROFILE\Documents\WindowsPowerShell\SCCMWatch.dat

}

<#
.SYNOPSIS
Adds a new SCCM Application or Software Update Group for which to monitor deployments.
.DESCRIPTION
Takes the name(s) of an SCCM Application or Software Update Group and adds it to the SCCMWatch.dat file located in your
$env:USERPROFILE\Documents\WindowsPowerShell folder if they are not already present. Get-SCCMDeploymentWatch will
reference these entries when querying deployments.
.PARAMETER Name
The name or array of names for an Application or a Software Update Group.
.EXAMPLE
PS C:\> Add-SCCMApplicationDeploymentWatch -Name 'Java 8 Update 131'

Adds the 'Java 8 Update 131' Application to SCCMWatch.dat if it is not already present.
.NOTES
Non-valid Applications or Software Update Groups do not cause any actual harm but should eventually be handled as they
can increase query time.
#>
function Add-SCCMDeploymentWatchList {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline,
            Position = 0,
            HelpMessage = 'application or software update group')]
        [Alias('DeploymentName')]
        [string[]]$Name
    )

    begin {

        # get the list of currently watched apps
        [System.Collections.ArrayList]$currentApps = @()
        Get-SCCMDeploymentWatchList -ErrorAction SilentlyContinue |
            ForEach-Object -Process { $currentApps.Add($_) | Out-Null}
        Write-Verbose -Message "currently watching $($currentApps.Count)"

    }

    process {

        # add the requested new apps (toss the index [int] returned)
        $Name | ForEach-Object -Process { $currentApps.Add($_) | Out-Null }

        <# TODO: add logic to determine if each app is valid
        foreach ($app in $Application) {
        }
        #>

    }

    end {

        # remove duplicate entries, ensure array returned on single entry to avoid casting error
        $currentApps = @($currentApps | Sort-Object -Unique)

        Write-Verbose -Message "now watching $($currentApps.Count)"
        Set-Content -Path $env:USERPROFILE\Documents\WindowsPowerShell\SCCMWatch.dat -Value $currentApps

    }

}

<#
.SYNOPSIS
Removes a currently watched SCCM Application or Software Update Group.
.DESCRIPTION
Will remove the name(s) of the provided Application or Software Update Group from
C:\Users\[USERNAME\Documents\WindowsPowerShell\SCCMWatch.dat if it is already present.
.PARAMETER Name
The name or array of names for an Application or a Software Update Group.
.EXAMPLE
PS C:\> Remove-SCCMApplicationDeploymentWatch -Name 'Adobe Flash','Adobe Reader DC'

Removes 'Adobe Flash' and 'Adobe Reader DC' from SCCMWatch.dat if they are present.
.NOTES
Attempting to remove an Application or Software Update Group that isn't present in the file will not result in errors.
#>
function Remove-SCCMDeploymentWatchList {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline,
            Position = 0,
            HelpMessage = 'application or software update group')]
        [Alias('DeploymentName')]
        [string[]]$Name
    )

    begin {

        # get the list of currently watched apps
        [System.Collections.ArrayList]$currentApps = @()
        Get-SCCMDeploymentWatchList -ErrorAction Stop |
            ForEach-Object -Process { $currentApps.Add($_) | Out-Null}
        Write-Verbose -Message "currently watching $($currentApps.Count)"

    }

    process {

        # remove each specified app, no error thrown if not present so don't bother checking
        $Name | ForEach-Object -Process { $currentApps.Remove($_) }

    }

    end {

        Write-Verbose -Message "now watching $($currentApps.Count)"
        Set-Content -Path $env:USERPROFILE\Documents\WindowsPowerShell\SCCMWatch.dat -Value $currentApps

    }

}

<#
.SYNOPSIS
Shows the status of all currently watched Application and Software Update Group deployments.
.DESCRIPTION
Gets the content from C:\Users\[USERNAME]\Documents\WindowsPowerShell\SCCMWatch.dat and formats the output of
each Get-CMDeployment query into a table.
.PARAMETER OnlyPhasedDeployments
Will return only the deployments for our phased collections. Good for reporting compliance.
.PARAMETER ShowAllUsers
Includes Available deployments to the 'All Users' collection. Usually omitted for better compliance reporting on
Required deployments.
.EXAMPLE
PS C:\> Show-SCCM-ApplicationDeploymentWatch
.NOTES
Additional information about the function.
#>
function Get-SCCMDeploymentWatch {
    [CmdletBinding()]
    [OutputType([SCCMDeploymentStats[]])]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline,
            Position = 0,
            HelpMessage = 'application or software update group')]
        [Alias('DeploymentName')]
        [string[]]$Name,
        [Alias('OPD')]
        [switch]$OnlyPhasedDeployments = $true,
        [Alias('SAU')]
        [switch]$ShowAllUsers
    )

    begin {

        # save current working location so we can run the CM commands on PRI
        Push-Location
        Set-Location -Path PRI:

    }

    process {

        foreach ($applicationName in $Name) {

            $deployments = Get-CMDeployment -SoftwareName $applicationName

            if (-not $deployments) {

                Write-Verbose -Message "no deployments for $applicationName or no such app"
                break

            }

            $deploymentStats = @()

            foreach ($deployment in $deployments) {

                if (-not $ShowAllUsers -and ($deployment.CollectionName -like "")) {

                    Write-Verbose -Message "suppressing 'All Users' collection"
                    continue

                }

                if ($OnlyPhasedDeployments -and ($deployment.CollectionName -notmatch "Phase \d")) {

                    Write-Verbose -Message "suppressing non-phased deployment, $($deployment.CollectionName)"
                    continue

                }

                $deploymentStats += New-Object -TypeName SCCMDeploymentStats(
                    $deployment.CollectionName,
                    $deployment.NumberTargeted,
                    $deployment.NumberSuccess,
                    $deployment.NumberInProgress,
                    $deployment.NumberErrors,
                    $deployment.NumberUnknown
                )

            }

        }

    }

    end {

        # create a 'Total' entry of all deployments of the current app
        $deploymentStats += New-Object -TypeName SCCMDeploymentStats(
            'Total',
            ($deploymentStats | Measure-Object -Property Active  -Sum).Sum,
            ($deploymentStats | Measure-Object -Property Success -Sum).Sum,
            ($deploymentStats | Measure-Object -Property Pending -Sum).Sum,
            ($deploymentStats | Measure-Object -Property Errors  -Sum).Sum,
            ($deploymentStats | Measure-Object -Property Unknown -Sum).Sum
        )

        Pop-Location

        Write-Output -InputObject $deploymentStats

    }

}

<#
.SYNOPSIS
Neatly displays the output for your list of watched Applications and Software Update Groups.
.DESCRIPTION
Outputs a header with timestamp to show when the command was last run, along with the name of each Application or
Software Update Group above each set of results to make the current status of each deployment clearly identifiable.
.PARAMETER Name
The name or array of names for an Application or a Software Update Group.
.EXAMPLE
PS C:\> Show-SCCMApplicationDeploymentWatch

Automatically passes the watched Applications or Software Update Groups from the SCCMWatch.dat file into the
Get-SCCMDeploymentWatch commandlet.
.EXAMPLE
PS C:\> Show-SCCMApplicationDeploymentWatch -Name 'Cisco AnyConnect VPN'

Shows the status of all current deployments for the 'Cisco AnyConnect VPN' application.
.NOTES
The use of Write-Host is acceptable here because this is a 'Show' function whose sole purpose is to neatly display
output in the console window.
#>
function Show-SCCMDeploymentWatch {
    [CmdletBinding()]
    param (
        [Parameter(
            Position = 0,
            ValueFromPipeline)]
        [Alias('DeploymentName')]
        [string[]]$Name = (Get-SCCMDeploymentWatchList)
    )

    begin {

        # show a fancy banner to indicate the time these deployment stats were gathered
        $bufferWidth = (Get-Host).UI.RawUI.BufferSize.Width
        $dateString = "[$(Get-Date)]"

        Write-Host ("-" * 10) -ForegroundColor Yellow -BackgroundColor Green -NoNewline
        Write-Host "$dateString" -ForegroundColor Blue -BackgroundColor Green -NoNewline
        Write-Host ("-" * ($bufferWidth - (10 + $dateString.Length))) -ForegroundColor Yellow -BackgroundColor Green

    }

    process {

        foreach ($deployment in $Name) {

            $deploymentStats = Get-SCCMDeploymentWatch -Name $deployment

            Write-Host ("-" * 5) -ForegroundColor Yellow -BackgroundColor Green -NoNewline
            Write-Host $deployment -ForegroundColor Blue -BackgroundColor Green -NoNewline
            Write-Host ("-" * ($bufferWidth - (5 + $deployment.Length))) -ForegroundColor Yellow -BackgroundColor Green

            $deploymentStats | Format-Table

        }

    }

}

# automatically load format data when module is imported
Update-FormatData -AppendPath (Join-Path $PSScriptRoot '*.ps1xml')

Export-ModuleMember -Function Get-SCCMDeploymentWatchList
Export-ModuleMember -Function Add-SCCMDeploymentWatchList
Export-ModuleMember -Function Remove-SCCMDeploymentWatchList
Export-ModuleMember -Function Get-SCCMDeploymentWatch
Export-ModuleMember -Function Show-SCCMDeploymentWatch