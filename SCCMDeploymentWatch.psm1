<#
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
    [int]$Error
    [int]$Unknown
    [double]$Compliance
    [int]$LowerThreshold
    [int]$UpperThreshold

    # Constructors
    SCCMDeploymentStats () { }

    SCCMDeploymentStats ([string]$CollectionName, [int]$Active, [int]$Success, [int]$Pending, [int]$Error, [int]$Unknown) {
        $this.CollectionName = $CollectionName
        $this.Active         = $Active
        $this.Success        = $Success
        $this.Pending        = $Pending
        $this.Error          = $Error
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
Returns string array of currently watched SCCM deployments.
.DESCRIPTION
Gets the contents of C:\Users\[USERNAME]\Documents\WindowsPowerShell\SCCMWatch.dat indicating all of the Applications
or Software Update Groups you are currently watching.
.EXAMPLE
PS C:\> Get-SCCMApplicationDeploymentWatch

Returns an array of the contents of SCCMWatch.dat which can be fed into some of the other functions.
.NOTES
Just a simple wrapper/helper function.
#>
function Get-SCCMApplicationDeploymentWatchList {
    [CmdletBinding()]
    [OutputType([array])]
    param ()

    [array](Get-Content -Path $env:USERPROFILE\Documents\WindowsPowerShell\SCCMWatch.dat)

}

<#
.SYNOPSIS
Adds a new SCCM Application or Software Update Group for which to monitor deployments.
.DESCRIPTION
Takes the name(s) of an SCCM Application or Software Update Group and adds it to the SCCMWatch.dat file located in your
$env:USERPROFILE\Documents\WindowsPowerShell folder if they are not already present. Get-SCCMApplicationDeploymentWatch
will reference these entries when querying deployments.
.PARAMETER Name
The name or array of names for an Application or a Software Update Group.
.EXAMPLE
PS C:\> Add-SCCMApplicationDeploymentWatch -Application 'WKS - Java 8 Update 131'

.NOTES
Non-valid Applications or Software Update Groups do not cause any actual harm but should eventually be handled
as they can increase query time.
#>
function Add-SCCMApplicationDeploymentWatch {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            Position = 0)]
        [Alias('ApplicationName')]
        [string[]]$Name
    )

    # get the list of currently watched apps
    [System.Collections.ArrayList]$currentApps = @()
    $currentApps.Add((Get-SCCMApplicationDeploymentWatchList -ErrorAction SilentlyContinue))
    Write-Verbose -Message "currently watching $($currentApps.Count)"

    # add the requested new apps
    $currentApps.Add($Name)

    # remove any duplicate entries
    $currentApps = $currentApps | Sort-Object -Unique

    <# TODO: add logic to determine if each app is valid
    foreach ($app in $Application) {
    }
    #>

    Write-Verbose -Message "now watching $($currentApps.Count)"
    Set-Content -Path $env:USERPROFILE\Documents\WindowsPowerShell\SCCMWatch.dat -Value $currentApps

}

<#
.SYNOPSIS
Removes a currently watched SCCM Application or Software Update Group.
.DESCRIPTION
Will remove the name(s) of the provided application(s) or software update group(s) from
C:\Users\[USERNAME\Documents\WindowsPowerShell\SCCMWatch.dat if it is already present.
.PARAMETER Application
A description of the Application parameter.
.EXAMPLE
PS C:\> Remove-SCCMApplicationDeploymentWatch -Application $value1

.NOTES
Additional information about the function.
#>
function Remove-SCCMApplicationDeploymentWatch {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            Position = 0)]
        [Alias('ApplicationName')]
        [string[]]$Name
    )

    # get the list of currently watched apps
    [System.Collections.ArrayList]$currentApps = @()
    $currentApps.Add((Get-SCCMApplicationDeploymentWatchList -ErrorAction Stop))
    Write-Verbose -Message "currently watching $($currentApps.Count)"

    foreach ($application in $Name) {

        Write-Verbose -Message "attempting to remove $application (if present)"
        $currentApps.Remove($application)

    }

    Write-Verbose -Message "now watching $($currentApps.Count)"
    Set-Content -Path $env:USERPROFILE\Documents\WindowsPowerShell\SCCMWatch.dat -Value $currentApps

}

<#
.SYNOPSIS
Shows the status of all currently watched Application and Software Update Group deployments.
.DESCRIPTION
Gets the content from C:\Users\[USERNAME]\Documents\WindowsPowerShell\SCCMWatch.dat and formats the output of
each Get-CMDeployment query into a table.
.PARAMETER OnlyPhasedDeployments
Will return only the deployments for our phased collections. Good for reporting compliance.
.PARAMETER SuppressAllUsers
Suppresses displaying any deployments to the "All Users" collection. Good for reporting compliance.
.PARAMETER Sorted
Sort output alphabetically instead of "as-is" from SCCMWatch.dat
.EXAMPLE
PS C:\> Show-SCCM-ApplicationDeploymentWatch
.NOTES
Additional information about the function.
#>
function Get-SCCMApplicationDeploymentWatch {
    [CmdletBinding()]
    param (
        [string[]]$Application,
        [Alias('OPD')]
        [switch]$OnlyPhasedDeployments = $true,
        [Alias('SAU')]
        [switch]$SuppressAllUsers = $true,
        [switch]$Sorted
    )

    # load the list of applications we are interested in seeing deployments for
    [System.Collections.ArrayList]$currentApps += Get-SCCMApplicationDeploymentWatchList -ErrorAction Stop
    Write-Verbose -Message "currently watching $($currentApps.Count)"

    if ($Sorted) {

        $currentApps = $currentApps | Sort-Object

    }

    $results = [ordered]@{}

    # save current working location so we can run the CM commands on PRI
    Push-Location
    Set-Location -Path PRI:

    foreach ($app in $currentApps) {

        $deployments = Get-CMDeployment -SoftwareName $app

        if (-not $deployments) {

            Write-Verbose -Message "no deployments for $app or no such app"
            continue

        }

        $deploymentStats = @()

        foreach ($deployment in $deployments) {

            if ($SuppressAllUsers -and ($deployment.CollectionName -like "")) {

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

        # create a "Total" entry of all deployments of the current app
        $deploymentStats += New-Object -TypeName SCCMDeploymentStats(
            "Total",
            ($deploymentStats | Measure-Object -Property Active  -Sum).Sum,
            ($deploymentStats | Measure-Object -Property Success -Sum).Sum,
            ($deploymentStats | Measure-Object -Property Pending -Sum).Sum,
            ($deploymentStats | Measure-Object -Property Error   -Sum).Sum,
            ($deploymentStats | Measure-Object -Property Unknown -Sum).Sum
        )

        $results.Add($app, $deploymentStats)

    }

    Pop-Location

    Write-Output -InputObject $results

}

<#
.SYNOPSIS
Neatly displays the output for your list of watched applications and software update groups.
.DESCRIPTION
Blah blah blah
.EXAMPLE
PS C:\> Show-SCCMApplicationDeploymentWatch

.NOTES
The use of Write-Host is acceptable here because this is a 'Show' function whose sole purpose is to neatly display
output in the console window.
#>
function Show-SCCMApplicationDeploymentWatch {
    [CmdletBinding()]
    param (
        [string[]]$Application = (Get-SCCMApplicationDeploymentWatchList)
    )

    $applications = Get-SCCMApplicationDeploymentWatch
    $bufferWidth  = (Get-Host).UI.RawUI.BufferSize.Width
    $dateString   = "[$(Get-Date)]"

    Write-Host ("-" * 10) -ForegroundColor Yellow -BackgroundColor Green -NoNewline
    Write-Host "$dateString" -ForegroundColor Blue -BackgroundColor Green -NoNewline
    Write-Host ("-" * ($bufferWidth - (10 + $dateString.Length))) -ForegroundColor Yellow -BackgroundColor Green

    foreach ($application in $applications.GetEnumerator()) {

        Write-Host ("-" * 5) -ForegroundColor Yellow -BackgroundColor Green -NoNewline
        Write-Host $application.Key -ForegroundColor Blue -BackgroundColor Green -NoNewline
        Write-Host ("-" * ($bufferWidth - (5 + $application.Key.Length))) -ForegroundColor Yellow -BackgroundColor Green

        $application.Value | Format-Table

    }

}

# automatically load format data when module is imported
Update-FormatData -AppendPath (Join-Path $psscriptroot "*.ps1xml")

Export-ModuleMember -Function Get-SCCMApplicationDeploymentWatchList
Export-ModuleMember -Function Add-SCCMApplicationDeploymentWatch
Export-ModuleMember -Function Remove-SCCMApplicationDeploymentWatch
Export-ModuleMember -Function Get-SCCMApplicationDeploymentWatch
Export-ModuleMember -Function Show-SCCMApplicationDeploymentWatch