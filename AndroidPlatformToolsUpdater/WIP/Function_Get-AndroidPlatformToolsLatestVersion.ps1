# Get-AndroidPlatformToolsLatestVersion
function Get-AndroidPlatformToolsLatestVersion {
    <#
        .SYNOPSIS
            Gets Android Platform Tools latest version number by reading from the website.
    #>


    # Input parameters
    [CmdletBinding()]
    [OutputType([System.Version])]
    Param()


    # Begin Function
    Begin {
        $Uri = [string] 'https://developer.android.com/studio/releases/platform-tools'
        Write-Verbose -Message ('Will try to get latest version number from "{0}".' -f $Uri)
    }


    # Process
    Process {
        Try {
            $WebPage = Invoke-WebRequest -Uri $Uri -ErrorAction 'Stop' | Select-Object -ExpandProperty 'Content'
            $Version = [System.Version]$(($WebPage.Split("`r`n") | Where-Object -FilterScript {$_ -like '<h4*'} | Select-Object -First 1).Split('>')[1].Split(' ')[0])
        }
        Catch {
            Write-Verbose -Message ('ERROR: Logic for getting out latest version of Android Platform Tools did not work.')
            $Version = [System.Version]('0.0')
        }
    }


    # End
    End {
        return $Version
    }
}
