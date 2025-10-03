function Get-VideoDurationTotal {
    <#
    .SYNOPSIS
        Calculates the total duration of video files and returns a detailed object with totals in seconds, minutes, and hours.
    .DESCRIPTION
        This function scans specified folders for common video file types (.mp4, .mov, .mkv, etc.),
        reads the length metadata from each file, and sums their durations.

        It can be filtered to only include files that match a specific name pattern (e.g., "Lab*").

        The function returns a rich custom object containing the total duration broken down into seconds,
        minutes, and hours, along with the number of files processed and the filter that was used.
    .PARAMETER Path
        An array of paths to the folder(s) containing the video files. If not provided, the function
        defaults to the current directory. This parameter accepts pipeline input.
    .PARAMETER Filter
        A string used to filter video files by name. It accepts wildcards (*, ?).
        For example, a filter of "Lab*" would only process video files whose names start with "Lab".
        The default value is "*", which includes all files.
    .EXAMPLE
        Get-VideoDurationTotal -Filter "REC-*.mp4"

        Scans the current directory for video files that start with "REC-" and outputs a detailed duration object.
    .EXAMPLE
        (Get-VideoDurationTotal -Path "C:\Videos").TotalHours

        Calculates the total duration for videos in the specified path and displays only the total hours.
    .EXAMPLE
        Get-ChildItem -Path "C:\Videos" -Directory | Get-VideoDurationTotal -Filter "Meeting*"

        Scans all subdirectories of C:\Videos and calculates the total duration of videos whose names
        start with "Meeting".
    .OUTPUTS
        PSCustomObject
        A custom object with the following properties:
        - TotalFiles: The total number of video files found and processed.
        - TotalSeconds: The grand total duration in seconds.
        - TotalMinutes: The grand total duration expressed in minutes (e.g., 90.5).
        - TotalHours: The grand total duration expressed in hours (e.g., 1.51).
        - FormattedString: A formatted string representing the duration (HH:MM:SS).
        - FilterApplied: The filename filter that was used for the scan.
        - ScannedPaths: A list of the folders that were processed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 0)]
        [string[]]$Path = ".",

        [Parameter(Mandatory = $false)]
        [string]$Filter = "*"
    )

    begin {
        # --- Initialization ---
        $grandTotalSeconds = 0
        $grandTotalFiles = 0
        $processedPaths = [System.Collections.Generic.List[string]]::new()
        $videoExtensions = @(".mp4", ".mov", ".mkv", ".avi", ".wmv", ".flv")

        try {
            $shell = New-Object -ComObject Shell.Application
        }
        catch {
            Write-Error "Failed to create Shell.Application COM object. This feature might not be available. Error: $($_.Exception.Message)"
            return
        }
    }

    process {
        foreach ($folderPath in $Path) {
            $resolvedPath = (Resolve-Path -Path $folderPath).ProviderPath

            if (-not (Test-Path -Path $resolvedPath -PathType Container)) {
                Write-Warning "The path '$resolvedPath' is not a valid folder. Skipping."
                continue
            }

            $processedPaths.Add($resolvedPath)
            Write-Verbose "Scanning folder: $resolvedPath with filter: '$Filter'"

            try {
                $folder = $shell.NameSpace($resolvedPath)
                
                $videoFiles = Get-ChildItem -Path $resolvedPath -File | Where-Object {
                    ($videoExtensions -contains $_.Extension) -and ($_.Name -like $Filter)
                }

                if ($null -eq $videoFiles -or $videoFiles.Count -eq 0) {
                    Write-Verbose "No video files matching the filter found in: $resolvedPath"
                    continue
                }

                $grandTotalFiles += $videoFiles.Count

                foreach ($file in $videoFiles) {
                    $fileItem = $folder.ParseName($file.Name)
                    $durationString = $folder.GetDetailsOf($fileItem, 27)

                    if (-not [string]::IsNullOrWhiteSpace($durationString)) {
                        try {
                            $timeSpan = [TimeSpan]::Parse($durationString)
                            $grandTotalSeconds += $timeSpan.TotalSeconds
                        }
                        catch {
                            Write-Warning "Could not parse duration for file: $($file.Name). Value was: '$durationString'. Skipping."
                        }
                    }
                }
            }
            catch {
                Write-Warning "An error occurred while processing '$resolvedPath'. Error: $($_.Exception.Message)"
            }
        }
    }

    end {
        if ($grandTotalFiles -gt 0) {
            $totalTimeSpan = [TimeSpan]::FromSeconds($grandTotalSeconds)
            $formattedTotal = "{0:D2}:{1:D2}:{2:D2}" -f [int]$totalTimeSpan.TotalHours, $totalTimeSpan.Minutes, $totalTimeSpan.Seconds

            # --- MODIFIED OUTPUT OBJECT ---
            # Added TotalMinutes and TotalHours for easy numeric access.
            [PSCustomObject]@{
                TotalFiles      = $grandTotalFiles
                TotalSeconds    = [math]::Round($grandTotalSeconds, 2)
                TotalMinutes    = [math]::Round($totalTimeSpan.TotalMinutes, 2)
                TotalHours      = [math]::Round($totalTimeSpan.TotalHours, 2)
                FormattedString = $formattedTotal
                FilterApplied   = $Filter
                ScannedPaths    = $processedPaths
            }
        }
        else {
             Write-Host "No video files matching the filter '$Filter' were found in the provided path(s)."
        }

        if ($shell) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
        }
    }
}

