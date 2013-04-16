################################################################################
#
# Copyright (c) 2013 Seth Wright <seth@crosse.org>
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#
################################################################################

################################################################################
<#
    .SYNOPSIS
    Transcodes a source video file.

    .DESCRIPTION
    Transcodes a source video file into a destination MP4 or MKV file.

    .INPUTS
    System.String. The names of multiple source files to transcode can be passed via the command line.

    .OUTPUTS
    None.

#>
################################################################################
function Out-M4V {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [Alias("File")]
            [object]
            # The source file to transcode.
            $InputFile,

            [Parameter(Mandatory=$true,
                ParameterSetName="Single")]
            [Parameter(Mandatory=$true,
                ParameterSetName="SingleChaptersDb")]
            [System.IO.FileInfo]
            # The output file.
            $OutputFile,

            [Parameter(Mandatory=$false)]
            [ValidateSet("MP4", "MKV", "Autodetect")]
            [string]
            # The format of the output file.  The default is MP4.
            $OutputFormat = "MP4",

            [Parameter(Mandatory=$false,
                ParameterSetName="Batch")]
            [Parameter(Mandatory=$false,
                ParameterSetName="BatchChaptersDb")]
            [System.IO.DirectoryInfo]
            # The output path.  The resulting file name will the the same as the input file, with an extension that depends on the output format.
            $OutputPath = (Get-Location).Path,

            [Parameter(Mandatory=$false)]
            [ValidateSet("480p", "720p", "1080p")]
            [string]
            # The maximum video resolution to support. Supported values for this parameter are 480p, 720p, and 1080p.
            $MaxVideoFormat,

            [Parameter(Mandatory=$false)]
            [ValidateRange(1, 51)]
            [int]
            # Set the video quality, from 1 to 51.  The default is 18.
            $VideoQuality = 18,

            [Parameter(Mandatory=$false)]
            [switch]
            # Indicates whether to optimize for HTTP streaming.  The default is true.  (This only applies to MP4-encoded output files.)
            $OptimizeForHttpStreaming = $true,

            [Parameter(Mandatory=$false)]
            [switch]
            # Indicates whether to always include a stereo, Dolby Pro Logic II version of the main audio track.  The default is true.
            $AlwaysIncludeStereoTrack = $true,

            [Parameter(Mandatory=$false)]
            [switch]
            # Indicates whether to include a Dolby Digital 5.1 (AC3) version of the main audio track when it is not in AC3 format (for instance, when the main audio track is a DTS track).  The default is true.
            $AlwaysIncludeAC3Track = $true,

            [Parameter(Mandatory=$true,
            [Parameter(Mandatory=$false)]
            [switch]
            # Adds an extra pass that scans subtitles matching the language of the first audio or the language selected by the -NativeLanguage parameter. The one that's only used 10 percent of the time or less is selected. This should locate subtitles for short foreign language segments. The default is true.
            $SubtitleScan = $true,

            [Parameter(Mandatory=$false)]
            [string]
            # Specifies the native language preference for subtitles.  When the default audio track does not match this language then select the first subtitle that does.  The format for this parameter is the desired language's ISO639-2 code. The default is "eng" (English).
            $NativeLanguage = "eng",

            [Parameter(Mandatory=$false)]
            [switch]
            # Indicates whether to overwrite an output file that already exists.
            $Force,

            [Parameter(Mandatory=$false)]
            [string]
            # The path to HandbrakeCLI.exe.  By default this is "C:\Program Files\Handbrake\HandbrakeCLI.exe".
            $HandbrakeCLIPath = "C:\Program Files\Handbrake\HandbrakeCLI.exe",

            [Parameter(Mandatory=$false)]
            [string]
            # The path to the MediaInfo.exe CLI program (NOT the GUI application, which unfortunately has the same name).  The default is to use the binaries in the same directory as this cmdlet's module manifest.
            $MediaInfoCLIPath = (Join-Path $PSScriptRoot "MediaInfo.exe"),

            [Parameter(Mandatory=$false)]
            [switch]
            # Shows what would happen if the cmdlet runs. The cmdlet is not run.
            $WhatIf
          )

    begin {
        Resolve-Path $HandbrakeCLIPath -ErrorAction Stop | Out-Null
        Write-Verbose "Found HandbrakeCLI: $HandbrakeCLIPath"

        Resolve-Path $MediaInfoCLIPath -ErrorAction Stop | Out-Null
        Write-Verbose "Found MediaInfo: $MediaInfoCLIPath"

        $OutputPath = (Resolve-Path $OutputPath -ErrorAction Stop).Path
        Write-Verbose "Output Path: $OutputPath"

        switch ($OutputFormat) {
            "MP4" { $format = "--format mp4" }
            "MKV" { $format = "--format mkv" }
            "Audodetect" { }
        }

        $handbrakeOptions = @(
                # Set output format.
                $format
                # Add chapter markers
                '--markers',
                # Use 64-bit mp4 files that can hold more than 4GB.
                "--large-file"
                # Set video library encoder
                "--encoder x264"
                # advanced encoder options in the same style as mencoder
                "--encopts `"b-adapt=2`""
                # Set video quality
                "--quality $VideoQuality"
                # Set video framerate
                "--rate 29.97"
                # Select peak-limited frame rate control.
                "--pfr"
                # Set audio codec to use when it is not possible to copy an
                # audio track without re-encoding.
                "--audio-fallback ffac3"
                # Store pixel aspect ratio with specified width
                "--loose-anamorphic"
                # Set the number you want the scaled pixel dimensions to divide
                # cleanly by.
                "--modulus 2"
                # Selectively deinterlaces when it detects combing
                "--decomb"
                )

        if ($OptimizeForHttpStreaming) {
            $handbrakeOptions += "--optimize"
        }

        if ($SubtitleScan) {
            $handbrakeOptions += "--subtitle scan"
        }

        if ($NativeLanguage) {
            $handbrakeOptions += "--native-language eng"
        }

        if ($Verbose) {
            $handbrakeOptions += "--verbose"
        } else {
            $handbrakeOptions += "--verbose 0"
        }

    }
    process {
        try {
            if ($InputFile -isnot [System.IO.FileInfo]) {
                $inter = Resolve-Path $InputFile
                $InputFile = [System.IO.FileInfo]$inter.Path
            }
            Write-Verbose "Input File: $InputFile"

            if ([String]::IsNullOrEmpty($OutputFile)) {
                switch ($OutputFormat) {
                    "MP4" {
                        $extension = ".m4v"
                    }
                    "MKV" {
                        $extension = ".mkv"
                    }
                    "Autodetect" {
                        Write-Error "Autodetect cannot be used with -OutputPath.  Use -OutputFile instead."
                        return
                    }
                }
                $outFile = Join-Path $OutputPath $InputFile.Name.Replace($InputFile.Extension, $extension)
            } else {
                $outFile = $OutputFile
            }

            if ((Test-Path $outFile) -eq $true -and
                    $Force -eq $false -and
                    $WhatIf -eq $false) {
                Write-Error "Output file already exists! ($outFile)"
                return
            }

            Write-Verbose "Output File: $outFile"

            $command = "& '$MediaInfoCLIPath' --Output=XML `"$($inputFile.FullName)`""
            [xml]$info = Invoke-Expression "$command"

            $audio = @($info.MediaInfo.File.Track | ? { $_.type -match "Audio" })
            if ($audio -eq $null) {
                Write-Error "Error getting audio track information from source."
                return
            }

            Write-Verbose "Processing source audio tracks"

            $audioTracks = @()
            for ($trackNumber = 1; $trackNumber -le $audio.Count; $trackNumber++) {
                Write-Verbose "Evaluating track $trackNumber of $($audio.Count)"
                $audioTrack = $audio[$trackNumber - 1]

                $trackTitle = $audioTrack.Title
                $trackLang = $audioTrack.Language

                Write-Verbose "Audio Track ${trackNumber}: Title: $trackTitle; Language: $trackLang"

                if ($audioTrack.Default -eq "Yes") {
                    $audioTitle = $audioTrack.Title
                    Write-Verbose "Track $trackNumber is the default audio track."

                    if ($AlwaysIncludeStereoTrack) {
                        # Transcode the main audio track into a stereo track.
                        $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy:aac -Mixdown "ProLogicII" -Name "Dolby Pro Logic II"
                    }

                    if ($audioTrack.Format -match "AC-3|MA" -or $AlwaysIncludeAC3Track) {
                        # Copy or Transcode the main audio track as an AC-3 track
                        # if it is:
                        # 1) an AC-3 track already;
                        # 2) a DTS-HD Master Audio track (because most things can't decode this yet); or
                        # 3) the -AlwaysIncludeAC3Track option is set (the default)
                        $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy:ac3 -Mixdown "6ch" -Name "Dolby Digital 5.1"
                    }

                    if ($audioTrack.Format -match "DTS") {
                        if ($audioTrack.Format_profile -match '^MA') {
                            # If a DTS-HD MA track exists, pass it through.
                            $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy:dtshd -Name "DTS-HD Master Audio"
                        } else {
                            # If a regular DTS track exists, pass it through.
                            $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy:dts -Name "DTS"
                        }
                    }
                } else {
                    Write-Verbose "Track $trackNumber is a secondary audio track."
                    # This is a secondary track; copy it as-is if possible, or
                    # else transcode it to AC3.
                    $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy -Name "$trackTitle ($trackLang)"
                }
            }

            #$audioTracks = $audioTracks | Sort-Object Track, Name, Encoder

            $trackNumbers = ($audioTracks | % { $_.Track }) -join ','
            $trackEncodings = ($audioTracks | % { $_.Encoder }) -join ','
            $trackNames = ($audioTracks | % { $_.Name }) -join ','
            $mixdown = ($audioTracks | % { $_.Mixdown}) -join ','

            $audioOptions = "--audio `"$trackNumbers`" --aencoder `"$trackEncodings`" --mixdown `"$mixdown`" --aname `"$trackNames`""

            if ([String]::IsNullOrEmpty($MaxVideoFormat)) {
                $video = $info.MediaInfo.File.Track | ? { $_.type -match "Video" }
                if ($video -eq $null) {
                    Write-Error "Error getting video track information from source."
                    return
                }
            } else {
                switch ($MaxVideoFormat) {
                    '480p' { $videoOptions = '--maxWidth 480' }
                    '720p' { $videoOptions = '--maxWidth 1280' }
                    '1080p' { $videoOptions = '--maxWidth 1920' }
                }
            }

            }
            $fileOptions = "--input `"$($inputFile.FullName)`" --output `"$outFile`""

            $command = "& '$HandbrakeCLIPath' $fileOptions "
            $command += $handbrakeOptions -join " "
            $command += " "
            $command += $audioOptions -join " "
            $command += " "
            $command += $videoOptions -join " "
            Write-Verbose $command

            if ($WhatIf -eq $false) {
                Invoke-Expression "$command"
            }
        }
        catch
        {
            throw
        }
    }
    end {
    }
}

function New-AudioTrack {
    param (
            [Parameter(Mandatory=$true)]
            [int]
            $Track,

            [Parameter(Mandatory=$true)]
            [string]
            $Name,

            [Parameter(Mandatory=$true)]
            [ValidateSet(
                "faac",
                "ffaac",
                "copy:aac",
                "ffac3",
                "copy:ac3",
                "copy:dts",
                "copy:dtshd",
                "lame",
                "copy:mp3",
                "vorbis",
                "ffflac",
                "copy"
                )]
            [string]
            $Encoder,

            [Parameter(Mandatory=$false)]
            [ValidateSet(
                "Auto",
                "Mono",
                "Stereo",
                "ProLogicI",
                "ProLogicII",
                "6ch"
                )]
            [string]
            $Mixdown = "Auto"
          )


    return New-Object PSObject -Property @{
        Track = $Track
        Name = $Name
        Encoder = $Encoder
        Mixdown = $Mixdown
    }
}


