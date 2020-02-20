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

            [Parameter(Mandatory=$false)]
            [switch]
            # Indicates whether to include a Dolby Digital 5.1 (AC3) version of High-Definition audio tracks (such as DTS-HD Master Audio tracks) when it is not in AC3 format.  The default is true.
            $IncludeAC3ForHDAudio = $true,

            [Parameter(Mandatory=$false)]
            [int[]]
            # An array of audio tracks to ignore and not add to the output file.  The audio tracks are numbered starting from one and are in the same order as MediaInfo reports.
            $IgnoreAudioTracks,

            [Parameter(Mandatory=$false,
                ParameterSetName="ChaptersDb")]
            [Parameter(Mandatory=$false,
                ParameterSetName="SingleChaptersDb")]
            [Parameter(Mandatory=$false,
                ParameterSetName="BatchChaptersDb")]
            [switch]
            # Indicates whether to attempt to look up chapter names on ChaptersDb.org.
            $LookupChapterNames,

            [Parameter(Mandatory=$true,
                ParameterSetName="ChaptersDb")]
            [Parameter(Mandatory=$true,
                ParameterSetName="SingleChaptersDb")]
            [Parameter(Mandatory=$true,
                ParameterSetName="BatchChaptersDb")]
            [ValidateNotNullOrEmpty()]
            [string]
            # Your API Key for the ChaptersDb.org website.
            $ChaptersDbApiKey,

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
            [switch]
            # Shows what would happen if the cmdlet runs. The cmdlet is not run.
            $WhatIf
          )

    begin {
        if ($IsWindows) {
            $HandbrakeCLIPath = "HandbrakeCLI.exe"
        } elseif ($IsMacOS) {
            $HandbrakeCLIPath = "HandbrakeCLI"
        } else {
            Write-Error "I don't know where to find HandbrakeCLI!"
        }

        $HandbrakeCLIPath = (Get-Command $HandbrakeCLIPath -ErrorAction Stop).Source
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

            $audio = @($info.MediaInfo.File.Track | Where-Object { $_.type -match "Audio" })
            if ($audio -eq $null) {
                Write-Error "Error getting audio track information from source."
                return
            }

            Write-Verbose "Processing source audio tracks"

            $audioTracks = @()
            for ($trackNumber = 1; $trackNumber -le $audio.Count; $trackNumber++) {
                if (($trackNumber) -in $IgnoreAudioTracks) {
                    Write-Verbose "Ignoring audio track $trackNumber"
                    continue
                }

                Write-Verbose "Evaluating track $trackNumber of $($audio.Count)"
                $audioTrack = $audio[$trackNumber - 1]

                $isDefaultTrack = ($audioTrack.Default -eq "Yes")
                $trackTitle = $audioTrack.Title
                $trackLang = $audioTrack.Language
                $trackFormat = $audioTrack.Format

                Write-Verbose "Audio Track ${trackNumber}: Title: $trackTitle; Format: $trackFormat; Format Profile: $($audioTrack.Format_profile); Language: $trackLang"

                if ($isDefaultTrack) {
                    Write-Verbose "`tTrack $trackNumber is the default audio track."

                    if ([String]::IsNullOrEmpty($trackTitle)) {
                        $trackTitle = "Main"
                    }

                    if ($AlwaysIncludeStereoTrack) {
                        Write-Verbose "`tAdding Pro Logic II version of default audio track."
                        # Transcode the main audio track into a stereo track.
                        $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy:aac -Mixdown "ProLogicII" -Name "Main Stereo (Dolby Pro Logic II / $trackLang)"
                    }
                } else {
                    Write-Verbose "`tTrack $trackNumber is a secondary audio track."
                }

                # Pass-through any DTS tracks.
                if ($trackFormat -match "^DTS") {
                    # Pass through any DTS tracks.
                    if ($audioTrack.Format_profile -match '^MA') {
                        $format = "DTS-HD Master Audio"
                    } elseif ($audioTrack.Format_profile -match '^ES') {
                        $format = "DTS-ES"
                    } else {
                        $format = "DTS"
                    }
                    Write-Verbose "`tPassing through $format track."
                    $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy -Name "$trackTitle ($format / $trackLang)"
                }

                # Copy or Transcode the main audio track as an AC-3 track
                # if it is:
                # 1) an AC-3 track already;
                # 2) the -AlwaysIncludeAC3Track option is set (the default) and this is the main audio track; or
                # 3) a DTS-HD Master Audio track (because most things can't decode this yet)
                if ($trackFormat -match "^AC-3" -or
                        ($AlwaysIncludeAC3Track -and $isDefaultTrack) -or
                        ($IncludeAC3ForHDAudio -and $audioTrack.Format_profile -match '^MA')) {

                    if ($audioTrack.Format_profile -match '^MA') {
                        Write-Verbose "`tAdding AC-3 track for compatibility."
                    } else {
                        Write-Verbose "`tAdding AC-3 track."
                    }

                    # Rename the audio track if it mentions lossless (because AC-3 is not) or
                    # 3/2+1.
                    if ($trackTitle -eq 'Lossless' -or $trackTitle -eq '3/2+1') {
                        $title = "Dolby Digital 5.1"
                    } else {
                        $title = $trackTitle
                    }

                    $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy:ac3 -Name "$title (AC-3 / $trackLang)"
                }

                # Finally, if the track is not an AC-3 track or DTS track, try to pass it through
                # or transcode it to AC-3.
                if ($trackFormat -notmatch '^AC-3|^DTS') {
                    Write-Verbose "`tAdding $trackFormat track."

                    # Transcode or pass through AC-3.
                    $audioTracks += New-AudioTrack -Track $trackNumber -Encoder copy -Name "$trackTitle ($trackFormat / $trackLang)"
                }
            }

            #$audioTracks = $audioTracks | Sort-Object Track, Name, Encoder

            $trackNumbers = ($audioTracks | Foreach-Object { $_.Track }) -join ','
            $trackEncodings = ($audioTracks | Foreach-Object { $_.Encoder }) -join ','
            $trackNames = ($audioTracks | Foreach-Object { $_.Name }) -join ','
            $mixdown = ($audioTracks | Foreach-Object { $_.Mixdown}) -join ','

            $audioOptions = "--audio `"$trackNumbers`" --aencoder `"$trackEncodings`" --mixdown `"$mixdown`" --aname `"$trackNames`""

            if ([String]::IsNullOrEmpty($MaxVideoFormat)) {
                $video = $info.MediaInfo.File.Track | Where-Object { $_.type -match "Video" }
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

            if ($LookupChapterNames) {
                $chapterCount = 0
                $menu = $info.Mediainfo.File.track | Where-Object { $_.type -eq 'Menu' }
                $enumerator = $menu.GetEnumerator()
                while ($enumerator.MoveNext() -eq $true) {
                    $chapterCount++
                }
                if ([String]::IsNullOrEmpty($video.Title) -eq $true) {
                    Write-Warning "Video track has no title!  Cannot look up chapter names."
                } else {
                    Write-Verbose "Contacting ChaptersDb.org (Title: $($video.Title), ChapterCount: $chapterCount)"
                    $chapterInfo = Get-ChapterInformation -Title $video.Title -ChapterCount $chapterCount -UseChaptersDb -ChaptersDbApiKey $ChaptersDbApiKey -BestResult
                    if ($chapterInfo -eq $null) {
                        Write-Warning "No chapter information could be retrieved from ChaptersDb.org."
                        $handbrakeOptions += "--markers"
                    } else {
                        $tempFile = [IO.Path]::GetTempFileName()
                        Write-Verbose "Chapters:"
                        for ($i = 1; $i -le $chapterInfo.Chapters.Count; $i++) {
                            Write-Verbose "${i}: $($chapterInfo.Chapters[$i - 1].Title)"
                            Write-Output "$i,$($chapterInfo.Chapters[$i - 1].Title.Replace(',', '\,'))" | `
                                Out-File -Append -FilePath $tempFile -Encoding ASCII
                        }

                        $handbrakeOptions += "--markers=`"$tempFile`""
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
                $startTime = Get-Date
                Invoke-Expression "$command"
                $elapsed = [Math]::Round(((Get-Date) - $starttime).TotalMinutes, 2)
                Write-Verbose "Encoding took $elapsed total minutes."
            }
        }
        catch
        {
            throw
        }
        finally {
            if ($tempFile -ne $null -and (Test-Path $tempFile)) {
                Write-Verbose "Removing chapters temp file $tempFile"
                Remove-Item $tempFile
            }
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

    Write-Verbose "`t--> Creating new track `"$Name`" from source track $Track"

    return New-Object PSObject -Property @{
        Track = $Track
        Name = $Name
        Encoder = $Encoder
        Mixdown = $Mixdown
    }
}

function Get-MediaInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [Alias("File")]
        [object]
        # The input file.
        $InputFile
    )

    begin {
        if ($IsWindows) {
            $MediaInfoCLIPath = "MediaInfo.exe"
        } elseif ($IsMacOS) {
            $MediaInfoCLIPath = "mediainfo"
        } else {
            Write-Error "I don't know where to find MediaInfo!"
        }

        $MediaInfoCLIPath = (Get-Command $MediaInfoCLIPath -ErrorAction Stop).Source
        Write-Verbose "Found MediaInfo: $MediaInfoCLIPath"
    }

    process {
        if ($InputFile -isnot [System.IO.FileInfo]) {
            $inter = Resolve-Path $InputFile
            $InputFile = [System.IO.FileInfo]$inter.Path
        }
        Write-Verbose "Input File: $InputFile"

        $command = "& '$MediaInfoCLIPath' --Output=JSON `"$($InputFile.FullName)`""
        $info = Invoke-Expression "$command" | ConvertFrom-Json
        return $info.media.track
    }
}

function Get-M4VTrack {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [Alias("File")]
        [object]
        # The source MKV file.
        $InputFile,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Menu", "Video", "Audio", "Text", "All")]
        [String]
        $TrackType = "All"
    )

    if ($InputFile -isnot [System.IO.FileInfo]) {
        $inter = Resolve-Path $InputFile
        $InputFile = [System.IO.FileInfo]$inter.Path
    }

    $tracks = @(Get-MediaInfo -InputFile $InputFile)
    if ($TrackType -ne "All") {
        $tracks = $tracks | Where-Object { $_."@type" -eq $TrackType }
    }

    return $tracks
}

function Export-SRTSubtitle {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [Alias("File")]
        [object]
        # The source MKV file from which to extract subtitles.
        $InputFile,

        [Parameter(Mandatory=$false,
                   ParameterSetName="Batch")]
        [System.IO.DirectoryInfo]
        # The output path.  The resulting file name will the the same as the input file, with an extension that depends on the output format.
        $OutputPath = (Get-Location).Path,

        [Parameter(Mandatory=$false)]
        [switch]
        # Extract all subtitles. Normally only subtitle tracks marked as "default" and "forced" are extracted.
        $All,

        [Parameter(Mandatory=$false)]
        [switch]
        # Indicates whether to overwrite an output file that already exists.
        $Force,

        [Parameter(Mandatory=$false)]
        [switch]
        # Shows what would happen if the cmdlet runs. The cmdlet is not run.
        $WhatIf
    )

    begin {
        if ($IsWindows) {
            $mkvextract = "mkvextract.exe"
        } else {
            $mkvextract = "mkvextract"
        }

        $mkvextract = (Get-Command $mkvextract -ErrorAction Stop).Source
        Write-Verbose "Found mkvextract: $mkvextract"

        $OutputPath = (Resolve-Path $OutputPath -ErrorAction Stop).Path
        Write-Verbose "Output Path: $OutputPath"
    }

    process {
        if ($InputFile -isnot [System.IO.FileInfo]) {
            $inter = Resolve-Path $InputFile
            $InputFile = [System.IO.FileInfo]$inter.Path
        }

        $tracks = Get-M4VTrack -InputFile $InputFile | ? { $_."@type" -ne "General" }
        for ($i = 0; $i -lt $tracks.Count; $i++) {
            $s = $tracks[$i]

            if ($s.CodecID -notmatch '(S_TEXT/(UTF8|ASCII)|PGS)') {
                Write-Information ("Skipping non-text track {0} ({1}, {2})" -f $i, $s."@type", $s.CodecID)
                continue
            }

            if ($s.Forced -eq "No" -and $s.Default -eq "No" -and $All -eq $false) {
                Write-Information ("Skipping non-default, non-forced subtitle track {0}" -f $i)
                Write-Verbose $s
                continue
            }

            if ($s.Forced -eq "Yes") {
                $forced = ".forced"
            }
            if ($s.Format -eq "PGS") {
                $ext = ".sup"
            } else {
                $ext = ".srt"
            }
            $append = ".{0}{1}{2}" -f $s.Language, $forced, $ext
            $outFile = Join-Path $OutputPath $InputFile.Name.Replace($InputFile.Extension, $append)

            if ((Test-Path $outFile) -eq $true -and
                    $Force -eq $false -and
                    $WhatIf -eq $false) {
                Write-Error "Output file already exists! ($outFile)"
                continue
            }

            Write-Verbose ("Writing subtitle track with ID {0} to {1}" -f $i, $outFile)
            $command = "& '$mkvextract' `"$($InputFile.FullName)`" tracks `"{0}:{1}`"" -f $i, $outFile
            if ($WhatIf -eq $false) {
                Invoke-Expression "$command"
            } else {
                Write-Information "Would execute: $command"
            }
        }
    }
}
