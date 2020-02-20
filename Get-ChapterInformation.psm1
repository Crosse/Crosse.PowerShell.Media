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
    Retrieves video chapter information from online sources.

    .DESCRIPTION
    Retrieves video chapter information from online sources.

    .INPUTS
    None.

    .OUTPUTS
    None.

#>
################################################################################
function Get-ChapterInformation {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The title to search for.
            $Title,

            [Parameter(Mandatory=$false)]
            [int]
            # The number of chapters.
            $ChapterCount,

            [Parameter(Mandatory=$false)]
            [switch]
            # Indicates whether to return all results or just a single result deemed the "best".
            $BestResult = $true
          )

    $escapedTitle = [Uri]::EscapeUriString($Title).Replace("*", "").Replace("'", "")

    $results = @()

    $chaptersDbApi = "https://chapterdb.plex.tv/chapters/search"
    $chaptersDbRequest = "{0}?title={1}&chapterCount={2}" -f $chaptersDbApi, $escapedTitle, $ChapterCount

    $response = Invoke-WebRequest -Uri $chaptersDbRequest -Method GET
    if ($response.StatusCode -ne 200) {
        Write-Error "Request unsuccessful. ($($response.StatusCode), $($response.StatusDescription))"
        return
    }

    $info = ([xml]($response.Content)).results.chapterInfo

    foreach ($result in $info) {
        $chaptersDbRequest = "https://chapterdb.plex.tv/chapters/{0}" -f $result.ref.chapterSetId
        $response = Invoke-WebRequest -Uri $chaptersDbRequest -Method GET
        if ($response.StatusCode -ne 200) {
            Write-Error "Request unsuccessful. ($($response.StatusCode), $($response.StatusDescription))"
            return
        }
        $chapterInfo = ([xml]($response.Content)).chapterInfo.chapters.chapter

        if ($chapterInfo.Count -ne $ChapterCount) {
            continue
        }

        $chapters = @()
        for ($index = 0; $index -lt $chapterInfo.Count; $index++) {
            $chapters += New-Object PSObject -Property @{
                Index = $index + 1
                Time = [TimeSpan]$chapterInfo[$index].time
                Title = $chapterInfo[$index].name
            }
        }

        $results += New-Object PSObject -Property @{
            Source = "ChaptersDb"
            ChaptersDbConfirmations = [int]$result.confirmations
            Title = $result.title
            Chapters = $chapters | Sort-Object Index
        }
    }

    Write-Verbose "Found $($chapters.Count) matches from ChaptersDb."

    if ($BestResult) {
        return @($results | Sort-Object Confirmations -Descending)[0]
    } else {
        return $results
    }
}
