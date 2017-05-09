# Word2Yammer.ps1
# PowerShell script to convert a Word document to Yammer text.
# © 2017 Serge van den Oever
# Features:
# - Word text is exported as UTF-8 so most special characters work
# - Leading spaces are replaced by non-breaking (UTF-8) spaces, so leading spaces are shown (useful for source-code in a post)
# - Dashes are replaced by non-breaking (UTF-8) dashes, so no strange line-breaks are introduced
# - leading and trailing empty lines are removed

#Requires -Version 3.0

param (
    [string]$Path = $null,
    [switch]$Version
)

function Word2Text {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Destination  
    )
    $word = New-Object -ComObject Word.Application
    $wordDoc  = $word.Documents.Open($Path, $false, $true)

    # Check Version of Word Installed and create UTF-8 (unicode) text file (see https://msdn.microsoft.com/en-us/library/office/ff839952.aspx)
    $wordVersion = $word.Version
    If ($wordVersion -eq '16.0' -Or $wordVersion -eq '15.0') {
        $wordDoc.SaveAs($Destination, 7) 
        $wordDoc.Close($false)  
    }
    ElseIf ($wordVersion -eq '14.0') 
    {
        $wordDoc.SaveAs([ref] $Destination,[ref] 7)
        $wordDoc.Close([ref]$false)
    }

    # Close Word
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)

    # Cleanup
    Remove-Variable word
}

function LeadingSpacesTotNonBreakingSpaces
{
    param(
        [string]$TextLine
    )

    if ($TextLine -match "(^\s+)") {
        $numSpaces = $matches[1].Length
        $text = ([string][char]0x00a0)*$numSpaces + $TextLine.Substring($numSpaces)
        return $text
    
    } else {
        return $TextLine
    }
}

# Foreach line:
# - trim trailing spaces
# - replace leading spaces by non-breaking spaces
# - replaces dashes by non-breaking dashes
# Determine the leading and trailing enpty lines, and write lines inbetween lines to file
# Append reference to this script to the file
function YammerizeText {
    param (
        [Parameter(Mandatory=$true)][string]$Path 
    )
    $lines = Get-Content -Path $Path
    for ($i=0; $i -lt $lines.Length; $i++) {
        $s = $lines[$i]
        $s = $s.TrimEnd();
        $s = LeadingSpacesTotNonBreakingSpaces -TextLine $s
        $s = $s.Replace('-', [char]0x2011)
        $lines[$i] = $s
    }

    # determine index of first non-empty line
    $startIndex = 0
    for ($i=0; $i -lt $lines.Length; $i++) {
        if ($lines[$i].Length -eq 0) {
            $startIndex++
        } else {
            break
        }
    }

    # determine index of last non-empty line
    $endIndex = $lines.Length-1
    for ($i=$lines.Length-1; $i -ge 0; $i--) {
        if ($lines[$i].Length -eq 0) {
            $endIndex--
        } else {
            break
        }
    }

    if ($startIndex -gt $endIndex) {
        $lines = $null
    } else {
        $lines = $lines[$startIndex..$endIndex]
    }

    # add lines about tool
    $lines += @(
        "",
        "==== Composed with https://github.com/svdoever/Word2Yammer ===="
    )


    Out-File -FilePath $Path -InputObject $lines -Encoding "UTF8" 
}

[decimal]$scriptVersion=1.00

if ($Version) {
    try {
        $scriptCode = (New-Object System.Net.WebClient).DownloadString('https://rawgit.com/svdoever/Word2Yammer/master/Word2Yammer.ps1')
    } catch {
        Write-Host "Error: "+ $_.Exception.Message
        exit -1    
    }
    $match = $scriptCode.Split("`n") | Select-String -pattern '[decimal]$scriptVersion=' -SimpleMatch

    # Format the version numbers of the two to 2 decimal places
    $oldVersion = $scriptVersion.ToString("N2", [CultureInfo]::InvariantCulture)
    $currentVersion = ([decimal]$match.line.Split("=")[1]).ToString("N2", [CultureInfo]::InvariantCulture)

    Write-Host "Your   version of Word2Yammer.ps1 is: $oldVersion"
    Write-Host "Latest version of Word2Yammer.ps1 is: $currentVersion"

    #  Compare the two version numbers and overwrite the old one with the new one.
    If ($currentVersion -gt $oldVersion)
    {
        Write-Host "New version available of Word2Yammer.ps1 script. Updating the script with the following command:"
        Write-Host "(New-Object System.Net.WebClient).DownloadString('https://rawgit.com/svdoever/Word2Yammer/master/Word2Yammer.ps1') > Word2Yammer.ps1"
    }
    exit 0
}

if ($Path -eq $null -or $Path -eq "") {
    Write-Host "Run the script as .\Word2Yammer.ps1 -Path MyWordsToTheWorld.docx"
    Write-Host "Run .\Word2Yammer.ps1 -Version to check for newer version"
    exit -1
}

$ResolvedPath = Resolve-Path -Path $Path
if ($ResolvedPath -eq $null) { 
    Write-Host "The specified path '$Path' does not exist."
    exit -1
}
$Path = $ResolvedPath.Path

$folder = [System.IO.Path]::GetDirectoryName($Path)
$filename = [System.IO.Path]::GetFileNameWithoutExtension($Path)
$outfile = "$folder\$filename.txt"

Word2Text -Path $Path -Destination $outfile
YammerizeText -Path $outfile