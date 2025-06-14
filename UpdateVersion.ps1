# UpdateVersion.ps1
# This script generates PolicyPlus\Version.vb with the current Git version.

$VersionFilePath = "PolicyPlus\Version.vb"

try {
    $RawGitVersion = git describe --always
    $GitVersion = $RawGitVersion.Trim()
}
catch {
    Write-Error "Error running 'git describe'. Make sure Git is installed and you are in a Git repository."
    Write-Error "Underlying error: $($_.Exception.Message)"
    exit 1
}

# Define the content for Version.vb
$FileContent = @"
' DO NOT MODIFY THIS FILE. To update it, run Version.ps1 again.
Module VersionHolder
    Public Const Version As String = "$GitVersion"
End Module
"@

$VersionFileDirectory = Split-Path -Path $VersionFilePath -Parent
if (-not (Test-Path $VersionFileDirectory)) {
    New-Item -ItemType Directory -Path $VersionFileDirectory -Force | Out-Null
}

# Write the content to Version.vb with UTF-8 encoding
try {
    Set-Content -Path $VersionFilePath -Value $FileContent -Encoding UTF8 -Force
    Write-Host "Successfully updated $VersionFilePath with version: $GitVersion"
}
catch {
    Write-Error "Failed to write to $VersionFilePath. Error: $($_.Exception.Message)"
}
