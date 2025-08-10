$ErrorActionPreference = 'SilentlyContinue'

function Get-Signature {
    [CmdletBinding()]
    param ([string[]]$FilePath)

    if (Test-Path -PathType Leaf -Path $FilePath) {
        $sig = Get-AuthenticodeSignature -FilePath $FilePath
        switch ($sig.Status) {
            'Valid' { return 'Valid Signature' }
            'NotSigned' { return 'Invalid Signature (NotSigned)' }
            'HashMismatch' { return 'Invalid Signature (HashMismatch)' }
            'NotTrusted' { return 'Invalid Signature (NotTrusted)' }
            default { return 'Invalid Signature ($($sig.Status))' }
        }
    } else {
        return 'File Was Not Found'
    }
}

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Warning 'Please run this script as Administrator.'
    Start-Sleep 5
    Exit
}

Clear-Host

Add-Type -AssemblyName PresentationFramework  # For MessageBox

# Setup BAM paths
$bamPaths = @(
    'HKLM:\SYSTEM\CurrentControlSet\Services\bam\UserSettings',
    'HKLM:\SYSTEM\CurrentControlSet\Services\bam\state\UserSettings'
)

# Collect all SID folders
$UsersList = @()
foreach ($bamPath in $bamPaths) {
    $UsersList += Get-ChildItem -Path $bamPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName
}
$Users = $UsersList | Sort-Object -Unique


# Build BAM entries
$Bam = foreach ($Sid in $Users) {
    foreach ($bamPath in $bamPaths) {
        $fullPath = Join-Path $bamPath $Sid
        if (Test-Path $fullPath) {
            try {
                $props = Get-ItemProperty -Path $fullPath
                $propNames = $props.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' }

                foreach ($p in $propNames) {
                    $value = $props.$($p.Name)
                    if ($value.Length -eq 24) {
                        # Convert value to hex string for FromFileTimeUtc
                        $Hex = [System.BitConverter]::ToString($value[7..0]) -replace '-', ''
                        $utcTime = [DateTime]::FromFileTimeUtc([Convert]::ToInt64($Hex, 16))
                        $localTime = $utcTime.ToLocalTime()

                        # Resolve user name
                        try {
                            $userName = (New-Object System.Security.Principal.SecurityIdentifier($Sid)).Translate([System.Security.Principal.NTAccount]).Value
                        } catch {
                            $userName = ''
                        }

                        # Attempt to get file path if property name looks like a path
                        $filePath = ''
                        if ($p.Name -like '*\*') {
                            try {
                                $relative = $p.Name.Substring(23)
                                $filePath = Join-Path 'C:\' $relative
                            } catch {
                                $filePath = ''
                            }
                        }

                        $signature = if ($filePath) { Get-Signature -FilePath $filePath } else { '' }

                        [PSCustomObject]@{
                            Application     = $p.Name
                            Path            = $filePath
                            Signature       = $signature
                            'LastRun (Local)'= $localTime
                            'LastRun (UTC)' = $utcTime
                            User            = $userName
                            SID             = $Sid
                            RegPath         = $bamPath
                            PropertyName    = $p.Name
                        }
                    }
                }
            } catch {}
        }
    }
}

do {
    $SelectedEntries = $Bam | Out-GridView -PassThru -Title 'Select BAM key entries to DELETE or Cancel to Exit'

    if (-not $SelectedEntries) { break }

    foreach ($entry in $SelectedEntries) {
    $fullRegPath = Join-Path $entry.RegPath $entry.SID
    $propertyName = $entry.PropertyName

    Write-Host "Attempting to delete BAM key entry '$propertyName' under '$fullRegPath'..." -ForegroundColor Yellow

    $propExists = Get-ItemProperty -Path $fullRegPath -Name $propertyName -ErrorAction SilentlyContinue
    $subkeyPath = Join-Path $fullRegPath $propertyName
    $subkeyExists = Test-Path -Path $subkeyPath

    if ($propExists) {
        try {
            Remove-ItemProperty -Path $fullRegPath -Name $propertyName -ErrorAction Stop
            Write-Host "Deleted property '$propertyName'" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to delete property '$propertyName'"
            Write-Warning $_.Exception.Message
        }
    } elseif ($subkeyExists) {
        try {
            Remove-Item -Path $subkeyPath -Recurse -Force -ErrorAction Stop
            Write-Host "Deleted subkey '$propertyName'" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to delete subkey '$propertyName'"
            Write-Warning $_.Exception.Message
        }
    } else {
        Write-Warning "Property or subkey '$propertyName' not found at $fullRegPath"
    }
}


    $result = [System.Windows.MessageBox]::Show(
        'Do you want to delete more BAM entries?',
        'Continue?',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
} while ($result -eq [System.Windows.MessageBoxResult]::Yes)

Write-Host 'Exiting script. Press Enter to close.' -ForegroundColor Cyan

try {
    if ($Host.Name -eq 'ConsoleHost') {
        [void][System.Console]::ReadKey($true)
    }
    else {
        Read-Host -Prompt 'Press Enter to exit...'
    }
}
catch {
    Read-Host -Prompt 'Press Enter to exit...'
}
