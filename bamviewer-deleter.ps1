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
            default { return "Invalid Signature ($($sig.Status))" }
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

Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml

# Setup BAM paths
$bamPaths = @(
    'HKLM:\SYSTEM\CurrentControlSet\Services\bam\UserSettings',
    'HKLM:\SYSTEM\CurrentControlSet\Services\bam\state\UserSettings'
)

function Load-BamEntries {
    $UsersList = @()
    foreach ($bamPath in $bamPaths) {
        $UsersList += Get-ChildItem -Path $bamPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName
    }
    $Users = $UsersList | Sort-Object -Unique

    $BamEntries = foreach ($Sid in $Users) {
        foreach ($bamPath in $bamPaths) {
            $fullPath = Join-Path $bamPath $Sid
            if (Test-Path $fullPath) {
                try {
                    $props = Get-ItemProperty -Path $fullPath
                    $propNames = $props.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' }

                    foreach ($p in $propNames) {
                        $value = $props.$($p.Name)
                        if ($value.Length -eq 24) {
                            $Hex = [System.BitConverter]::ToString($value[7..0]) -replace '-', ''
                            $utcTime = [DateTime]::FromFileTimeUtc([Convert]::ToInt64($Hex, 16))
                            $localTime = $utcTime.ToLocalTime()

                            try {
                                $userName = (New-Object System.Security.Principal.SecurityIdentifier($Sid)).Translate([System.Security.Principal.NTAccount]).Value
                            } catch {
                                $userName = ''
                            }

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

                            # Determine signature status for coloring
                            $sigStatus = 'Invalid'
                            if ($signature -eq 'Valid Signature') {
                                $sigStatus = 'Valid'
                            } elseif ($signature -eq 'File Was Not Found') {
                                $sigStatus = 'FileNotFound'
                            }

                            [PSCustomObject]@{
                                Application     = $p.Name
                                Path            = $filePath
                                Signature       = $signature
                                SignatureStatus = $sigStatus
                                'LastRunLocal'  = $localTime
                                'LastRunUTC'    = $utcTime
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
    return $BamEntries
}

# XAML for the WPF window
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="BAM Entries Deleter" Height="450" Width="1000"
        Background="Black" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="45"/>
        </Grid.RowDefinitions>

        <DataGrid Name="dgBamEntries" AutoGenerateColumns="False"
                  CanUserAddRows="False" CanUserDeleteRows="False"
                  SelectionMode="Extended" SelectionUnit="FullRow"
                  GridLinesVisibility="Horizontal"
                  HorizontalGridLinesBrush="White"
                  Background="Black" Foreground="White"
                  AlternatingRowBackground="#222222" RowBackground="Black"
                  Grid.Row="0" FontSize="14">
            <DataGrid.Resources>
                <!-- Column Headers -->
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Foreground" Value="Green"/>
                    <Setter Property="Background" Value="Black"/>
                </Style>

                <!-- Row Color Logic -->
                <Style TargetType="DataGridRow">
                    <Setter Property="Foreground" Value="White"/>
                    <Style.Triggers>
                        <DataTrigger Binding="{Binding SignatureStatus}" Value="Valid">
                            <Setter Property="Foreground" Value="LimeGreen"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding SignatureStatus}" Value="Invalid">
                            <Setter Property="Foreground" Value="Red"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding SignatureStatus}" Value="FileNotFound">
                            <Setter Property="Foreground" Value="Red"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding SignatureStatus}" Value="Deleted">
                            <Setter Property="Foreground" Value="Red"/>
                        </DataTrigger>
                    </Style.Triggers>
                </Style>
            </DataGrid.Resources>
            <DataGrid.Columns>
                <DataGridTextColumn Header="Application" Binding="{Binding Application}" Width="2*"/>
                <DataGridTextColumn Header="Signature" Binding="{Binding Signature}" Width="2*"/>
                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="1.5*"/>
                <DataGridTextColumn Header="Path" Binding="{Binding Path}" Width="3*"/>
                <DataGridTextColumn Header="Last Run (Local)" Binding="{Binding LastRunLocal}" Width="1.7*"/>
            </DataGrid.Columns>
        </DataGrid>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Grid.Row="1">
            <Button Name="btnDelete" Content="Delete Selected" Width="130" Margin="0,0,10,0"/>
            <Button Name="btnRefresh" Content="Refresh" Width="90" Margin="0,0,10,0"/>
            <Button Name="btnClose" Content="Close" Width="90"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Find controls
$dgBamEntries = $Window.FindName("dgBamEntries")
$btnDelete = $Window.FindName("btnDelete")
$btnRefresh = $Window.FindName("btnRefresh")
$btnClose = $Window.FindName("btnClose")

# Observable collection for grid
$BamData = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()

function Refresh-GridData {
    $BamData.Clear()
    $entries = Load-BamEntries
    foreach ($entry in $entries) {
        $BamData.Add($entry)
    }
    $dgBamEntries.ItemsSource = $BamData
}

# Load initial data
Refresh-GridData

# Delete button
$btnDelete.Add_Click({
    $selectedItems = $dgBamEntries.SelectedItems
    if ($selectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No BAM entries selected for deletion.", "Warning", "OK", "Warning") | Out-Null
        return
    }

    $result = [System.Windows.MessageBox]::Show("Are you sure you want to DELETE the selected BAM entries?", "Confirm Deletion", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }

    foreach ($entry in $selectedItems) {
        $fullRegPath = Join-Path $entry.RegPath $entry.SID
        $propertyName = $entry.PropertyName

        try {
            $propExists = Get-ItemProperty -Path $fullRegPath -Name $propertyName -ErrorAction SilentlyContinue
            $subkeyPath = Join-Path $fullRegPath $propertyName
            $subkeyExists = Test-Path -Path $subkeyPath

            if ($propExists) {
                Remove-ItemProperty -Path $fullRegPath -Name $propertyName -ErrorAction Stop
                $entry.Signature = "Deleted"
                $entry.SignatureStatus = "Deleted"
            } elseif ($subkeyExists) {
                Remove-Item -Path $subkeyPath -Recurse -Force -ErrorAction Stop
                $entry.Signature = "Deleted"
                $entry.SignatureStatus = "Deleted"
            } else {
                [System.Windows.MessageBox]::Show("Property or subkey '$propertyName' not found at $fullRegPath", "Warning", "OK", "Warning") | Out-Null
            }
        } catch {
            [System.Windows.MessageBox]::Show("Failed to delete '$propertyName': $($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
        }
    }

    $dgBamEntries.Items.Refresh()
})

$btnRefresh.Add_Click({ Refresh-GridData })
$btnClose.Add_Click({ $Window.Close() })

# Show window
$Window.ShowDialog() | Out-Null

Write-Host 'Exiting script. Press Enter to close.' -ForegroundColor Cyan
try {
    if ($Host.Name -eq 'ConsoleHost') {
        [void][System.Console]::ReadKey($true)
    } else {
        Read-Host -Prompt 'Press Enter to exit...'
    }
} catch {
    Read-Host -Prompt 'Press Enter to exit...'
}
