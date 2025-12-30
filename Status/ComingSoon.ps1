Add-Type -AssemblyName PresentationFramework

[System.Windows.MessageBox]::Show(
    "This function is coming soon.",
    "Advance Setup Tools",
    [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Information
) | Out-Null

exit
