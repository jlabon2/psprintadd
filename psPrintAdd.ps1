## PS PrintAdd
$ErrorActionPreference = 'Continue'

if ($host.name -eq 'ConsoleHost') {
    $SW_HIDE, $SW_SHOW = 0, 5
    $TypeDef = '[DllImport("User32.dll")]public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
    Add-Type -MemberDefinition $TypeDef -Namespace Win32 -Name Functions
    $hWnd = (Get-Process -Id $PID).MainWindowHandle
    $Null = [Win32.Functions]::ShowWindow($hWnd, $SW_HIDE)
    $global:baseConfigPath = Split-Path $PSCommandPath
}

else { $global:baseConfigPath = 'C:\psPrintAdd' }

Add-Type -AssemblyName 'System.Windows.Forms'
foreach ($dll in ((Get-ChildItem -Path (Join-Path $global:baseConfigPath lib) -Filter *.dll).FullName)) { $null = [System.Reflection.Assembly]::LoadFrom($dll) }

Import-Module (Join-Path $global:baseConfigPath -ChildPath internal.psm1) -Force

New-HashTables
Reset-PrintConfig -ConfigHash $configHash
Set-WPFControls -XAMLPath  (Join-Path $global:baseConfigPath -ChildPath Window.xaml) -TargetHash $syncHash    
'Window' | Set-RSDataContext -SyncHash $syncHash -DataContext $configHash.printerSettings

$configHash.printServList = Get-Content (Join-Path $baseConfigPath -ChildPath 'printsrvlist.json') | ConvertFrom-Json
$syncHash.PrintServer_Sel.ItemsSource = $configHash.printServList

#region Find_Items

$syncHash.Find_StartFind.Add_Click( {
        if ($global:RSJobCleanup) { [void]$global:RSJobCleanup.PowerShell.Runspace.Dispose() }
        New-RS -ScriptBlock { Find-PrinterInfo -PrinterConfig $configHash.printerSettings[0] } -RunspaceName UpdatePrintInfo
    })

$syncHash.Find_PrinterURL.add_Click( { Start-Process $configHash.printerSettings[0].PrinterURL })

$syncHash.Find_IPAddressText.Add_TextChanged( {

        if ($configHash.printerSettings[0].PrinterOnline) { Reset-Tool }

        if ($syncHash.Find_IPAddressText.Text -match "^.+[.].+[.].+[.].+$") {
            try {
                [IPAddress]$syncHash.Find_IPAddressText.Text              
                @('Find_ValidIP', 'Find_StartFind') | Set-Tag -SyncHash $syncHash -Tag True 
            }
            catch { @('Find_ValidIP', 'Find_StartFind') | Set-Tag -SyncHash $syncHash -Tag False }     
        }
        else {
            @('Find_ValidIP', 'Find_StartFind') | Set-Tag -SyncHash $syncHash 
        }
    })


#endregion

#region DHCP_Items


$syncHash.DHCP_FindLease.Add_Click( {
        New-RS -ScriptBlock { Find-DHCPLease -IPAddress $configHash.printerSettings[0].PrinterIP } -RunspaceName FindLease
    })

$syncHash.DHCP_ReserveLease.Add_Click( {
        New-RS -ScriptBlock { Add-PrinterIPReservation -PrinterConfig $configHash.printerSettings[0] } -RunspaceName AddReservation
    })

#endregion

#region Print_Server
$syncHash.PrintServer_Sel.Add_SelectionChanged( {
        $syncHash.Port_Sel.ItemsSource = $null
        $syncHash.Driver_Sel.ItemsSource = $null
        $syncHash.Port_InUseName.Text = $null
    })

#endregion

#region Port_Items
$syncHash.Port_GetList.Add_Click( {
        New-RS -ScriptBlock { Get-PortList -PrintServer $configHash.printerSettings[0].PrintServer } -RunspaceName GetPrinterPortList
    })

$syncHash.Port_Add.Add_Click( {
        New-RS -ScriptBlock { Add-PrintServerPort -PrintConfig $configHash.printerSettings[0] } -RunspaceName AddPrinterPort
    })

$syncHash.Port_Sel.Add_SelectionChanged( { Get-DuplicatePort })

#endRegion

#region Driver_Items

$syncHash.Driver_GetList.Add_Click( {
        New-RS -ScriptBlock { Get-DriverList -PrintServer $configHash.printerSettings[0].PrintServer } -RunspaceName GetPrinterDriverList
    })

$syncHash.Driver_Add.Add_Click( {
        New-RS -ScriptBlock { Add-PrintDriver -PrintServer $configHash.printerSettings[0].PrintServer } -RunspaceName AddPrintDriver
    })

#endregion

#region Prop_items

$syncHash.Prop_Name.Add_TextChanged( {
        if ($syncHash.Prop_Name.Text -in $configHash.PortList.Printer) { Update-Field -ControlName Prop_NameDuplicate -SyncHash $syncHash -Value 'Visible' -Property 'Visibility' }
        else { Update-Field -ControlName Prop_NameDuplicate -SyncHash $syncHash -Value 'Collapsed' -Property 'Visibility' }

    })

[System.Windows.RoutedEventHandler]$nameTextChangeEvent = {
    $props = $configHash.printerSettings[0]
    if ($synchash.Prop_UseModel.isChecked) { $string = $syncHash.Prop_Campus.Text + $syncHash.Prop_Department.Text + '-' + ($props.PrinterModel -replace "Brother " -replace '-' -replace ' ') }
    else { $string = $syncHash.Prop_Campus.Text + $syncHash.Prop_Department.Text + '-' }

    $syncHash.Prop_Name.Text = $string.Trim()
}

$syncHash.Prop_Campus.AddHandler([System.Windows.Controls.TextBox]::TextChangedEvent, $nameTextChangeEvent)
$syncHash.Prop_Department.AddHandler([System.Windows.Controls.TextBox]::TextChangedEvent, $nameTextChangeEvent)

[System.Windows.RoutedEventHandler]$locationTextChangeEvent = {
    $string = [string]::Empty
    @('Prop_Campus', 'Prop_Building', 'Prop_Room') | ForEach-Object {
        if (![string]::IsNullOrWhiteSpace($syncHash.$_.Text)) { $string = $string + "$($_ -replace 'Prop_' -replace 'Building','Bldg' -replace 'Room','Rm'): $($syncHash.$_.Text) " }
    }
    $syncHash.Prop_Location.Text = $string.Trim()
}

$syncHash.Prop_Campus.AddHandler([System.Windows.Controls.TextBox]::TextChangedEvent, $locationTextChangeEvent)
$syncHash.Prop_Building.AddHandler([System.Windows.Controls.TextBox]::TextChangedEvent, $locationTextChangeEvent)
$syncHash.Prop_Room.AddHandler([System.Windows.Controls.TextBox]::TextChangedEvent, $locationTextChangeEvent)
#endregion

#region Exception_Items

$syncHash.Exception_Text.Add_TextChanged( {
        if (![string]::IsNullOrWhiteSpace($syncHash.Exception_Text.Text)) {
            $syncHash.Exception_Dialog.IsOpen = $true
        }
    })

$syncHash.Exception_Close.Add_Click( {
        $syncHash.Exception_Dialog.IsOpen = $false
        $syncHash.Exception_Text.Clear()
    })

#endregion

#region Main_Items

$syncHash.Main_Reset.Add_Click( {
        Reset-Tool
    })

$syncHash.Main_Add.Add_Click( { Add-PrinterToServer -PrintSettings $configHash.printerSettings[0] })

$syncHash.Main_OpenPrintMgmt.Add_Click( {
        & printmanagement.msc
    })
#endregion
$syncHash.Window.ShowDialog()