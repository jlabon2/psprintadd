function Set-WPFControls {
    param (
        [Parameter(Mandatory)]$XAMLPath,
        [Parameter(Mandatory)][Hashtable]$TargetHash
    ) 

    $inputXML = Get-Content -Path $XAMLPath
    
    $inputXML = $inputXML -replace 'mc:Ignorable="d"' -replace ' x:Class="RepGen.MainWindow"'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML

    $xmlReader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $XAML)
    
    try { $TargetHash.Window = [Windows.Markup.XamlReader]::Load($xmlReader) }
    catch { Write-Warning -Message "Unable to parse XML, with error: $($Error[0])" }

    ## Load each named control into PS hashtable
   foreach ($controlName in ($XAML.SelectNodes('//*[@Name]').Name)) { $TargetHash.$controlName = $TargetHash.Window.FindName($controlName) }

}

function Set-Tag {
 [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)][String]$ControlName,
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [String]$Tag   
    )
    

    Process {
    $SyncHash.Window.Dispatcher.invoke([action] { $SyncHash.$controlName.Tag = $Tag })
    }
}

function New-HashTables {


    # Stores WPF controls
    $global:syncHash = [hashtable]::Synchronized(@{ })

    # Stores config'd vars
    $global:configHash = [hashtable]::Synchronized(@{ })

}


function Start-RS {
  param(
    [ScriptBlock]$ScriptBlock = { },
    [HashTable] $Variables = @{ } 
  )
  
  $rs = ([RunspaceFactory]::CreateRunspace()).Open()
  
  $ps = [powershell]::Create()
  $ps.Runspace = $rs
      
  [void]$ps.AddScript($ScriptBlock)
  
  ForEach ($Variable in $Variables.GetEnumerator()) { $ps.AddParameter($Variable.Name, $Variable.Value) | Out-Null }

  $ps.BeginInvoke()
  
}

function Reset-PrintConfig {
 [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$ConfigHash 
    )
    
    $configHash.printerSettings = New-Object -TypeName System.Collections.ObjectModel.ObservableCollection[Object]

    [void]$configHash.printerSettings.Add([PSCustomObject]@{
        PrinterOnline   = $null
        PrinterIP       = $null
        PrinterModel    = 'Unknown'
        PrinterURL      = 'Unknown'
        PrinterLease    = $null
        PrinterDHCP     = $null
        PrinterReserved = $null
        DHCPScopeID     = $null
        PrinterMAC      = $null
        PrintServer     = $null
        PrinterPortName = $null
        PrinterDriver   = $null
        PropCampus      = $null
        PropBuilding    = $null
        PropRoom        = $null
        PropDepartment  = $null
        PropName        = $null
        PropLocation    = $null
    })
}

function Update-Field {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$ControlName,
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        $Value,
        [Parameter(Mandatory)][String]$Property
    )

    $SyncHash.Window.Dispatcher.invoke([action] { process {$SyncHash.$controlName.$Property = $value }})}


function Find-PrinterInfo {
    [CmdletBinding()]
    Param ( 
    [Parameter(Mandatory, ValueFromPipeline)]$PrinterConfig)
    
    Update-Field -controlName Find_Progress -SyncHash $syncHash -Value Visible -Property Visibility
    Update-Field -controlName Find_Model -SyncHash $syncHash -Value 'Unknown' -Property Text
    Update-Field -controlName Find_Connectivity -SyncHash $syncHash -Property Text
    Update-Field -controlName Find_PrinterURL -SyncHash $syncHash -Property NavigateUri

    Start-Sleep -Seconds 1

    $connectivity = Test-Connection -ComputerName $PrinterConfig.PrinterIP -Quiet -Count 1

    if ($connectivity) {
        #$PrinterConfig.PrinterOnline = 'Online'
        Update-Field -controlName Find_Connectivity -SyncHash $syncHash -Value Online -Property Text
        try {
            $uri = "http:\\$($PrinterConfig.PrinterIP)"
            $siteTitle = (Invoke-WebRequest -Uri $uri -ErrorAction Stop).ParsedHtml.Title
            if ($siteTitle) { Update-Field -controlName Find_Model -SyncHash $syncHash -Value $siteTitle -Property Text } 
            if (![string]::IsNullOrEmpty($uri -replace 'http:\\')) {Update-Field -controlName Find_PrinterURL -SyncHash $syncHash -Property NavigateUri -Value $uri} 
        }

        catch { 
            
            Update-Field -controlName Find_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility 
        }

    }

    else { Update-Field -controlName Find_Connectivity -SyncHash $syncHash -Value Offline -Property Text }

     Update-Field -controlName Find_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility
    
}


function Set-RSDataContext {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$controlName,
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)]$DataContext   
    )

    Process {
        $SyncHash.Window.Dispatcher.invoke([action] { $SyncHash.$controlName.DataContext = $DataContext })
    }
}

function New-RS {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [string]$RunspaceName,

        [Parameter(Mandatory=$True)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory=$False)]
        [switch]$MirrorCurrentEnv = $True,

        [Parameter(Mandatory=$False)]
        [switch]$Wait
    )

    #region >> Helper Functions

    function NewUniqueString {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory=$False)]
            [string[]]$ArrayOfStrings,
    
            [Parameter(Mandatory=$True)]
            [string]$PossibleNewUniqueString
        )
    
        if (!$ArrayOfStrings -or $ArrayOfStrings.Count -eq 0 -or ![bool]$($ArrayOfStrings -match "[\w]")) {
            $PossibleNewUniqueString
        }
        else {
            $OriginalString = $PossibleNewUniqueString
            $Iteration = 1
            while ($ArrayOfStrings -contains $PossibleNewUniqueString) {
                $AppendedValue = "_$Iteration"
                $PossibleNewUniqueString = $OriginalString + $AppendedValue
                $Iteration++
            }
    
            $PossibleNewUniqueString
        }
    }

    #endregion >> Helper Functions

    #region >> Runspace Prep

    # Create Global Variable Names that don't conflict with other exisiting Global Variables
    $ExistingGlobalVariables = Get-Variable -Scope Global
    $DesiredGlobalVariables = @("RSSyncHash","RSJobCleanup","RSJobs")
    if ($ExistingGlobalVariables.Name -notcontains 'RSSyncHash') {
        $GlobalRSSyncHashName = NewUniqueString -PossibleNewUniqueString "RSSyncHash" -ArrayOfStrings $ExistingGlobalVariables.Name
        Invoke-Expression "`$global:$GlobalRSSyncHashName = [hashtable]::Synchronized(@{})"
        $globalRSSyncHash = Get-Variable -Name $GlobalRSSyncHashName -Scope Global -ValueOnly
    }
    else {
        $GlobalRSSyncHashName = 'RSSyncHash'

        # Also make sure that $RunSpaceName is a unique key in $global:RSSyncHash
        if ($RSSyncHash.Keys -contains $RunSpaceName) {
            $RSNameOriginal = $RunSpaceName
            $RunSpaceName = NewUniqueString -PossibleNewUniqueString $RunSpaceName -ArrayOfStrings $RSSyncHash.Keys
            if ($RSNameOriginal -ne $RunSpaceName) {
                Write-Warning "The RunspaceName '$RSNameOriginal' already exists. Your new RunspaceName will be '$RunSpaceName'"
            }
        }

        $globalRSSyncHash = $global:RSSyncHash
    }
    if ($ExistingGlobalVariables.Name -notcontains 'RSJobCleanup') {
        $GlobalRSJobCleanupName = NewUniqueString -PossibleNewUniqueString "RSJobCleanup" -ArrayOfStrings $ExistingGlobalVariables.Name
        Invoke-Expression "`$global:$GlobalRSJobCleanupName = [hashtable]::Synchronized(@{})"
        $globalRSJobCleanup = Get-Variable -Name $GlobalRSJobCleanupName -Scope Global -ValueOnly
    }
    else {
        $GlobalRSJobCleanupName = 'RSJobCleanup'
        $globalRSJobCleanup = $global:RSJobCleanup
    }
    if ($ExistingGlobalVariables.Name -notcontains 'RSJobs') {
        $GlobalRSJobsName = NewUniqueString -PossibleNewUniqueString "RSJobs" -ArrayOfStrings $ExistingGlobalVariables.Name
        Invoke-Expression "`$global:$GlobalRSJobsName = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())"
        $globalRSJobs = Get-Variable -Name $GlobalRSJobsName -Scope Global -ValueOnly
    }
    else {
        $GlobalRSJobsName = 'RSJobs'
        $globalRSJobs = $global:RSJobs
    }
    $GlobalVariables = @($GlobalSyncHashName,$GlobalRSJobCleanupName,$GlobalRSJobsName)
    #Write-Host "Global Variable names are: $($GlobalVariables -join ", ")"

    # Prep an empty pscustomobject for the RunspaceNameResult Key in $globalRSSyncHash
    $globalRSSyncHash."$RunspaceName`Result" = [pscustomobject]@{}

    #endregion >> Runspace Prep


    ##### BEGIN Runspace Manager Runspace (A Runspace to Manage All Runspaces) #####

    $globalRSJobCleanup.Flag = $True

    if ($ExistingGlobalVariables.Name -notcontains 'RSJobCleanup') {
        #Write-Host '$global:RSJobCleanup does NOT already exists. Creating New Runspace Manager Runspace...'
        $RunspaceMgrRunspace = [runspacefactory]::CreateRunspace()
        if ($PSVersionTable.PSEdition -ne "Core") {
            $RunspaceMgrRunspace.ApartmentState = "STA"
        }
        $RunspaceMgrRunspace.ThreadOptions = "ReuseThread"
        $RunspaceMgrRunspace.Open()

        # Prepare to Receive the Child Runspace Info to the RunspaceManagerRunspace
        $RunspaceMgrRunspace.SessionStateProxy.SetVariable("JobCleanup",$globalRSJobCleanup)
        $RunspaceMgrRunspace.SessionStateProxy.SetVariable("jobs",$globalRSJobs)
        $RunspaceMgrRunspace.SessionStateProxy.SetVariable("SyncHash",$globalRSSyncHash)

        $globalRSJobCleanup.PowerShell = [PowerShell]::Create().AddScript({

            ##### BEGIN Runspace Manager Runspace Helper Functions #####

            # Load the functions we packed up
            $FunctionsForSBUse | foreach { Invoke-Expression $_ }

            ##### END Runspace Manager Runspace Helper Functions #####

            # Routine to handle completed Runspaces
            $ProcessedJobRecords = [System.Collections.ArrayList]::new()
            $SyncHash.ProcessedJobRecords = $ProcessedJobRecords
            while ($JobCleanup.Flag) {
                if ($jobs.Count -gt 0) {
                    $Counter = 0
                    foreach($job in $jobs) { 
                        if ($ProcessedJobRecords.Runspace.InstanceId.Guid -notcontains $job.Runspace.InstanceId.Guid) {
                            $job | Export-CliXml "$HOME\job$Counter.xml" -Force
                            $CollectJobRecordPrep = Import-CliXML -Path "$HOME\job$Counter.xml"
                            Remove-Item -Path "$HOME\job$Counter.xml" -Force
                            $null = $ProcessedJobRecords.Add($CollectJobRecordPrep)
                        }

                        if ($job.AsyncHandle.IsCompleted -or $job.AsyncHandle -eq $null) {
                            [void]$job.PSInstance.EndInvoke($job.AsyncHandle)
                            $job.Runspace.Dispose()
                            $job.PSInstance.Dispose()
                            $job.AsyncHandle = $null
                            $job.PSInstance = $null
                        }
                        $Counter++
                    }

                    # Determine if we can have the Runspace Manager Runspace rest
                    $temparray = $jobs.clone()
                    $temparray | Where-Object {
                        $_.AsyncHandle.IsCompleted -or $_.AsyncHandle -eq $null
                    } | foreach {
                        $temparray.remove($_)
                    }

                    <#
                    if ($temparray.Count -eq 0 -or $temparray.AsyncHandle.IsCompleted -notcontains $False) {
                        $JobCleanup.Flag = $False
                    }
                    #>

                    Start-Sleep -Seconds 5

                    # Optional -
                    # For realtime updates to a GUI depending on changes in data within the $globalRSSyncHash, use
                    # a something like the following (replace with $RSSyncHash properties germane to your project)
                    <#
                    if ($RSSyncHash.WPFInfoDatagrid.Items.Count -ne 0 -and $($RSSynchash.IPArray.Count -ne 0 -or $RSSynchash.IPArray -ne $null)) {
                        if ($RSSyncHash.WPFInfoDatagrid.Items.Count -ge $RSSynchash.IPArray.Count) {
                            Update-Window -Control $RSSyncHash.WPFInfoPleaseWaitLabel -Property Visibility -Value "Hidden"
                        }
                    }
                    #>
                }
            } 
        })

        # Start the RunspaceManagerRunspace
        $globalRSJobCleanup.PowerShell.Runspace = $RunspaceMgrRunspace
        $globalRSJobCleanup.Thread = $globalRSJobCleanup.PowerShell.BeginInvoke()
    }

    ##### END Runspace Manager Runspace #####


    ##### BEGIN New Generic Runspace #####

    $GenericRunspace = [runspacefactory]::CreateRunspace()
    if ($PSVersionTable.PSEdition -ne "Core") {
        $GenericRunspace.ApartmentState = "STA"
    }
    $GenericRunspace.ThreadOptions = "ReuseThread"
    $GenericRunspace.Open()

    # Pass the $globalRSSyncHash to the Generic Runspace so it can read/write properties to it and potentially
    # coordinate with other runspaces
    $GenericRunspace.SessionStateProxy.SetVariable("SyncHash",$globalRSSyncHash)

    # Pass $globalRSJobCleanup and $globalRSJobs to the Generic Runspace so that the Runspace Manager Runspace can manage it
    $GenericRunspace.SessionStateProxy.SetVariable("JobCleanup",$globalRSJobCleanup)
    $GenericRunspace.SessionStateProxy.SetVariable("Jobs",$globalRSJobs)
    $GenericRunspace.SessionStateProxy.SetVariable("ScriptBlock",$ScriptBlock)

    # Pass all other notable environment characteristics 
    if ($MirrorCurrentEnv) {
        [System.Collections.ArrayList]$SetEnvStringArray = @()

        $VariablesNotToForward = @('globalRSSyncHash','RSSyncHash','globalRSJobCleanUp','RSJobCleanup',
        'globalRSJobs','RSJobs','ExistingGlobalVariables','DesiredGlobalVariables','$GlobalRSSyncHashName',
        'RSNameOriginal','GlobalRSJobCleanupName','GlobalRSJobsName','GlobalVariables','RunspaceMgrRunspace',
        'GenericRunspace','ScriptBlock')

        $Variables = Get-Variable
        foreach ($VarObj in $Variables) {
            if ($VariablesNotToForward -notcontains $VarObj.Name) {
                try {
                    $GenericRunspace.SessionStateProxy.SetVariable($VarObj.Name,$VarObj.Value)
                }
                catch {
                    Write-Verbose "Skipping `$$($VarObj.Name)..."
                }
            }
        }

        # Set Environment Variables
        $EnvVariables = Get-ChildItem Env:\
        if ($PSBoundParameters['EnvironmentVariablesToForward'] -and $EnvironmentVariablesToForward -notcontains '*') {
            $EnvVariables = foreach ($VarObj in $EnvVariables) {
                if ($EnvironmentVariablesToForward -contains $VarObj.Name) {
                    $VarObj
                }
            }
        }
        $SetEnvVarsPrep = foreach ($VarObj in $EnvVariables) {
            if ([char[]]$VarObj.Name -contains '(' -or [char[]]$VarObj.Name -contains ' ') {
                $EnvStringArr = @(
                    'try {'
                    $('    ${env:' + $VarObj.Name + '} = ' + "@'`n$($VarObj.Value)`n'@")
                    '}'
                    'catch {'
                    "    Write-Verbose 'Unable to forward environment variable $($VarObj.Name)'"
                    '}'
                )
            }
            else {
                $EnvStringArr = @(
                    'try {'
                    $('    $env:' + $VarObj.Name + ' = ' + "@'`n$($VarObj.Value)`n'@")
                    '}'
                    'catch {'
                    "    Write-Verbose 'Unable to forward environment variable $($VarObj.Name)'"
                    '}'
                )
            }
            $EnvStringArr -join "`n"
        }
        $SetEnvVarsString = $SetEnvVarsPrep -join "`n"

        $null = $SetEnvStringArray.Add($SetEnvVarsString)

        # Set Modules
        $Modules = Get-Module
        if ($PSBoundParameters['ModulesToForward'] -and $ModulesToForward -notcontains '*') {
            $Modules = foreach ($ModObj in $Modules) {
                if ($ModulesToForward -contains $ModObj.Name) {
                    $ModObj
                }
            }
        }

        $ModulesNotToForward = @('MiniLab')

        $SetModulesPrep = foreach ($ModObj in $Modules) {
            if ($ModulesNotToForward -notcontains $ModObj.Name) {
                $ModuleManifestFullPath = $(Get-ChildItem -Path $ModObj.ModuleBase -Recurse -File | Where-Object {
                    $_.Name -eq "$($ModObj.Name).psd1"
                }).FullName

                $ModStringArray = @(
                    '$tempfile = [IO.Path]::Combine([IO.Path]::GetTempPath(), [IO.Path]::GetRandomFileName())'
                    "if (![bool]('$($ModObj.Name)' -match '\.WinModule')) {"
                    '    try {'
                    "        Import-Module '$($ModObj.Name)' -NoClobber -ErrorAction Stop 2>`$tempfile"
                    '    }'
                    '    catch {'
                    '        try {'
                    "            Import-Module '$ModuleManifestFullPath' -NoClobber -ErrorAction Stop 2>`$tempfile"
                    '        }'
                    '        catch {'
                    "            Write-Warning 'Unable to Import-Module $($ModObj.Name)'"
                    '        }'
                    '    }'
                    '}'
                    'if (Test-Path $tempfile) {'
                    '    Remove-Item $tempfile -Force'
                    '}'
                )
                $ModStringArray -join "`n"
            }
        }
        $SetModulesString = $SetModulesPrep -join "`n"

        $null = $SetEnvStringArray.Add($SetModulesString)
    
        # Set Functions
        $Functions = Get-ChildItem Function:\ | Where-Object {![System.String]::IsNullOrWhiteSpace($_.Name)}
        if ($PSBoundParameters['FunctionsToForward'] -and $FunctionsToForward -notcontains '*') {
            $Functions = foreach ($FuncObj in $Functions) {
                if ($FunctionsToForward -contains $FuncObj.Name) {
                    $FuncObj
                }
            }
        }
        $SetFunctionsPrep = foreach ($FuncObj in $Functions) {
            $FunctionText = Invoke-Expression $('@(${Function:' + $FuncObj.Name + '}.Ast.Extent.Text)')
            if ($($FunctionText -split "`n").Count -gt 1) {
                if ($($FunctionText -split "`n")[0] -match "^function ") {
                    if ($($FunctionText -split "`n") -match "^'@") {
                        Write-Warning "Unable to forward function $($FuncObj.Name) due to heredoc string: '@"
                    }
                    else {
                        'Invoke-Expression ' + "@'`n$FunctionText`n'@"
                    }
                }
            }
            elseif ($($FunctionText -split "`n").Count -eq 1) {
                if ($FunctionText -match "^function ") {
                    'Invoke-Expression ' + "@'`n$FunctionText`n'@"
                }
            }
        }
        $SetFunctionsString = $SetFunctionsPrep -join "`n"

        $null = $SetEnvStringArray.Add($SetFunctionsString)

        $GenericRunspace.SessionStateProxy.SetVariable("SetEnvStringArray",$SetEnvStringArray)
    }

    $GenericPSInstance = [powershell]::Create()

    # Define the main PowerShell Script that will run the $ScriptBlock
    $null = $GenericPSInstance.AddScript({
        $SyncHash."$RunSpaceName`Result" | Add-Member -Type NoteProperty -Name Done -Value $False
        $SyncHash."$RunSpaceName`Result" | Add-Member -Type NoteProperty -Name Errors -Value $null
        $SyncHash."$RunSpaceName`Result" | Add-Member -Type NoteProperty -Name ErrorsDetailed -Value $null
        $SyncHash."$RunspaceName`Result".Errors = [System.Collections.ArrayList]::new()
        $SyncHash."$RunspaceName`Result".ErrorsDetailed = [System.Collections.ArrayList]::new()
        $SyncHash."$RunspaceName`Result" | Add-Member -Type NoteProperty -Name ThisRunspace -Value $($(Get-Runspace)[-1])
        [System.Collections.ArrayList]$LiveOutput = @()
        $SyncHash."$RunspaceName`Result" | Add-Member -Type NoteProperty -Name LiveOutput -Value $LiveOutput
        

        
        ##### BEGIN Generic Runspace Helper Functions #####

        # Load the environment we packed up
        if ($SetEnvStringArray) {
            foreach ($obj in $SetEnvStringArray) {
                if (![string]::IsNullOrWhiteSpace($obj)) {
                    try {
                        Invoke-Expression $obj
                    }
                    catch {
                        $null = $SyncHash."$RunSpaceName`Result".Errors.Add($_)

                        $ErrMsg = "Problem with:`n$obj`nError Message:`n" + $($_ | Out-String)
                        $null = $SyncHash."$RunSpaceName`Result".ErrorsDetailed.Add($ErrMsg)
                    }
                }
            }
        }

        ##### END Generic Runspace Helper Functions #####

        ##### BEGIN Script To Run #####


        $Result = Invoke-Expression -Command $ScriptBlock.ToString()
        $SyncHash."$RunSpaceName`Result" | Add-Member -Type NoteProperty -Name Output -Value $Result
        
       

        ##### END Script To Run #####

        $SyncHash."$RunSpaceName`Result".Done = $True
    })

    # Start the Generic Runspace
    $GenericPSInstance.Runspace = $GenericRunspace

    if ($Wait) {
        # The below will make any output of $GenericRunspace available in $Object in current scope
        $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
        $GenericAsyncHandle = $GenericPSInstance.BeginInvoke($Object,$Object)

        $GenericRunspaceInfo = [pscustomobject]@{
            Name            = $RunSpaceName + "Generic"
            PSInstance      = $GenericPSInstance
            Runspace        = $GenericRunspace
            AsyncHandle     = $GenericAsyncHandle
        }
        $null = $globalRSJobs.Add($GenericRunspaceInfo)

        #while ($globalRSSyncHash."$RunSpaceName`Done" -ne $True) {
        while ($GenericAsyncHandle.IsCompleted -ne $True) {
            #Write-Host "Waiting for -ScriptBlock to finish..."
            Start-Sleep -Milliseconds 10
        }

        $globalRSSyncHash."$RunspaceName`Result".Output
        #$Object
    }
    else {
        $HelperRunspace = [runspacefactory]::CreateRunspace()
        if ($PSVersionTable.PSEdition -ne "Core") {
            $HelperRunspace.ApartmentState = "STA"
        }
        $HelperRunspace.ThreadOptions = "ReuseThread"
        $HelperRunspace.Open()

        # Pass the $globalRSSyncHash to the Helper Runspace so it can read/write properties to it and potentially
        # coordinate with other runspaces
        $HelperRunspace.SessionStateProxy.SetVariable("SyncHash",$globalRSSyncHash)

        # Pass $globalRSJobCleanup and $globalRSJobs to the Helper Runspace so that the Runspace Manager Runspace can manage it
        $HelperRunspace.SessionStateProxy.SetVariable("JobCleanup",$globalRSJobCleanup)
        $HelperRunspace.SessionStateProxy.SetVariable("Jobs",$globalRSJobs)

        # Set any other needed variables in the $HelperRunspace
        $HelperRunspace.SessionStateProxy.SetVariable("GenericRunspace",$GenericRunspace)
        $HelperRunspace.SessionStateProxy.SetVariable("GenericPSInstance",$GenericPSInstance)
        $HelperRunspace.SessionStateProxy.SetVariable("RunSpaceName",$RunSpaceName)

        $HelperPSInstance = [powershell]::Create()

        # Define the main PowerShell Script that will run the $ScriptBlock
        $null = $HelperPSInstance.AddScript({
            ##### BEGIN Script To Run #####

            # The below will make any output of $GenericRunspace available in $Object in current scope
            $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
            $GenericAsyncHandle = $GenericPSInstance.BeginInvoke($Object,$Object)

            $GenericRunspaceInfo = [pscustomobject]@{
                Name            = $RunSpaceName + "Generic"
                PSInstance      = $GenericPSInstance
                Runspace        = $GenericRunspace
                AsyncHandle     = $GenericAsyncHandle
            }
            $null = $Jobs.Add($GenericRunspaceInfo)

            #while ($SyncHash."$RunSpaceName`Done" -ne $True) {
            while ($GenericAsyncHandle.IsCompleted -ne $True) {
                #Write-Host "Waiting for -ScriptBlock to finish..."
                Start-Sleep -Milliseconds 10
            }

            ##### END Script To Run #####
        })

        # Start the Helper Runspace
        $HelperPSInstance.Runspace = $HelperRunspace
        $HelperAsyncHandle = $HelperPSInstance.BeginInvoke()

        $HelperRunspaceInfo = [pscustomobject]@{
            Name            = $RunSpaceName + "Helper"
            PSInstance      = $HelperPSInstance
            Runspace        = $HelperRunspace
            AsyncHandle     = $HelperAsyncHandle
        }
        $null = $globalRSJobs.Add($HelperRunspaceInfo)
    }

    ##### END Generic Runspace
}



function Find-DHCPLease {
    param(
        [Parameter(Mandatory)]$IPAddress
    )

    begin {
        Update-Field -ControlName DHCP_Progress -SyncHash $syncHash -Value Visible -Property Visibility
        Update-Field -ControlName DHCP_LeaseFound -SyncHash $syncHash -Property Text -Value No
        Update-Field -ControlName DHCP_Reservation -SyncHash $syncHash -Property Text -Value No
        Update-Field -ControlName DHCP_Server  -SyncHash $syncHash -Property Text  -Value Unknown


        function IPInRange {
            [cmdletbinding()]
            [outputtype([System.Boolean])]
            param(
                # IP Address to find.
                [parameter(Mandatory,
                    Position = 0)]
                [validatescript( {
                        ([System.Net.IPAddress]$_).AddressFamily -eq 'InterNetwork'
                    })]
                [string]
                $IPAddress,

                # Range in which to search using CIDR notation. (ippaddr/bits)
                [parameter(Mandatory)]
                [string]
                $NetworkIP,

                [parameter(Mandatory)]
                [string]
                $Mask

            )

            Function NetMaskToCIDR {
                [CmdletBinding()]
                Param(
                    [String]$SubnetMask = '255.255.255.0'
                )
                $byteRegex = '^(0|128|192|224|240|248|252|254|255)$'
                $invalidMaskMsg = "Invalid SubnetMask specified [$SubnetMask]"
                Try {
                    $netMaskIP = [IPAddress]$SubnetMask
                    $addressBytes = $netMaskIP.GetAddressBytes()

                    $strBuilder = New-Object -TypeName Text.StringBuilder

                    $lastByte = 255
                    foreach ($byte in $addressBytes) {

                        # Validate byte matches net mask value
                        if ($byte -notmatch $byteRegex) {
                            Write-Error -Message $invalidMaskMsg `
                                -Category InvalidArgument `
                                -ErrorAction Stop
                        }
                        elseif ($lastByte -ne 255 -and $byte -gt 0) {
                            Write-Error -Message $invalidMaskMsg `
                                -Category InvalidArgument `
                                -ErrorAction Stop
                        }

                        [void]$strBuilder.Append([Convert]::ToString($byte, 2))
                        $lastByte = $byte
                    }

                    ($strBuilder.ToString().TrimEnd('0')).Length
                }
                Catch {
                    Write-Error -Exception $_.Exception `
                        -Category $_.CategoryInfo.Category
                }
            }

            # Split range into the address and the CIDR notation
            [String]$CIDRAddress = $NetworkIP
            [int]$CIDRBits = NetMaskToCidr -SubnetMask $Mask

            # Address from range and the search address are converted to Int32 and the full mask is calculated from the CIDR notation.
            [int]$BaseAddress = [System.BitConverter]::ToInt32((([System.Net.IPAddress]::Parse($CIDRAddress)).GetAddressBytes()), 0)
            [int]$Address = [System.BitConverter]::ToInt32(([System.Net.IPAddress]::Parse($IPAddress).GetAddressBytes()), 0)
            [int]$Mask = [System.Net.IPAddress]::HostToNetworkOrder(-1 -shl ( 32 - $CIDRBits))

            # Determine whether the address is in the range.
            if (($BaseAddress -band $Mask) -eq ($Address -band $Mask)) {
                $true
            }
            else {
                $false
            }
        }

        function Resolve-Scope {
            [CmdletBinding(DefaultParameterSetName = 'HostName')]
            param (
                [parameter(Mandatory = $true)][IPAddress]$IPAddress,
                [parameter(Mandatory = $true)]$ScopeList
            )



            Process {

          
                foreach ($scope in $scopeList) {
                    if (IPInRange -IPAddress ([string]$ipAddress) -NetworkIP $scope.scopeID -Mask $scope.SubnetMask) { $foundScope = $scope; break }            
                }
            }   
    

            End {
                $foundScope 
            }
        }

    }
    
    process {
        $scopeList = foreach ($dhcpServer in (Get-DhcpServerInDC)) {   
            Get-DhcpServerv4Scope -ComputerName $dhcpServer.DNSName | Select-Object ScopeID, SubnetMask, Name, @{Label = 'DHCPServer'; Expression = { $dhcpServer.DNSName } }
        } 

        $scope = Resolve-Scope -IPAddress $IPAddress -ScopeList $scopeList 
   
        if ($scope) {
            $client = Get-DhcpServerv4Lease -ComputerName $scope.DHCPServer -ScopeId $scope.ScopeId | Where-Object {$_.IPAddress -eq $IPAddress} | Select-Object IPAddress, scopeID, ClientID, AddressState, @{Label = 'DHCPServer'; Expression = { $scope.DHCPServer } }         

            if ($client) {
                if ($client.AddressState -match 'Reservation') { Update-Field -controlName DHCP_Reservation -SyncHash $syncHash -Property Text -Value Yes }
                Update-Field -controlName DHCP_LeaseFound -SyncHash $syncHash -Property Text -Value Yes
                Update-Field -controlName DHCP_Server -SyncHash $syncHash -Property Text -Value $scope.DHCPServer
                Update-Field -controlName DHCP_ScopeID -SyncHash $syncHash -Property Text -Value $scope.ScopeId
                Update-Field -controlName DHCP_ClientID -SyncHash $syncHash -Property Text -Value $client.ClientID
               
            }
        }

        Update-Field -ControlName DHCP_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility
    }
}

function Add-PrinterIPReservation {
    [CmdletBinding()]
    Param ( 
    [Parameter(Mandatory, ValueFromPipeline)]$PrinterConfig)

    Update-Field -ControlName DHCP_Progress -SyncHash $syncHash -Value Visible -Property Visibility
    Start-Sleep -Seconds 1
    try {
        Add-DhcpServerv4Reservation -ComputerName $PrinterConfig.PrinterDHCP -ScopeId $printerConfig.DHCPScopeID -IPAddress $PrinterConfig.PrinterIP -ClientId $printerConfig.PrinterMAC -ErrorAction Stop
        Update-Field -ControlName DHCP_Reservation -SyncHash $syncHash -Value Yes -Property Text
        Update-Field -ControlName DHCP_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility
   }

    catch {
         Update-Field -ControlName Exception_Text -SyncHash $syncHash -Value $_ -Property Text
         Update-Field -ControlName DHCP_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility
    }

    

}

function Get-PortList {
    param ($PrintServer) 
    
    Update-Field -ControlName Port_Progress -SyncHash $syncHash -Value Visible -Property Visibility
    Start-Sleep -Seconds 1

    
    if (!(Test-Connection -ComputerName $printServer -Quiet -Count 1)) {
        Update-Field -ControlName Exception_Text -SyncHash $syncHash -Value "Target print server $PrintServer is offline" -Property Text
    }
    else {
        try {
            [array]$portList = (Get-PrinterPort -ComputerName $PrintServer -ErrorAction Stop | Where-Object {$_.Description -eq 'Standard TCP/IP Port'}).Name
            Update-Field -ControlName Port_Sel -SyncHash $syncHash -Value $portList -Property ItemsSource
        }
        catch {
            Update-Field -ControlName Exception_Text -SyncHash $syncHash -Value $_ -Property Text           
       }
   }
    Update-Field -ControlName Port_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility
}

function Add-PrintServerPort {
    param ($PrintConfig) 
    
    Update-Field -ControlName Port_Progress -SyncHash $syncHash -Value Visible -Property Visibility
    Start-Sleep -Seconds 1
    try {     
        if ($syncHash.Port_Sel.Items -contains $printConfig.PrinterIP) { Update-Field -ControlName Port_Duplicate -SyncHash $syncHash -Value Visible -Property Visibility }
        else {
            $return = (cscript c:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r $printConfig.PrinterIP -h $printConfig.PrinterIP -s $printConfig.PrintServer) | Select-Object -Skip 3 -ErrorAction SilentlyContinue
            if ($return -and $return[0] -notlike "Created/updated port*") {
                Update-Field -ControlName Exception_Text -SyncHash $syncHash -Value $return -Property Text
            }
        }

       Get-PortList -PrintServer $PrintConfig.PrintServer
    }
    catch {
         Update-Field -ControlName Exception_Text -SyncHash $syncHash -Value 'Port list update failed' -Property Text
         Update-Field -ControlName Port_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility
   }
}

function Get-DriverList {
    param ($PrintServer) 
    
    Update-Field -ControlName Driver_Progress -SyncHash $syncHash -Value Visible -Property Visibility
    Start-Sleep -Seconds 1

    if (!(Test-Connection -ComputerName $printServer -Quiet -Count 1)) {
        Update-Field -ControlName Exception_Text -SyncHash $syncHash -Value "Target print server $PrintServer is offline" -Property Text
    }
    else {

        try {
            [array]$driverList = (Get-PrinterDriver -ComputerName $PrintServer | Where-Object {$_.Manufacturer -ne 'Microsoft'}).Name
            Update-Field -ControlName Driver_Sel -SyncHash $syncHash -Value $driverList -Property ItemsSource
        }
        catch {
             Update-Field -ControlName Exception_Text -SyncHash $syncHash -Value 'Driver list update failed' -Property Text             
       }      
    }
    Update-Field -ControlName Driver_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility
}

function Add-PrintDriver {
    param ($PrintServer) 
    
    Update-Field -ControlName Driver_Progress -SyncHash $syncHash -Value Visible -Property Visibility
    Start-Sleep -Seconds 1
    try {     
        Start-Process -FilePath printui -ArgumentList "/s /t 2 /c\\$($PrintServer)" -Wait       
        Get-DriverList -PrintServer $PrintServer
    }
    catch {
         Update-Field -ControlName Exception_Text -SyncHash $syncHash -Value $_ -Property Text
         Update-Field -ControlName Driver_Progress -SyncHash $syncHash -Value Collapsed -Property Visibility
   }
}

function Reset-Tool {

    # Reset IP derived settings
     @('DHCP_Server', 'DHCP_Reservation', 'DHCP_LeaseFound', 'Find_Connectivity' ) | ForEach-Object {Update-Field -SyncHash $syncHash -Property Text -ControlName $_}

    # Reset entered properties
    $syncHash.Prop_Wrap.Children | ForEach-Object {$_.Clear()}

    # Reset checkbox defaults
    ($syncHash.Prop_ServerSpool, $syncHash.Prop_SNMP, $syncHash.Prop_PublishAD) | ForEach-Object {$_.IsChecked = $true}
    ($syncHash.Prop_URL, $syncHash.DHCP_Static) | ForEach-Object {$_.IsChecked = $false}

    # Reset printer sel boxes
    $syncHash.PrintServer_Sel.SelectedValue = $null
    $syncHash.Port_Sel.SelectedValue = $null
    $syncHash.Driver_Sel.SelectedValue = $null
}
