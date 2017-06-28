<#
.SYNOPSIS
    PowerShell script for deploying patches via a GUI interface.
.DESCRIPTION
    PowerShell script for deploying patches via a GUI interface.
.NOTES
    Author:  Stefan M. Werner
    Website: http://getninjad.com
#>

# get current directoy
$currentdir             = Split-Path $($MyInvocation.MyCommand.Path)

# Set max number of concurrent threads
$threads				= 6

# Set Switch Defaults
$defaultSwitchExe 		= '/quiet /norestart'
$defaultSwitchMsu 		= '/quiet /norestart'
$defaultSwitchMsi 		= '/qn /norestart'
$defaultSwitchMsp 		= '/qn'

Function Add-DeployOption
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $key,
        
        [Parameter(Mandatory = $true)]
        [string]
        $value
    )
    
    $obj = New-Object -TypeName PSObject -Property @{
        key = $key
        value  = $value
    }

    # Add results to array
    $script:array += $obj
}

function Add-ComputerName 
{
    Param (
        [Parameter(Mandatory=$true)]
        [string]
        $ComputerName
    )  
	
    # Clear text box contents
    $syncHash.Window.txtComputerName.Clear()
	
    # Check that input value isn't blank
    If ($ComputerName -eq '') 
    {
        Write-Warning "You can't add a blank line."
    } 
    Else 
    {
	
        # Check that computer isn't already in list
        $listItems = $syncHash.Window.listComputers.Items
        $duplicate = $false

        Foreach ($listItem in $listItems) 
        {
            If ($listItem -eq $ComputerName)
            {
                $duplicate = $true
				
                Write-Warning "$ComputerName can't be added as it is already in the list."
            }
        }
			
        If ($duplicate -ne $true)
        {	
            # Add to ListBox
            $syncHash.Window.listComputers.Items.Add("$ComputerName")
        }
    }
}

Function Add-Package 
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $Filename,
        
        [Parameter(Mandatory=$true)]
        [string]
        $FilePath,
        
        [Parameter(Mandatory=$true)]
        [string]
        $PackageName,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet('x86','x64','all')]
        [string]
        $Type
    )
    
    Copy-Item -Path "$FilePath\$Filename" -Destination "$currentdir\deploy\$PackageName"
    
    $syncHash.Counters.top++
    
    [int]$pct = ($syncHash.Counters.top/$syncHash.Counters.FileCount) * 100
    $syncHash.Window.ProgressBarTop.Value = $pct
    
    $syncHash.Counters.bottom++
    
    [int]$pct = ($syncHash.Counters.bottom/$syncHash.Counters.AllCount) * 100
    $syncHash.Window.ProgressBarBottom.Value = $pct
    
    If ($Type -eq 'x86')
    {
        $bat = "$currentdir\deploy\$PackageName\Install-86.bat"
    }
    ElseIf ($Type -eq 'x64')
    {
        $bat = "$currentdir\deploy\$PackageName\Install-64.bat"
    }
    ElseIf ($Type -eq 'all')
    {
        $bat = "$currentdir\deploy\$PackageName\Install-all.bat"
    }
    
    Add-Content -Path $bat -value "cd C:\TempInstall\$PackageName"
     
    If ($Filename -like "*.msi")
    {
        Add-Content -Path $bat -Value "START /WAIT MsiExec.exe /i $Filename $($syncHash.Window.txtSwitches.Text)"
    }
    ElseIf ($Filename -like "*.exe")
    {
        Add-Content -Path $bat -Value "START /WAIT $Filename $($syncHash.Window.txtSwitches.Text)"
    }
    ElseIf ($Filename -like "*.msu")
    {
        Add-Content -Path $bat -Value "START /WAIT WUSA.exe $Filename $($syncHash.Window.txtSwitches.Text)"
    }
    ElseIf ($Filename -like "*.msp")
    {
        Add-Content -Path $bat -Value "START /WAIT MsiExec.exe /p $Filename $($syncHash.Window.txtSwitches.Text)"
    }
    
    $syncHash.Counters.top++
    
    [int]$pct = ($syncHash.Counters.top/$syncHash.Counters.FileCount) * 100
    $syncHash.Window.ProgressBarTop.Value = $pct
    
    $syncHash.Counters.bottom++
    
    [int]$pct = ($syncHash.Counters.bottom/$syncHash.Counters.AllCount) * 100
    $syncHash.Window.ProgressBarBottom.Value = $pct
}

Function Get-RunspaceData 
{
    Do 
    {
        $more = $false        
		
        Foreach($runspace in $runspaces) 
        {
            If ($runspace.Runspace.isCompleted) 
            {
                $runspace.powershell.dispose()
                $runspace.Runspace = $null
                $runspace.powershell = $null
            } 
            ElseIf ($runspace.Runspace -ne $null) 
            {
                $more = $true
            }
        }
		
        If ($more) 
        {
            Start-Sleep -Milliseconds 100
        }
		
        #Clean out unused runspace jobs
        $temphash = $runspaces.clone()
		
        $temphash | Where {	$_.runspace -eq $Null } | ForEach {
            Write-Verbose ("Removing {0}" -f $_.computer)
            $Runspaces.remove($_)
        }
    } while ($more)
}

function Convert-XAMLtoWindow
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $XAML,
        
        [string[]]
        $NamedElements,
        
        [switch]
        $PassThru
    )
    
    Add-Type -AssemblyName PresentationFramework
    
    $reader = [System.XML.XMLReader]::Create([System.IO.StringReader]$XAML)
    $result = [System.Windows.Markup.XAMLReader]::Load($reader)
    foreach($Name in $NamedElements)
    {
        $result | Add-Member NoteProperty -Name $Name -Value $result.FindName($Name) -Force
    }
    
    if ($PassThru)
    {
        $result
    }
    else
    {
        $result.ShowDialog()
    }
}

$xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   xmlns:d="http://schemas.microsoft.com/expression/blend/2008" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="d"
   MinHeight="400"
   Width="900"
   
   SizeToContent="Height"
   Title="Patch Manager v.2.0.1" ResizeMode="NoResize" d:DesignHeight="530.48" WindowStyle="SingleBorderWindow" BorderBrush="Black" Background="White" OpacityMask="{x:Null}" WindowStartupLocation="CenterScreen" ScrollViewer.VerticalScrollBarVisibility="Disabled">
    <DockPanel RenderTransformOrigin="0.749,0.928">
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_File">
                <MenuItem x:Name="menuExit" Header="_Exit" />
            </MenuItem>
        </Menu>
        <Canvas x:Name="canStart" Height="483" VerticalAlignment="Top" Width="890">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnStart" Content="Start" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <TextBlock x:Name="textStart" Canvas.Left="46" TextWrapping="Wrap" Canvas.Top="79" Height="245" Width="442"><Run FontWeight="Bold" FontSize="18.667" Text="Welcome to Patch Manager! "/><LineBreak/><Run/><LineBreak/><Run FontSize="16" Text="Patch Manager is a GUI driven PowerShell script, giving you access to a simple point and click interface with the power to remotely deploy patches and software with a few simple mouse clicks"/><Run Text=". "/><LineBreak/><Run/><LineBreak/><Run FontSize="16" Text="The script is structure in the familiar Windows Wizard style that most should be familiar with. "/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="Click &quot;start&quot; when you are ready to begin."/></TextBlock>
            <Image Source="C:\Users\1005245768E\Pictures\people_cartoon_wizard_in_sorcerers_robe_and_a_magi.jpg" x:Name="image" Height="316" Canvas.Left="506" Canvas.Top="47" Width="326"/>
        </Canvas>
        <Canvas x:Name="canStep1" Height="483" VerticalAlignment="Top" Width="890" Visibility="Collapsed">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnNext" Content="Next" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <Button x:Name="btnBack" Content="Back" Height="33" Canvas.Left="12" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <RadioButton x:Name="radioImport" Content="Import from file" Canvas.Left="569" Canvas.Top="172" FontSize="16" GroupName="BuildList" Height="21" IsChecked="True"/>
            <RadioButton x:Name="radioManual" Content="Manually add systems" Canvas.Left="569" Canvas.Top="206" FontSize="16" GroupName="BuildList" Height="21"/>
            <TextBlock x:Name="textBlock" Height="272" Canvas.Left="52" TextWrapping="Wrap" Canvas.Top="64" Width="388"><Run FontWeight="Bold" FontSize="18.667" Text="Step 1:"/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="Before we get started, please select how you would like to build your list of target systems. "/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="You can either import your list from a CSV, or TXT file. If you have a list exported from the ServiceDesk, you can import it here. TXT files must be in the typical one device per line format."/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="Alternately you can manually type in computer names to add to your list. "/></TextBlock>
            <Separator Height="30" Canvas.Left="285" Margin="0" Canvas.Top="199" Width="394" RenderTransformOrigin="0.5,0.5">
                <Separator.RenderTransform>
                    <TransformGroup>
                        <ScaleTransform/>
                        <SkewTransform/>
                        <RotateTransform Angle="90"/>
                        <TranslateTransform/>
                    </TransformGroup>
                </Separator.RenderTransform>
            </Separator>
        </Canvas>
        <Canvas x:Name="canStep2" Height="483" VerticalAlignment="Top" Width="890" Visibility="Collapsed">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnImport" Content="Import" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16" IsEnabled="False"/>
            <Button x:Name="btnBack2" Content="Back" Height="33" Canvas.Left="12" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <Button Name="btnOpenFile" Canvas.Left="516" Canvas.Top="231" Height="33" Width="149" FontSize="16">Open file</Button>
            <TextBox x:Name="txtFile" Height="33" Canvas.Left="205" TextWrapping="Wrap" Canvas.Top="231" Width="295" FontSize="16" Padding="5"/>
            <TextBlock x:Name="textBlock1" Height="99" Canvas.Left="190" TextWrapping="Wrap" Canvas.Top="114" Width="509"><Run FontWeight="Bold" FontSize="18.667" Text="Step 2:"/><LineBreak/><Run/><LineBreak/><Run FontSize="16" Text="Select a TXT or CSV file to import and click Import in the bottom right corner. "/></TextBlock>
        </Canvas>
        <Canvas x:Name="canStep3" Height="483" VerticalAlignment="Top" Width="890" Visibility="Collapsed">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnNext2" Content="Next" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16" IsEnabled="False"/>
            <Button x:Name="btnBack3" Content="Back" Height="33" Canvas.Left="12" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <TextBlock x:Name="textBlock2" Height="394" Canvas.Left="12" TextWrapping="Wrap" Canvas.Top="20" Width="396"><Run FontWeight="Bold" FontSize="18.667" Text="Step 3:"/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="Verify your list of systems to ensure it was properly imported, and make any necessary changes. This is the last opportunity to edit the list before proceeding with the deployment."/></TextBlock>
            <ListBox x:Name="listComputers" Height="311" Canvas.Left="431" Canvas.Top="20" Width="443" FontSize="16"/>
            <TextBox x:Name="txtComputerName" Height="33" Canvas.Left="430" TextWrapping="Wrap" Canvas.Top="343"  FontSize="16" Padding="5" Width="444"/>
            <Button x:Name="btnAddComputer" Content="Add" Height="33" Canvas.Left="430" Canvas.Top="381" Width="220" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <Button x:Name="btnRemoveComputer" Content="Remove" Height="33" Canvas.Left="656" Canvas.Top="381" Width="220" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
        </Canvas>
        <Canvas x:Name="canStep4" Height="483" VerticalAlignment="Top" Width="890" Visibility="Collapsed">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnNext3" Content="Next" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <Button x:Name="btnBack4" Content="Back" Height="33" Canvas.Left="12" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <TextBlock x:Name="textBlock3" Height="394" Canvas.Left="12" TextWrapping="Wrap" Canvas.Top="20" Width="365"><Run FontWeight="Bold" FontSize="18.667" Text="Step 4:"/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="Here you will select the files you would like to install. You can select separate files for x86 and x64 systems, or single file which will be applied to both. "/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="At this point you must also determine what credential you would like to use for this installation. By default we use the local system account. "/></TextBlock>
            <TextBox x:Name="txtFileName" Height="33" Canvas.Left="430" TextWrapping="Wrap" Canvas.Top="131"  FontSize="16" Padding="5" Width="278"/>
            <Button x:Name="btnAddFile" Content="Add" Height="33" Canvas.Left="716" Canvas.Top="131" Width="158" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <TextBox x:Name="txtFileNameX64" Height="33" Canvas.Left="430" TextWrapping="Wrap" Canvas.Top="264"  FontSize="16" Padding="5" Width="278" IsEnabled="False"/>
            <Button x:Name="btnAddFileX64" Content="Add" Height="33" Canvas.Left="716" Canvas.Top="264" Width="158" RenderTransformOrigin="0.491,0.484" FontSize="16" IsEnabled="False"/>
            <TextBlock x:Name="textBlock4" Height="60" Canvas.Left="430" TextWrapping="Wrap" Text="Select a file you would like to install. If there are separate versions for x86 and x64 systems, select the x86 file first. " Canvas.Top="66" Width="434" FontSize="16"/>
            <TextBlock x:Name="textBlock5" Height="33" Canvas.Left="430" TextWrapping="Wrap" Text="Select a file for x64 systems." Canvas.Top="225" Width="444" FontSize="16"/>
            <CheckBox x:Name="checkBox" Content="Install file based on system architecture." Canvas.Left="440" Canvas.Top="188"/>
            <ComboBox x:Name="comboBox" Height="33" Canvas.Left="430" Canvas.Top="369" Width="444" FontSize="16" Padding="6,5,5,5">
                <ComboBoxItem x:Name="CurrentUser" Content="Current User" HorizontalAlignment="Left" Width="442"/>
                <ComboBoxItem x:Name="LocalSystem" Content="Local System (default)" HorizontalAlignment="Left" Width="442" IsSelected="True"/>
                <ComboBoxItem x:Name="ProvideCred" Content="Provide Credential" HorizontalAlignment="Left" Width="442" IsEnabled="False"/>
            </ComboBox>
            <Label x:Name="label" Content="Install As:" Canvas.Left="430" Canvas.Top="331" FontSize="16" Height="33"/>
            <Separator Height="30" Canvas.Left="205" Margin="0" Canvas.Top="199" Width="394" RenderTransformOrigin="0.5,0.5">
                <Separator.RenderTransform>
                    <TransformGroup>
                        <ScaleTransform/>
                        <SkewTransform/>
                        <RotateTransform Angle="90"/>
                        <TranslateTransform/>
                    </TransformGroup>
                </Separator.RenderTransform>
            </Separator>
        </Canvas>
        <Canvas x:Name="canStep5" Height="483" VerticalAlignment="Top" Width="890" Visibility="Collapsed">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnNext4" Content="Next" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <Button x:Name="btnBack5" Content="Back" Height="33" Canvas.Left="12" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <TextBlock x:Name="textBlock6" Height="394" Canvas.Left="12" TextWrapping="Wrap" Canvas.Top="20" Width="365"><Run FontWeight="Bold" FontSize="18.667" Text="Step 5:"/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="We set some default switches for you based on the type of files you have chosen for this deployment. Please ensure these are correct and make any corrections as needed."/></TextBlock>
            <Label x:Name="label1" Content="Install Switches:" Canvas.Left="442" Canvas.Top="66" FontSize="16" FontWeight="Bold" Padding="0"/>
            <TextBox x:Name="txtSwitches" Height="33" Canvas.Left="442" TextWrapping="Wrap" Canvas.Top="92" Width="432" FontSize="16" Padding="5"/>
            <Separator Height="30" Canvas.Left="205" Margin="0" Canvas.Top="199" Width="394" RenderTransformOrigin="0.5,0.5">
                <Separator.RenderTransform>
                    <TransformGroup>
                        <ScaleTransform/>
                        <SkewTransform/>
                        <RotateTransform Angle="90"/>
                        <TranslateTransform/>
                    </TransformGroup>
                </Separator.RenderTransform>
            </Separator>
            <TextBlock x:Name="textBlock7" Height="264" Canvas.Left="442" TextWrapping="Wrap" Canvas.Top="150" Width="432"><Run Text="Default Switches:"/><LineBreak/><Run/><LineBreak/><Run Text=".msu&#x9;-quiet -norestart"/><LineBreak/><Run Text=".msi&#x9;-qn -norestart"/><LineBreak/><Run Text=".exe&#x9;-quiet -norestart"/><LineBreak/><Run/><LineBreak/><Run Text="Based on suggested silent install switches per Microsoft knowledge base articles 227091, 912203, and 934307."/></TextBlock>
            <Label x:Name="label1_x64" Content="x64 Install Switches:" Canvas.Left="442" Canvas.Top="309" FontSize="16" FontWeight="Bold" Padding="0" Visibility="Collapsed"/>
            <TextBox x:Name="txtSwitches_x64" Height="33" Canvas.Left="442" TextWrapping="Wrap" Canvas.Top="335" Width="432" FontSize="16" Padding="5" Visibility="Collapsed"/>
        </Canvas>
        <Canvas x:Name="canStep6" Height="483" VerticalAlignment="Top" Width="890" Visibility="Collapsed">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnDeploy" Content="Deploy!" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <Button x:Name="btnBack6" Content="Back" Height="33" Canvas.Left="12" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <TextBlock x:Name="textBlock8" Height="394" Canvas.Left="12" TextWrapping="Wrap" Canvas.Top="20" Width="365"><Run FontWeight="Bold" FontSize="18.667" Text="Step 6:"/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="Please confirm your settings and click &quot;Deploy!&quot; when you confident in your selections. To make changes simply click the back button."/></TextBlock>
            <DataGrid x:Name="datDeploymentSetting" Height="409" Canvas.Left="406" Canvas.Top="10" Width="468" IsReadOnly="True" CanUserAddRows="False" CanUserDeleteRows="False">
                <DataGrid.Columns>
                    <DataGridTextColumn x:Name="Key" Header="Key" Width="150" Binding="{Binding key}"/>
                    <DataGridTextColumn x:Name="Value" Header="Value" Width="318" Binding="{Binding value}"/>
                </DataGrid.Columns>
            </DataGrid>
        </Canvas>
        <Canvas x:Name="canStep7" Height="483" VerticalAlignment="Top" Width="890" Visibility="Collapsed">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnClean" Content="Cleanup" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16" IsEnabled="False"/>
            <ProgressBar x:Name="ProgressBarTop" Height="32" Canvas.Left="10" Canvas.Top="145" Width="870"/>
            <ProgressBar x:Name="ProgressBarBottom" Height="32" Canvas.Left="10" Canvas.Top="251" Width="870"/>
            <Label x:Name="lblProgressTop" Content="Label" Height="33" Canvas.Left="10" Canvas.Top="107" Width="506" FontSize="16" Padding="0"/>
            <Label x:Name="lblProgressBottom" Content="Label" Height="33" Canvas.Left="10" Canvas.Top="213" Width="506" FontSize="16" Padding="0"/>
        </Canvas>
        <Canvas x:Name="canStep8" Height="483" VerticalAlignment="Top" Width="890" Visibility="Collapsed">
            <Separator Height="16" Canvas.Left="10" Margin="0" Canvas.Top="419" Width="864"/>
            <Button x:Name="btnDone" Content="Done" Height="33" Canvas.Left="725" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <Button x:Name="btnNew" Content="New Deployment" Height="33" Canvas.Left="12" Canvas.Top="440" Width="149" RenderTransformOrigin="0.491,0.484" FontSize="16"/>
            <TextBlock x:Name="textBlock9" Height="125" Canvas.Left="60" TextWrapping="Wrap" Canvas.Top="135" Width="365"><Run FontWeight="Bold" FontSize="18.667" Text="Deployment Completed"/><LineBreak/><Run FontSize="16"/><LineBreak/><Run FontSize="16" Text="Your deployment has completed. You choose to run a new deployment below, or click done to exit the application. "/></TextBlock>
            <Image Source="C:\Users\1005245768E\Pictures\done.png" x:Name="image2" Height="300" Canvas.Left="489" Canvas.Top="51" Width="300"/>
        </Canvas>
    </DockPanel>

</Window>
'@

$syncHash = [Hashtable]::Synchronized(@{})
$syncHash.Window = Convert-XAMLtoWindow -XAML $xaml -NamedElements 'menuExit', 'canStart', 'btnStart', 'textStart', 'image', 'canStep1', 'btnNext', 'btnBack', 'radioImport', 'radioManual', 'textBlock', 'canStep2', 'btnImport', 'btnBack2', 'btnOpenFile', 'txtFile', 'textBlock1', 'canStep3', 'btnNext2', 'btnBack3', 'textBlock2', 'listComputers', 'txtComputerName', 'btnAddComputer', 'btnRemoveComputer', 'canStep4', 'btnNext3', 'btnBack4', 'textBlock3', 'txtFileName', 'btnAddFile', 'txtFileNameX64', 'btnAddFileX64', 'textBlock4', 'textBlock5', 'checkBox', 'comboBox', 'CurrentUser', 'LocalSystem', 'ProvideCred', 'label', 'canStep5', 'btnNext4', 'btnBack5', 'textBlock6', 'label1', 'txtSwitches', 'textBlock7', 'label1_x64', 'txtSwitches_x64', 'canStep6', 'btnDeploy', 'btnBack6', 'textBlock8', 'datDeploymentSetting', 'Key', 'Value', 'canStep7', 'btnClean', 'ProgressBarTop', 'ProgressBarBottom', 'lblProgressTop', 'lblProgressBottom', 'canStep8', 'btnDone', 'btnNew', 'textBlock9', 'image2' -PassThru

$syncHash.Window.image.Source = "$currentdir\app\img\wizard.jpg"
$syncHash.Window.image2.Source = "$currentdir\app\img\done.png"
$syncHash.Window.icon = "$currentdir\app\img\Patch_management_icon.png"

$runspaces = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))

#$syncHash.Window.canStart.Visibility = 'Visible'
#$syncHash.Window.canStep6.Visibility = 'Collapsed'   

$syncHash.Window.btnStart.add_Click(
    {
        $syncHash.Window.canStart.Visibility = 'Collapsed'
        $syncHash.Window.canStep1.Visibility = 'Visible'        
    }
)

$syncHash.Window.btnNext.add_Click(
    {
        $syncHash.Window.canStep1.Visibility = 'Collapsed'

        If ($syncHash.Window.radioImport.IsChecked -eq $true)
        {
            $syncHash.Window.canStep2.Visibility = 'Visible' 
        }
        Else
        {
            $syncHash.Window.canStep3.Visibility = 'Visible'
        } 
    }
)

$syncHash.Window.btnImport.add_Click(
    {
        If ($syncHash.Window.txtFile.Text -like "*.csv")
        { 
            # Import CSV file
            $rows = Import-Csv -Path $syncHash.Window.txtFile.Text
		    
            # Add Computers to List
            foreach ($row in $rows) 
            {
                Add-ComputerName -ComputerName $row.ComputerName
            }
        }
        ElseIf ($syncHash.Window.txtFile.Text -like "*.txt")
        {
            # Import text file
            $list = Get-Content -Path $syncHash.Window.txtFile.Text

            # Add Computers to List
            foreach ($item in $list) 
            {
                Add-ComputerName -ComputerName $item
            }
        }
        
        $syncHash.Window.canStep2.Visibility = 'Collapsed'
        $syncHash.Window.canStep3.Visibility = 'Visible'        
    }
)

$syncHash.Window.btnNext2.add_Click(
    {
        $syncHash.Window.canStep3.Visibility = 'Collapsed'
        $syncHash.Window.canStep4.Visibility = 'Visible'        
    }
)

$syncHash.Window.btnNext3.add_Click(
    {
        $next = $true
        
        If ($syncHash.Window.txtFileName.Text -ne '')
        {
            If ($syncHash.Window.checkBox.IsChecked)
            {
                If ($syncHash.Window.txtFileNameX64.Text -eq '')
                {
                    $next = $false
                }
                
                $syncHash.Window.label1_x64.Visibility = 'Visible'
                $syncHash.Window.txtSwitches_x64.Visibility = 'Visible'
            }
            Else
            {
                $syncHash.Window.label1_x64.Visibility = 'Collapsed'
                $syncHash.Window.txtSwitches_x64.Visibility = 'Collapsed'
            }
        }
        Else
        {
            $next = $false
        }
        
        If ($next)
        { 
            $syncHash.Window.canStep4.Visibility = 'Collapsed'
            $syncHash.Window.canStep5.Visibility = 'Visible'
            
            $file = $syncHash.Window.txtFileName.Text
            
            If ($file -like "*.exe")
            {
                $syncHash.Window.txtSwitches.Text = $defaultSwitchExe
            }
            ElseIf ($file -like "*.msi")
            {
                $syncHash.Window.txtSwitches.Text = $defaultSwitchMsi
            }
            ElseIf ($file -like "*.msu")
            {
                $syncHash.Window.txtSwitches.Text = $defaultSwitchMsu
            }
            
            If ($syncHash.Window.txtFileNameX64.Text -eq '')
            {
                $file = $syncHash.Window.txtFileNameX64.Text
            
                If ($file -like "*.exe")
                {
                    $syncHash.Window.txtSwitches_x64.Text = $defaultSwitchExe
                }
                ElseIf ($file -like "*.msi")
                {
                    $syncHash.Window.txtSwitches_x64.Text = $defaultSwitchMsi
                }
                ElseIf ($file -like "*.msu")
                {
                    $syncHash.Window.txtSwitches_x64.Text = $defaultSwitchMsu
                }
            }
        }
        Else
        {
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup("You must provide file(s) for this deployment.",0,"Error",0x0)
        }
    }
)

$syncHash.Window.btnNext4.add_Click(
    {
        $syncHash.Window.canStep5.Visibility = 'Collapsed'
        $syncHash.Window.canStep6.Visibility = 'Visible'
        
        # Create empty array
        $script:array = @()
        
        # Add Deployment Settings to array
        $count = $($($syncHash.Window.txtFileName.text).Split('\')).count
        
        Add-DeployOption -key 'Install File' -value $($($syncHash.Window.txtFileName.text).Split('\'))[$count -1]
        Add-DeployOption -key 'Source Directory' -value $($syncHash.Window.txtFileName.Text).Replace($($($syncHash.Window.txtFileName.text).Split('\'))[$count -1],'')
        
        Add-DeployOption -key 'Install Switches' -value $syncHash.Window.txtSwitches.Text
        
        If ($syncHash.Window.checkBox.IsChecked)
        {
            $count = $($($syncHash.Window.txtFileNameX64.text).Split('\')).count
        
            Add-DeployOption -key 'Install File (x64)' -value $($($syncHash.Window.txtFileNameX64.text).Split('\'))[$count -1]
            Add-DeployOption -key 'Source Directory (x64)' -value $($syncHash.Window.txtFileNameX64.Text).Replace($($($syncHash.Window.txtFileNameX64.text).Split('\'))[$count -1],'')
            
            Add-DeployOption -key 'Install Switches (x64)' -value $syncHash.Window.txtSwitches.Text
        }
                
        Add-DeployOption -key 'Credential' -value $syncHash.Window.comboBox.Text
        
        $syncHash.Window.datDeploymentSetting.ItemsSource = $script:array
    }
)

Function Test-Install {
    Param (
        $command, 
        $remoteinstall,
        $destpath, 
        $sourcepath,
        $PackageName,
        $syncHash
    )
	
    New-Item -Path $destpath -Type Directory -Force
    Copy-Item  -Path $sourcepath -Destination $destpath -Recurse -Force
    
    If ($syncHash.Window.checkBox.IsChecked)
    { 
        If ($(Get-WmiObject -Class Win32_OperatingSystem -Property OsArchitecture).OsArchitecture -eq '64-bit')
        {
            $installpath = "$destpath\$PackageName\Install-64.bat"
        }
        ElseIf ($(Get-WmiObject -Class Win32_OperatingSystem -Property OsArchitecture).OsArchitecture -eq '32-bit')
        {
            $installpath = "$destpath\$PackageName\Install-86.bat"
        }
    }
    Else
    {
        $installpath = "$destpath\$PackageName\Install-all.bat"
    }
    
    Try
    { 
        If ($syncHash.window.CurrentUser.IsSelected)
        {
            & $command $remoteinstall -h -accepteula $installpath
        }
        ElseIf ($syncHash.window.LocalSystem.IsSelected)
        { 
            & $command $remoteinstall -h -accepteula -s $installpath
        }
    }
    
    Catch
    {
        $_
    }
    
    Remove-Item -Path $destpath -Recurse -Force

    $syncHash.Counters.top++

    [int]$pct = ($syncHash.Counters.top/$syncHash.Counters.ComputerCount) * 100
    $syncHash.Window.ProgressBarTop.Value = $pct
    
    $syncHash.Window.lblProgressTop.Content = "Running install process - $($syncHash.Counters.top) of $($syncHash.Counters.ComputerCount) complete."
    
    $syncHash.Counters.bottom++

    [int]$pct = ($syncHash.Counters.bottom/$syncHash.Counters.AllCount) * 100
    $syncHash.Window.ProgressBarBottom.Value = $pct
    
    If ($pct -eq 100)
    {
        $syncHash.Window.btnClean.IsEnabled = $true
    }
}

$syncHash.Window.btnDeploy.add_Click(
    {
        # Set Counters
        $syncHash.Counters = @{}
        $syncHash.Counters.ComputerCount = $syncHash.Window.listComputers.Items.Count
        $syncHash.Counters.FileCount     = $(If ($syncHash.Window.checkBox.IsChecked) { 2 } Else { 1 })
        $syncHash.Counters.AllCount      = $syncHash.Counters.ComputerCount + ($syncHash.Counters.FileCount * 2)
        
        $syncHash.Counters.top           = 0
        $syncHash.Counters.bottom        = 0
        
        $syncHash.Counters
        
        # Set Progress Bar Labels
        $syncHash.Window.lblProgressTop.Content    = "Building Package"
        $syncHash.Window.lblProgressBottom.Content = 'Step 1 of 2 - Buliding Deployment Package'
        
        # Show current canvas
        $syncHash.Window.canStep6.Visibility = 'Collapsed'
        $syncHash.Window.canStep7.Visibility = 'Visible'
        
        # Create temp folder for deployment package
        If (-not(Test-Path -Path "$currentdir\deploy"))
        {
            New-Item -Path $currentdir -Name 'deploy' -ItemType Directory
        }
    
        $PackageName = $(Get-Date -Format 'yyyyMMdd-HHmm')
        New-Item -Path "$currentdir\deploy" -Name $PackageName -ItemType Directory
        
        # Build deployment package
        If ($syncHash.Window.checkBox.IsChecked)
        {
            $count = $($($syncHash.Window.txtFileName.text).Split('\')).count
            Add-Package -Filename $($($syncHash.Window.txtFileName.text).Split('\'))[$count -1] -FilePath $($syncHash.Window.txtFileName.Text).Replace($($($syncHash.Window.txtFileName.text).Split('\'))[$count -1],'') -PackageName $PackageName -Type 'x86'
            
            $count = $($($syncHash.Window.txtFileNameX64.text).Split('\')).count
            Add-Package -Filename $($($syncHash.Window.txtFileNameX64.text).Split('\'))[$count -1] -FilePath $($syncHash.Window.txtFileNameX64.Text).Replace($($($syncHash.Window.txtFileNameX64.text).Split('\'))[$count -1],'') -PackageName $PackageName -Type 'x64'
        }
        Else
        {
            $count = $($($syncHash.Window.txtFileName.text).Split('\')).count
            Add-Package -Filename $($($syncHash.Window.txtFileName.text).Split('\'))[$count -1] -FilePath $($syncHash.Window.txtFileName.Text).Replace($($($syncHash.Window.txtFileName.text).Split('\'))[$count -1],'') -PackageName $PackageName -Type 'all'
        }
        
        Start-Sleep -Seconds 1
        
        # Reset counter
        $syncHash.Counters.top = 0
        
        # Reset progress bar
        $syncHash.Window.ProgressBarTop.Value = '0'
        
        # Set Progress Bar Labels
        $syncHash.Window.lblProgressTop.Content    = "Running install process - $($syncHash.Counters.top) of $($syncHash.Counters.ComputerCount) complete."
        $syncHash.Window.lblProgressBottom.Content = 'Step 2 of 2 - Deployment'
        
        # Create Runspace Pool
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $sessionstate.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry('HashTable',$syncHash,$null)))
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $threads, $sessionstate, $Host)

        $runspacepool.Open()
        
        Write-Host 'Runspace Open'
        
        $ScriptBlock = {
            Param (
                $command, 
                $remoteinstall,
                $destpath, 
                $sourcepath,
                $PackageName
            )
			
            If (-not(Test-Path -Path $destpath))
            {
                New-Item -Path $destpath -Type Directory -Force -ErrorAction Stop
            }
            
            Copy-Item  -Path $sourcepath -Destination $destpath -Recurse -Force -ErrorAction Stop
            
            If ($syncHash.Window.checkBox.IsChecked)
            { 
                If ($(Get-WmiObject -Class Win32_OperatingSystem -Property OsArchitecture).OsArchitecture -eq '64-bit')
                {
                    $installpath = "$destpath\$PackageName\Install-64.bat"
                }
                ElseIf ($(Get-WmiObject -Class Win32_OperatingSystem -Property OsArchitecture).OsArchitecture -eq '32-bit')
                {
                    $installpath = "$destpath\$PackageName\Install-86.bat"
                }
            }
            Else
            {
                $installpath = "$destpath\$PackageName\Install-all.bat"
            }
            
            Try
            { 
                If ($syncHash.window.CurrentUser.IsSelected)
                {
                    & $command $remoteinstall -h -accepteula $installpath
                }
                ElseIf ($syncHash.window.LocalSystem.IsSelected)
                { 
                    & $command $remoteinstall -h -accepteula -s $installpath
                }
            }
            
            Catch
            {
                $_
            }
            
            Remove-Item -Path $destpath -Recurse -Force

            $syncHash.Counters.top++
    
            [int]$pct = ($syncHash.Counters.top/$syncHash.Counters.ComputerCount) * 100
            $syncHash.Window.ProgressBarTop.Value = $pct
            
            $syncHash.Window.lblProgressTop.Content = "Running install process - $($syncHash.Counters.top) of $($syncHash.Counters.ComputerCount) complete."
            
            $syncHash.Counters.bottom++
    
            [int]$pct = ($syncHash.Counters.bottom/$syncHash.Counters.AllCount) * 100
            $syncHash.Window.ProgressBarBottom.Value = $pct
            
            If ($pct -eq 100)
            {
                $syncHash.Window.btnClean.IsEnabled = $true
            }
        }

        $command           = "$currentdir\app\PsExec.exe"
        $destpath          = 'C:\TempInstall'
        $script:sourcepath = "$currentdir\deploy\$PackageName"
        
        $($syncHash.Window.listComputers.Items)
        
        Write-Host 'testing'
        
        Foreach ($Computer in $($syncHash.Window.listComputers.Items))
        { 
            Write-Host 'starting loop'
            $remoteinstall = "\\$Computer"         
            Write-Host $remoteinstall

            #Test-Install -command $command -remoteinstall $remoteinstall -destpath $destpath -sourcepath $script:sourcepath -PackageName $PackageName -synchash $synchash
            
            # Create the powershell instance and supply the scriptblock and parameters
            $powershell = [Management.Automation.PowerShell]::Create().AddScript($ScriptBlock).AddArgument($command).AddArgument($remoteinstall).AddArgument($destpath).AddArgument($script:sourcepath).AddArgument($PackageName).AddArgument($syncHash)
            
            # Add the runspace to the PowerShell instance
            $powershell.RunspacePool = $runspacepool
            
            # Create a temporay collection for each runspace
            $temp = "" | Select-Object PowerShell,Runspace,Computer
            $temp.Computer = $Computer
            $temp.PowerShell = $powershell
            
            # Save the handle output when calling BeginInvoke(),this will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            
            $runspaces.Add($temp) | Out-Null
            
            Write-Host 'loop '
        }  
    }
)

$syncHash.Window.btnClean.add_Click(
    {
        #Get-RunspaceData
        
        #$runspacepool.close()
        
        Remove-Item -Path $script:sourcepath -Recurse -Force -ErrorAction Continue
        
        # Show current canvas
        $syncHash.Window.canStep7.Visibility = 'Collapsed'
        $syncHash.Window.canStep8.Visibility = 'Visible'
    }
)

$syncHash.Window.btnNew.add_Click(
    {
        # Show current canvas
        $syncHash.Window.canStep8.Visibility = 'Collapsed'
        $syncHash.Window.canStart.Visibility = 'Visible'
    }
)

$syncHash.Window.btnDone.add_Click(
    {
        $syncHash.Window.Close()
    }
)

$syncHash.Window.btnBack6.add_Click(
    {
        $syncHash.Window.canStep6.Visibility = 'Collapsed'
        $syncHash.Window.canStep5.Visibility = 'Visible'
    }
)

$syncHash.Window.btnBack5.add_Click(
    {
        $syncHash.Window.canStep5.Visibility = 'Collapsed'
        $syncHash.Window.canStep4.Visibility = 'Visible'
    }
)

$syncHash.Window.btnBack4.add_Click(
    {
        $syncHash.Window.canStep4.Visibility = 'Collapsed'
        $syncHash.Window.canStep3.Visibility = 'Visible'        
    }
)

$syncHash.Window.btnBack3.add_Click(
    {
        $syncHash.Window.canStep3.Visibility = 'Collapsed'
        
        If ($syncHash.Window.radioImport.IsChecked -eq $true)
        {
            If ($syncHash.Window.txtFile.Text -like "*.csv")
            { 
                # Import CSV file
                $rows = Import-Csv $syncHash.Window.txtFile.Text
		
                # Remove Computers from List
                foreach ($row in $rows) 
                {
                    $syncHash.Window.listComputers.Items.Remove($row.ComputerName) | Out-Null
                }
            }
            ElseIf ($syncHash.Window.txtFile.Text -like "*.txt")
            { 
                # Import text file
                $list = Get-Content -Path $syncHash.Window.txtFile.Text

                # Add Computers to List
                foreach ($item in $list) 
                {
                    $syncHash.Window.listComputers.Items.Remove($item) | Out-Null
                }
            }

            $syncHash.Window.txtFile.Clear()
            $syncHash.Window.btnImport.IsEnabled = $false
            $syncHash.Window.btnNext2.IsEnabled = $false
            $syncHash.Window.canStep2.Visibility = 'Visible' 
        }
        Else
        {
            $syncHash.Window.canStep1.Visibility = 'Visible'
        }      
    }
)

$syncHash.Window.btnBack2.add_Click(
    {
        $syncHash.Window.canStep2.Visibility = 'Collapsed'
        $syncHash.Window.canStep1.Visibility = 'Visible'        
    }
)

$syncHash.Window.btnBack.add_Click(
    {
        $syncHash.Window.canStep1.Visibility = 'Collapsed'
        $syncHash.Window.canStart.Visibility = 'Visible'        
    }
)

$syncHash.Window.btnOpenFile.add_Click(
    {
        # Open load file dialog window
        $fd = New-Object system.windows.forms.openfiledialog
        #$fd.InitialDirectory =  $currentdir + "\groups"
        $fd.Filter = "All Files|*.*|Comma Separated Values file|*.csv|Text|*.txt"
        $fd.MultiSelect = $false
        $fd.showdialog()
		
        If ($fs.filename -ne '')
        { 
            # Set text box
            $syncHash.Window.txtFile.Text = $fd.filename

            # Enable Import button
            $syncHash.Window.btnImport.IsEnabled = $true
        }
    }
)

$syncHash.Window.btnImport.add_Click(
    {
        $syncHash.Window.btnNext2.IsEnabled = $true
    }
)

$syncHash.Window.btnAddComputer.add_Click(
    {
        Add-ComputerName -ComputerName $syncHash.Window.txtComputerName.Text
        $syncHash.Window.btnNext2.IsEnabled = $true
    }
)

$syncHash.Window.txtComputerName.add_KeyDown(
    {
        If ($args[1].key -eq 'Enter') 
        {
            Add-ComputerName -ComputerName $syncHash.Window.txtComputerName.Text
            $syncHash.Window.btnNext2.IsEnabled = $true
        }
    }
)

$syncHash.Window.btnRemoveComputer.add_Click(
    {
        $Computer = $syncHash.Window.listComputers.SelectedItem
        $syncHash.Window.listComputers.Items.Remove($Computer) | Out-Null
        
        If ($($syncHash.Window.listComputers.Items).Count -eq 0)
        {
            $syncHash.Window.btnNext2.IsEnabled = $false
        }
    }
)

$syncHash.Window.listComputers.add_KeyDown(
    {
        If ($args[1].key -eq 'Delete')
        {
            $Computer = $syncHash.Window.listComputers.SelectedItem
            $syncHash.Window.listComputers.Items.Remove($Computer) | Out-Null
        
            If ($($syncHash.Window.listComputers.Items).Count -eq 0)
            {
                $syncHash.Window.btnNext2.IsEnabled = $false
            }
        }
    }
)

$syncHash.Window.btnAddFile.add_Click(
    {
        # Open load file dialog window
        $fd = New-Object system.windows.forms.openfiledialog
        #$fd.InitialDirectory =  $currentdir + "\groups"
        $fd.Filter = "All Files|*.*"
        $fd.MultiSelect = $false
        $fd.showdialog()
		
        If ($fs.filename -ne '')
        { 
            # Set text box
            $syncHash.Window.txtFileName.Text = $fd.filename
        }
    }
)

$syncHash.Window.btnAddFileX64.add_Click(
    {
        # Open load file dialog window
        $fd = New-Object system.windows.forms.openfiledialog
        #$fd.InitialDirectory =  $currentdir + "\groups"
        $fd.Filter = "All Files|*.*"
        $fd.MultiSelect = $false
        $fd.showdialog()
		
        If ($fs.filename -ne '')
        { 
            # Set text box
            $syncHash.Window.txtFileNameX64.Text = $fd.filename
        }
    }
)

$syncHash.Window.checkBox.add_Checked(
    {
        $syncHash.Window.txtFileNameX64.IsEnabled = $true
        $syncHash.Window.btnAddFileX64.IsEnabled = $true
    }
)

$syncHash.Window.checkBox.add_Unchecked(
    {
        $syncHash.Window.txtFileNameX64.IsEnabled = $false
        $syncHash.Window.btnAddFileX64.IsEnabled = $false
    }
)

$syncHash.Window.menuExit.add_Click(
    {
        $syncHash.Window.Close()        
    }
)

$syncHash.Window.ShowDialog() | Out-Null