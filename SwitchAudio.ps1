    #region [Title & Description]
    
     
    # Title: SwitchAudio GUI Project
    # Date Created: 12/10/2020
    # Date Last Updated: 01/01/2021
    # Version: 1.4
    # Description: Change audio devices with the click of a button!
    #
    # Requirements:
    # - Powershell 4.0 or greater
    # - frgnca's module AudioDeviceCmdlets (git: https://github.com/frgnca/AudioDeviceCmdlets)
    # 
    # Change Log:
    # - 1.0 
    # - 1.1 Added check for pre-requisite modules; prompt to install if not found.
    # - 1.2 Added check for admin priveleges; script will close and try to open with as an administrator, prompting UAC
    # - 1.3 Refined code; cleaned up comments; properly regioned code
    # - 1.4 Added [Table of Contents] region
    # ----- Created seperate group boxes based on object property
    # ----- Wrapped group boxes in tablelayoutpanel
    # ----- Split [Build Dynamic Form] into [Generate Button Objects], [Generate Form Objects], and [Define Button Action]
    # ----- Consolidated form formatting into [Form Properties] region for easy customization
    # ----- Generalized code; Added functionality to use different objects in place of AudioDeviceCmdlets
    
    
    #endregion
    #region [Table of Contents]
    
    #
    # *If you want to make changes, edit the [Form Attributes] region.
    # **If you want to re-purpose, edit [Form Attributes], [Build Object Array], [Define Button Action], and any associated object properties in the [Generate Form Objects] and [Generate Button Objects] regions
    # 
    # [Title & Description] - Title and Description
    # [Startup Preflight] - check and initialize required libraries/modules; can remove/adjust if necessary
    # *[Form Properties] - form formatting; can adjust any variable in here, changes here are purely aesthetic
    # **[Build Object Array] - generate an array list with a specific object type; can change object type, may require adjustments to object property use elsewhere
    # **[Define Button Action] - create script block that is applied to each button; adjust this to change what the button clicks do
    # **[Generate Button Objects] - generates buttons based on the defined object array
    # **[Generate Form Objects] - put together the Windows Form, generate groups based on object type; only adjust this if you know what you're doing
    # [Start Form] - add controls to form; start form; only adjust this if you know what you're doing
    
    
    #endregion
    #region [Startup Preflight]
    
    
    #Intializations
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Collections") 
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $wshell = New-Object -ComObject ("Wscript.Shell")
    
    #Check for AudioDeviceCmdlets Powershell Module
    #prompt to install if not found
    if(Get-Module -ListAvailable -Name AudioDeviceCmdlets){
    	Import-Module -Name AudioDeviceCmdLets
    }
    else{
    	#Check if running as admin
    	if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    		Start-Process PowerShell -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
    		exit;
    	}
    	
    	Start-Sleep 2
    	
    	$installModule = $wshell.Popup("This machine is missing a required Powershell module: AudioDeviceCmdlets.`n`nDownload and install?", 0, "Missing Required Module!", 0x1)
    	if($installModule -eq 1){
    		Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    		Install-Module -Name AudioDeviceCmdlets
    	}
    	else{
    		return
    	}
    }
    
    
    #endregion 
    #region [Form Properties]
    
    
    #form properties
    $formTitle = "Switch Audio"
    $formBackColor = "#2a2a72"
    $buttonMaxColumns = 3
    
    #groupbox properties
    #title of each group box is set to the type property of the object
    $groupBoxFontSize = 10
    $groupBoxFontType = 'Microsoft Sans Serif'
    $groupBoxFontStyle = [System.Drawing.FontStyle]::Bold
    $groupBoxFontColor = '#FFFFFF'
    
    #button properties
    $buttonFontSize = 9
    $buttonFontType = 'Microsoft Sans Serif'
    $buttonFontStyle = [System.Drawing.FontStyle]::Bold
    $buttonFontColor = '#FFFFFF'
    $buttonWidth = 100
    $buttonHeight = 70
    
    
    #endregion
    #region [Build Object Array]
    
    
    #Get list of all audio devices
    #Create array of button objects
    $objectList = Get-AudioDevice -List | Sort-Object -Property Type
    [System.Windows.Forms.Button []]$buttonList = [System.Collections.ArrayList]@()
    $selection = @()
    #$objectList | Select-Object -Property Index, Name | Sort-Object -Property Index | Out-GridView -Title "Select data file to reconcile." -Passthru | ForEach-Object {$selection += $_.Name}
    ForEach($device in $selection){
    	if ($device -eq "Microphone (Realtek High Definition Audio)"){
    	echo "Found default microphone!"
    	}
    }
    
    
    #endregion
    #region [Define Button Action]
    
    
    #define the script block that will activate on each button click. 
    #$this is a placeholder object from the current index of the object array
    #parameters are limited to the suitable object properties
    [scriptblock]$buttonAction = {
    	#create popup window to confirm choice
    	$confirmChoice = $wshell.Popup("Change audio device to " + $this.Text + '?', 0, "Confirm Choice", 0x1)
    	#only execute if choice is OK
    	if ($confirmChoice -eq 1)
    	{
    		#change audio device to the device that matches the button's name property
    		Set-AudioDevice -Index $this.Tag
    	}
    }
    
    
    #endregion
    #region [Generate Button Objects]
    
    
    #parse audio device array
    #dynamically generate buttons
    #add click to change default input/output to selection
    #add buttons to button object list
    
    foreach ($object in $objectList)
    {
    	#newbutton - Forms.Button Object
    	#newbutton - use Forms.Button.Name property to pass audio device index to button click
    	#newbutton - set properties in [Form Properties] 
    	$newButton = New-Object System.Windows.Forms.Button
    	$newButtonName = ($object | Select-Object -ExpandProperty Name -Property BaseName)
    	$newButtonType = ($object | Select-Object -Property Type -ExpandProperty Type)
    	if($newButtonName.Length -ge 37){
    		$newButtonName = $newButtonName.SubString(0, 33) + "..."
    	}
    	$newButton.Text = $newButtonName
    	$newButton.Name = $newButtonType
    	$newButton.Tag = $object.Index	
    	$toolTip = New-Object System.Windows.Forms.ToolTip
    	$toolTip.SetToolTip($newButton, ($object | Select-Object -ExpandProperty Name -Property BaseName))
    	$newButton.Font = [System.Drawing.Font]::new($buttonFontType, $buttonFontSize, $buttonFontStyle)
    	$newButton.ForeColor = $buttonFontColor
    	$newButton.BackColor = "#7a89c2"
    	$newButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    	$newButton.FlatStyle = 1
    	$newButton.Padding = '5,5,5,5'
    	$newButton.Add_Click($buttonAction)
    	
    	#add new button object to the button object list
    	$buttonList += $newButton
    }
    
    
    #endregion
    #region [Generate Form Objects]
    
    
    #define groupbox list
    #define tablelayoutpanel list
    [System.Windows.Forms.GroupBox []]$groupList = [System.Collections.ArrayList]@()
    [System.Windows.Forms.TableLayoutPanel []]$tableList = [System.Collections.ArrayList]@()
    
    #dynamically generate button grid layout and grouping based on object property: type
    foreach ($type in ($objectList | Select-Object -Property Type -ExpandProperty Type -Unique)){
    		
    	#groupbox - Forms.GroupBox Object
    	#groupbox - set properties in [Form Properties] 
    	#groupbox - dynamically generates group boxes based on type property
    	$groupBox = New-Object System.Windows.Forms.GroupBox
    	$groupBox.AutoSize = $true
    	$groupBox.AutoSizeMode = 'GrowAndShrink'
    	$groupBox.Padding = '10,10,10,10'
    	$groupBox.Anchor = 'None'
    	$groupBox.Dock = 'Fill'
    	$groupBox.Text = $type
    	$groupBox.Tag = $type
    	$groupBox.Font = [System.Drawing.Font]::new($groupBoxFontType, $groupBoxFontSize, $groupBoxFontStyle)
    	$groupBox.ForeColor = $groupBoxFontColor
    	
    	#tablepanel - Forms.TabelLayoutPanel Object
    	#tablepanel - set properties in [Form Properties] 
    	#tablepanel - dynamically generates tablepanellayouts based on type property
    	$tablePanel = New-Object System.Windows.Forms.TableLayoutPanel
    	$tablePanel.AutoSize = $true
    	$tablePanel.AutoSizeMode = 'GrowAndShrink'
    	$tablePanel.Margin = 20
    	$tablePanel.Dock = 'Fill'
    	$tablePanel.Tag = $type
    	$tablePanel.RowCount = ([math]::ceiling($objectList.Length / $buttonMaxColumns))
    	$tablePanel.ColumnCount = $buttonMaxColumns
    	
    	#add buttons to tablepanel by type
    	#add tablepanel to groupbox by type
    	$tablePanel.Controls.AddRange(($buttonList | Where-Object -Property Name -EQ -Value $type))
    	$groupBox.Controls.Add($tablePanel)
    	
    	#add new groupbox object to the groupbox object list
    	$groupList += $groupBox
    }
    
    #formpanel - Forms.TabelLayoutPanel Object
    #formpanel - declare tablelayoutpanel to place group boxes in
    $formPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $formPanel.AutoSize = $true
    $formPanel.AutoSizeMode = 'GrowAndShrink'
    $formPanel.Margin = 20
    $formPanel.Dock = 'Fill'
    $formPanel.Tag = 'FormTablePanel'
    $formPanel.RowCount = $groupList.Length
    $formPanel.ColumnCount = 1
    
    #form - Forms.Form Object
    #form - height and width sized to accomadate table sizing; disabled resize; disabled maximize button
    #form - start center screen
    #form - set other properties in [Form Properties] 
    $form = New-Object System.Windows.Forms.Form
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = '5,5,5,5'
    $form.FormBorderStyle = 'FixedDialog'
    $form.StartPosition = 'CenterScreen'
    $form.MaximizeBox = $false
    $form.BackColor = $formBackColor
    $form.text = $formTitle
    
    #icon - System.Drawing.Icon Object
    #icon - declares Icon object as local image; only works when running in Powershell, will not work when compiled into an executable.
    #$Icon = New-Object system.drawing.icon ("C:\Users\*USER*\Documents\icon.ico")
    #$form.Icon = $Icon
    
    
    #endregion
    #region [Start Form]
    
    
    #add all the group boxes to the form
    $formPanel.Controls.AddRange($groupList)
    $form.Controls.Add($formPanel)
    
    #Show the form
    [void]$form.ShowDialog()
    
    
    #endregion    
