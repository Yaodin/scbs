####---------------------------------------------------------------------------------------------####
# Service Center Backup Script (SCBS)																#
# Written by: Austin Heyne																			#
# 																									#
# Program provides an interface to create HW Tickets in Service-Now and perform a managed backup	#
# to a network drive (Norfile). 																	#
# 																									#
# --Information--																					#
# User guide is located in pbworks.																	#
# This program maps norfile, interfaces with the Service-Now API, collects system information		#
# and performs a managed, multithreaded backup using 7z and an internal job handeling funcion. 		#
# This allows backup options to be easily controled by the end users and improves performance in 	#
# certain sitiuations. API interface streamlines ticket creation and autopopulates computer 		#
# information in the ticket. 																		#
####---------------------------------------------------------------------------------------------####

function GenerateForm {

#--Program States--#
#
# Program state is changed on button press. Available states below.
# Make all state changes through the setGuiState function. eg setGuiState "StartNew"
#
#-Name------State-----------------------------------Gui-----------------------------Buttons---------#
# Splash	| Choose to make new ticket or update	| Two selections				| New	 Update #
# StartNew	| Initial state. Next is Select			| Caller and User input fields	|          Next #
# StartExist| Initial state. Next is Select			| HW Number input field			|          Next #
# Select	| Choose Files to backup				| Directory Tree View			| Back     Next #
# Confirm	| Confirm Files to backup				| RTB display of files			| Back   Backup #
# Backup	| Backing up files and verify 			| RTB display of status			|        Cancel #
# Done		| Backup and verification done			| RTB display of status			|          Exit #
# Error		| Error occured, back returns to select	| RTB display of status			| Back     Exit #
#-------------------------------------------------------------------------------------B2---------B1-#

#region Imported Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion Imported Assemblies

#region GUI Elements
$SCBS = New-Object System.Windows.Forms.Form
$B_Advanced = New-Object System.Windows.Forms.Button
$CB_Monolithic = New-Object System.Windows.Forms.CheckBox
$CB_Store = New-Object System.Windows.Forms.CheckBox
$PB_splash = New-Object System.Windows.Forms.PictureBox
$PB_StartPic = New-Object System.Windows.Forms.PictureBox
$TB_HWNumber = New-Object System.Windows.Forms.TextBox
$label_HWNumber = New-Object System.Windows.Forms.Label
$CoB_Make = New-Object System.Windows.Forms.ComboBox
$label_Make = New-Object System.Windows.Forms.Label
$label_SN = New-Object System.Windows.Forms.Label
$TB_SN = New-Object System.Windows.Forms.TextBox
$TB_Description = New-Object System.Windows.Forms.TextBox
$label_Description = New-Object System.Windows.Forms.Label
$label_PartnershipGroup = New-Object System.Windows.Forms.Label
$label_userPhone = New-Object System.Windows.Forms.Label
$TB_UserPhone = New-Object System.Windows.Forms.TextBox
$TB_CallerPhone = New-Object System.Windows.Forms.TextBox
$label_callerPhone = New-Object System.Windows.Forms.Label
$CoB_PartnershipGroup = New-Object System.Windows.Forms.ComboBox
$CoB_Warranty = New-Object System.Windows.Forms.ComboBox
$label_Warranty = New-Object System.Windows.Forms.Label
$RTB_Output = New-Object System.Windows.Forms.RichTextBox
$RTB_Comments = New-Object System.Windows.Forms.RichTextBox
$label_Comments = New-Object System.Windows.Forms.Label
$TB_Dropoff = New-Object System.Windows.Forms.TextBox
$label_Dropoff = New-Object System.Windows.Forms.Label
$TB_Pickup = New-Object System.Windows.Forms.TextBox
$label_Pickup = New-Object System.Windows.Forms.Label
$TB_User = New-Object System.Windows.Forms.TextBox
$label_User = New-Object System.Windows.Forms.Label
$TB_Caller = New-Object System.Windows.Forms.TextBox
$label_Caller = New-Object System.Windows.Forms.Label
$B_Right = New-Object System.Windows.Forms.Button
$B_Left = New-Object System.Windows.Forms.Button
$B_UpdateTicket = New-Object System.Windows.Forms.Button
$B_NewTicket = New-Object System.Windows.Forms.Button
$statusBar = New-Object System.Windows.Forms.StatusBar
$treeView = New-Object System.Windows.Forms.TreeView
$timer_backupControl = New-Object System.Windows.Forms.Timer
$timer_norfileConnect = New-Object System.Windows.Forms.Timer
$toolTip1 = New-Object System.Windows.Forms.ToolTip
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion GUI Elements

#region GUI Handlers
Write-Verbose "Define Handlers"

#--Splash Screen Controls--#
Write-Verbose "Define Button Update Ticket"
$B_UpdateTicket_OnClick={
	# HW Ticket already exists, update ticket
	setGuiState "StartExist"
	$script:StartState = "Exist"	#Varible for determining which state the program started in.
}

Write-Verbose "Define Button New Ticket"
$B_NewTicket_OnClick={
	# Create new hardware ticket
	setGuiState "StartNew"
	$script:StartState = "New"		#Varible for determining which state the program started in.
}

#--Main Program Controls--#
Write-Verbose "Define Button 1"
$B_Right_OnClick={
	Write-Verbose "Button 1 pressed in program state: $script:ProgramState"
	switch($script:ProgramState){

	# This is switching on the current program state so the code block that is run is determined by the
	# current program state, not the state the program is going into. Program state is thus changed at
	# the end of the code block to reduce confusion.
		
#---Start---#
		"StartNew" {
			# Gather basic data
			$model = gwmi -Class Win32_ComputerSystem | Select-Object Model
			
			[Hashtable]$data = @{
				"caller4x4" = $TB_Caller.Text.Trim(" ");
				"callerPhone" = $TB_CallerPhone.Text.Trim(" ").Replace("-","").Replace(".","").Replace(",","");
				"user4x4" = $TB_User.Text.Trim(" ");
				"userPhone" = $TB_UserPhone.Text.Trim(" ").Replace("-","").Replace(".","").Replace(",","");
				"warranty" = $CoB_Warranty.Text;
				"make" = $CoB_Make.Text;
				"model" = $model.Model;
				"chassis" = getChassis;
				"os" = getOS;
				"username" = (whoami).split("\")[1];
				"assignmentgroup" = $script:assignmentGroup;
				"partnership" = $CoB_PartnershipGroup.Text;
				"serial" = $TB_SN.Text.Trim(" ");
				"pickup" = $TB_Pickup.Text.Trim(" ");
				"dropoff" = $TB_Dropoff.Text.Trim(" ");
				"description" = $TB_Description.Text.Trim(" ");
				"comments" = $RTB_Comments.Text;
			}
			
			validateData $data
			
			if($script:validData){
				if(!$script:NewCaseCreated){
					# Create Ticket and Get Returned HW Number, Caller and Affected User
					snQuery -new -data $data 					
				} else {
					# Ticket has already been created, update it instead.
					snQuery -update -data $data -sysID $script:SysID
				}
				
				[System.Xml.XmlDocument]$xml = $script:lastWebRequestResult
				$result = $xml.response.result
				
				if($result.opened_by.display_value -match $data.caller4x4 -and $result.u_customer.display_value -match $data.user4x4){
					# Case created correctly
					setupStartEnv			
					setGuiState "Select"
				} else {
					$msg = ""
					if($result.opened_by.display_value -notmatch $data.caller4x4){
						$msg += "The caller 4x4 is not valid.`n"
					}
					if($result.u_customer.display_value -notmatch $data.user4x4){
						$msg += "The user 4x4 is not valid."
					}
					popupError "Invalid Input" $msg
				}
			}
		}
		
#--StartExist--#
		"StartExist" {
			$script:hwNumber = $TB_HWNumber.Text.Trim(" ").ToUpper()
			
			if(!$script:hwNumber.StartsWith("HW")){
				$script:hwNumber = "HW" + $script:hwNumber
			}
			
			if($script:hwNumber -match $script:regex_hw){
				$TB_HWNumber.Text = $script:hwNumber
				# Valid hwNumber format, querySN
				snQuery -verify -hwNumber $script:hwNumber
				
				if($script:lastWebRequest){
					# Valid hwNumber
					setupStartEnv
					setGuiState "Select"
				}
			} else {
				popupError "Invalid Hardware Number" "The provided hardware number is not valid."
			}				
		}
		
#---Select---#
		"Select" {
			# Gather selections
			$script:CheckedNodes = New-Object System.Collections.ArrayList
			getCheckedNodes $treeView.Nodes $script:CheckedNodes
			
			# Display selections on RTB
			updateRTB
			setGuiState "Confirm"
		}	
		
#---Confirm---#		
		"Confirm" {		
			# Removes all files already in the folder. Needed if a backup fails and is being restarted.
			Remove-Item $($script:netPath + "*.*")
			
		#---Create Job Queue---#	
			$script:jobQueue = makeJobQueue 
			
		#---Create Jobs---#
			# Create 'Report Building' Job
			# This job uses msinfo32.exe /report to dump system data to a file and then 
			#	upload it to the backup directory.
			# This job has a different Backup method so we cannot use the makeJob function
			$jobObject = New-Object PSObject -Property $jobObjectProperties
			$jobObject.Name = "Collect System Information"
			$jobObject.Path = "Collect System Information"
			Add-Member -InputObject $jobObject -Name Backup -MemberType ScriptMethod -Value {
				# Run the job
				$this.Job = Start-Job -ArgumentList $script:netPath -ScriptBlock {
					param($destination)
					
					# Make temp folder for compInfo
					$tempDir = Join-Path $Env:APPDATA SCBS
					New-Item -ItemType Directory -Path $tempDir -ErrorAction SilentlyContinue
					
					$filename = Join-Path $tempDir "compInfo.txt"
					# Remove if already exists
					Remove-Item $filename -ErrorAction SilentlyContinue
					
					# Gather and save info
					msinfo32.exe /report $filename
					
					# Wait for msinfo to gather the data and start saving it
					while(!(Test-Path $filename)){
						sleep 1
					}
					
					# Wait for msinof to finish saving
					do{
						$size0 = $(Get-Item $filename).Length
						sleep 1
						$size1 = $(Get-Item $filename).Length
					} until($size0 -eq $size1)
					
					# Move the compInfo file to the backup folder
					Move-Item -Path $filename -Destination $(Join-Path $destination "compInfo.txt") -Force
					
					# Cleanup
					Remove-Item $tempDir -Recurse
				} 
			}
			
			Add-Member -InputObject $jobObject -Name State -MemberType ScriptMethod -Value {
				#get the state of the job
				if($this.Job.State -eq $null){
					return "Unprocessed"
				} else {
					return $this.Job.State
				}
			}
			
			# Add the 'Report Building' Job to the jobQueue
			$script:jobQueue.JobList += $jobObject
			
			# Create jobs for each selected node
			foreach($Node in $script:CheckedNodes){
				Write-Debug $Node.Name
				Write-Debug $Node.Tag
				
				# Add them to the Job List
				$script:jobQueue.JobList += makeJob $Node.Name $Node.Tag
			}
			
			setGuiState "Backup"		
			
			# Start the timer loop to being checking jobs
			$timer_backupControl.Start()
		}	
		
#---Backup---#
		"Backup" {
			# Cancel button pressed
		
			# 6 = yes; 7 = no
			$msg = "Are you sure you want to cancel the backup?`nAll finised backups will remain for now but will be deleted if the backup is restarted!"
			$answer = popupError "Confirm Cancel" $msg 0x4
			if($answer -eq '6'){
				# We're sure we want to cancel
				Write-Warning "Backup canceled by user"
				$statusBar.Text = "Canceling..."
				$timer_backupControl.Stop()			
				foreach($_ in $script:jobQueue.JobList){
					# Kill all jobs
					Stop-Job $_.Job -ErrorAction SilentlyContinue
				}			
				$script:jobQueue.UpdateGUI()
				setGuiState "Error"	
			} else {
				return
			}			
		}	
		
#---Done---#
		"Done" {
			Write-Verbose "Done"
			$SCBS.Close()
		}
		
#---Error---#
		"Error" {
			Write-Verbose "Done"
			$SCBS.Close()
		}
	}
}

Write-Verbose "Define Button 2"
$B_Left_OnClick={ #back button, reverts program state depeding on current state
	Write-Verbose $("Button 2 pressed in program state: " + $script:ProgramState) 
	switch($script:ProgramState){
		"StartExist" {
			setGuiState "Splash"
		}
		"StartNew" {
			setGuiState "Splash"
		}
		"Select" {
			# Determine what Start State to go back to.
			if($script:StartState -eq "Exist"){
				setGuiState "StartExist"
			} elseif($script:StartState -eq "New"){
				setGuiState "StartNew" 
			}
		}
		"Confirm" {
			setGuiState "Select" 
		}
		"Error" {
			# Clearout old errors
			$script:ErrorList = @()
			
			updateRTB
			setGuiState "Confirm" 
		}
	}
}

Write-Verbose "Define Button Advanced"
$B_Advanced_OnClick={
	Write-Verbose "Advanced Button Pressed"
	$B_Advanced.Visible = $false
	$CB_Monolithic.Visible = $true
	$CB_Store.Visible = $true
}

#--Timer Controls--#
Write-Verbose "Define Backup Control Timer"
$timer_backupControl_Tick={
	#--Backup Control Timer--#
	# This timer controls the backup job queue. Most heavy lifting is done by the manageJobQueue 
	# function. Also regulatea the activity graphic.
	# 
	Write-Verbose "Timer 1 Tick"
	# This code runs every half second
	$timer_backupControl.Stop()	#stop the timer so that we don't encounter any race conditions with the next timer tick
	
	$script:tickcount++		# Tracks how many times the timer has ticked
	Write-Debug $("TickCount: " + $script:tickcount) 
	
	# Spinning bar graphic
	updateBarGraphic "Backing up" 
	
	# Used to disable cancel button for half a second to prevent accidental clicks 
	# (actual 1s with .5s timer)
	if(-not $B_Right.Enabled){
		Write-Verbose "Enable Button"
		sleep .5
		$B_Right.Enabled = $true
	}
	
	manageJobQueue 

	# Update Gui
	$script:jobQueue.UpdateGUI()
	
	# Backup state exit control
	if($script:ProgramState -eq "Backup"){
		Write-Verbose "Restarting Timer"
		$timer_backupControl.Start()
	}
}

Write-Verbose "Define Norfile Connect Timer"
$timer_norfileConnect_Tick={
	#--Norfile Connect Timer--#
	# Controls norfile maping job and post run functions.
	
	Write-Verbose "Timer 2 Tick"
	$timer_norfileConnect.Stop()	# Stop the timer so that we don't encounter any race conditions with the next timer tick
	
	updateBarGraphic "Mapping Norfile" 
	
	# List of status mesages for the job that results in an error for the script
	$badStatus = "AtBreakpoint","Blocked","Disconnected","Failed","Stopped","Stopping","Suspended","Suspending"
	
	# Grab the map state from the job object
	$mapState = $script:mapNorfileJob.State()
	if($mapState -eq "Unprocessed"){
		# Job hasn't started yet. 
		# Considered restarting but caused issues with maping the drive repeatedly...
		$timer_norfileConnect.Start()
		return
	} elseif($mapState -eq "Completed"){
		#  Map is done, test it 
		try{
			$script:netPath = Join-Path $script:DriveLetter "\"
		} catch {
			popupError "Drive mapping failure" "Failed to properly map norfile. Check network connection." 0x0
			$SCBS.Close()
			return
		}
		if(!(Test-Path $script:netPath)){
			# Test failed, queuery retry
			$answer = popupError "Drive mapping failure" "Failed to properly map norfile." 0x5
			if($answer -eq '4'){
				# Recreate the job and start it
				$script:mapNorfileJob = makeMapNorfileJob
				$script:mapNorfileJob.Map()
				$timer_norfileConnect.Start()
				return
			} elseif($answer -eq '2'){
				# Quit
				$SCBS.Close()
				return
			}
		} else {
			# Drive mapped successfully, verify version
			Write-Debug ("Drive mapped: " + $script:DriveLetter)
			verifyVersion
			if($script:ValidVersion){ 
				Write-Verbose "Version verified"
				$statusBar.Text = "Norfile Mapped. Ready to Continue"
				$B_Right.Enabled = $true
			} else {
				# Old version, forbid running
				Write-Error "Old version"
				$msg = "This version of SCBS is out of date. Please go to servicesapps.ou.edu/dml/software to download a new version. Would you like to download the new version now?"
				$answer = popupError "Out of Date Version" $msg 0x4
				if($answer -eq '6'){
					# Yes, download new verision
					start 'download url' # Redacted
					$SCBS.close()
				} else {
					# No, close, can't proceed with old version
					$SCBS.close()
				}
			}
		}
		return
	} elseif($mapState -eq "Running"){
		# Map is still running
		$timer_norfileConnect.Start()
		return
	} elseif($badStatus -contains $mapState){
		# Something bad happened while mapping, queuery retry
		$answer = popupError "Drive mapping failure" "Failed to properly map norfile." 0x5
		if($answer -eq '4'){
			# Recreate the job and start it
			$script:mapNorfileJob = makeMapNorfileJob
			$script:mapNorfileJob.Map()
			$timer_norfileConnect.Start()
			return
		} else {
			# Quit
			$SCBS.Close()
			return
		}
	}
}

#--Data Controls--#
Write-Verbose "Define CoB Warranty SelectedIndexChanged"
$handler_CoB_Warranty_SelectedIndexChanged={
	Write-Verbose $CoB_Warranty.SelectedItem
	switch($CoB_Warranty.SelectedItem){
		"Dell Warranty" {$CoB_Make.SelectedItem = "Dell"}
		"HP Warranty" {$CoB_Make.SelectedItem = "HP"}
		"Lenovo Warranty" {$CoB_Make.SelectedItem = "Lenovo"}
		default {$CoB_Make.SelectedIndex = '-1'}
	}
}

#--Form Controls--#
Write-Verbose "Define Load"
$FormEvent_Load={
	# Code runs on form load
	Write-Verbose "Form Event Load"
	setGuiState "Splash"
	
	# Make sure we can get to needed resources
	verifyConnection
	
	Write-Verbose "Make map norfile job"
	# Make 'mapNorfile' job and start it
	$script:mapNorfileJob = makeMapNorfileJob
	$script:mapNorfileJob.Map()
	
	# Start job timer control
	$timer_norfileConnect.Start()

	#Setup Tooltips (hover text)
	
	$text = "4x4 of IT Rep."
	$toolTip1.SetToolTip($label_Caller,$text)
	$toolTip1.SetToolTip($TB_Caller,$text)
	
	$text = "4x4 of the Hardware's Owner."
	$toolTip1.SetToolTip($label_User,$text)
	$toolTip1.SetToolTip($TB_User,$text)
	
	$text = "Phone Number of IT Rep."
	$toolTip1.SetToolTip($label_callerPhone,$text)
	$toolTip1.SetToolTip($TB_CallerPhone,$text)
	
	$text = "Phone Number of Hardware's Owner."
	$toolTip1.SetToolTip($label_userPhone,$text)
	$toolTip1.SetToolTip($TB_UserPhone,$text)
	
	$text = "Warranty of the device being checked in. If the device is no longer under warranty select 'Out of Warranty'."
	$toolTip1.SetToolTip($label_Warranty,$text)
	$toolTip1.SetToolTip($CoB_Warranty,$text)
	
	$text = "Make or manufacturer of device."
	$toolTip1.SetToolTip($label_Make,$text)
	$toolTip1.SetToolTip($CoB_Make,$text)
	
	$text = "Partnership Group of the Affected User. If none available select No Partnership."
	$toolTip1.SetToolTip($label_PartnershipGroup,$text)
	$toolTip1.SetToolTip($CoB_PartnershipGroup,$text)
	
	$text = "Service Tag or Serial Number of device. If non available provide any unique identifying number on device."
	$toolTip1.SetToolTip($label_SN,$text)
	$toolTip1.SetToolTip($TB_SN,$text)
	
	$text = "Location the device will be picked up."
	$toolTip1.SetToolTip($label_Pickup,$text)
	$toolTip1.SetToolTip($TB_Pickup,$text)
	
	$text = "Location the device will be returned to after completion of service."
	$toolTip1.SetToolTip($label_Dropoff,$text)
	$toolTip1.SetToolTip($TB_Dropoff,$text)
	
	$text = "Additional comments to be added to the incident in Service-Now."
	$toolTip1.SetToolTip($label_Comments,$text)
	$toolTip1.SetToolTip($RTB_Comments,$text)	
	
	$text = "Advanced backup settings."
	$toolTip1.SetToolTip($B_Advanced,$text)
	
	$text = "Place backup in a single archive. Greatly hinders performance."
	$toolTip1.SetToolTip($CB_Monolithic,$text)
	
	$text = "Store backup in archive rather than compressing it."
	$toolTip1.SetToolTip($CB_Store,$text)
	
	#Autofill data
	$TB_SN.Text = gwmi win32_bios | Select-Object -ExpandProperty SerialNumber
}

Write-Verbose "Define State Correction"
$OnLoadForm_StateCorrection={
	Write-Verbose "Form Event State Correction"
	#Correct the initial state of the form to prevent the .Net maximized form issue
	$SCBS.WindowState = $InitialFormWindowState
}

Write-Verbose "Define Resize"
$FormEvent_Resize={
	Write-Verbose "Form Event Resize"
	#Handles dynamic resizing of some gui elements.
	#since you can't anchor gui elements to other gui elements this calculates sizes and positions and sets them manually.

	[int]$WindowWidth = $SCBS.Width	 #Width of form
	$WindowWidth = $WindowWidth - 17 #this compensates for the difference between ClientSize and Window min Size
	[int]$Margin = '12'				 #Distance from edge of form to edge of items
	[int]$Spacer = '7'				 #Distance between pickup and dropoff TBs
	
	[int]$TBSize = ($WindowWidth - ($Margin * 2) - $Spacer)/2
	$TBSize = [Math]::Floor($TBSize)
	$TBSize = [Math]::Abs($TBSize)
	$DropoffPosition = ($Margin + $Spacer + $TBSize)
	
	# $TB_Pickup Size
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = $TBSize
	$TB_Pickup.Size = $System_Drawing_Size
	
	# $label_Dropoff Move
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $DropoffPosition
	$System_Drawing_Point.Y = $label_Dropoff_Draw_Y
	$label_Dropoff.Location = $System_Drawing_Point
	
	# $TB_Dropoff Size
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = $TBSize
	$TB_Dropoff.Size = $System_Drawing_Size
	
	# $TB_Dropoff Move
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $DropoffPosition
	$System_Drawing_Point.Y = $TB_Dropoff_Draw_Y
	$TB_Dropoff.Location = $System_Drawing_Point
	
	#--Splash Page--#
	$L_Centerline = ($WindowWidth - $Spacer) / 2
	$R_Centerline = $L_Centerline + $Spacer
	$B_NewTicket_Draw_X = [Math]::Abs([Math]::Floor($L_Centerline - $B_NewTicket_Size_W))
	$B_UpdateTicket_Draw_X = [Math]::Abs([Math]::Floor($R_Centerline))
	
	# $B_NewTicket
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $B_NewTicket_Draw_X
	$System_Drawing_Point.Y = $B_NewTicket_Draw_Y
	$B_NewTicket.Location = $System_Drawing_Point
		
	# $B_UpdateTicket
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $B_UpdateTicket_Draw_X
	$System_Drawing_Point.Y = $B_UpdateTicket_Draw_Y
	$B_UpdateTicket.Location = $System_Drawing_Point
	
	#--Select Page--#
	$L_Centerline = ($WindowWidth - $Spacer) / 2
	$M_Centerline = ($WindowWidth / 2)
	$R_Centerline = $L_Centerline + $Spacer
	$B_Advanced_Draw_X = [Math]::Abs([Math]::Floor($M_Centerline - ($B_Advanced.Width / 2)))
	$CB_Monolithic_Draw_X = [Math]::Abs([Math]::Floor($L_Centerline - $CB_Monolithic.Width))
	$CB_Store_Draw_X = [Math]::Abs([Math]::Floor($R_Centerline))
	
	# $B_Advanced
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $B_Advanced_Draw_X
	$System_Drawing_Point.Y = $B_Right.Location.Y
	$B_Advanced.Location = $System_Drawing_Point
	
	# $CB_Monolithic
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $CB_Monolithic_Draw_X
	$System_Drawing_Point.Y = $B_Right.Location.Y
	$CB_Monolithic.Location = $System_Drawing_Point
	
	# $CB_Store
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $CB_Store_Draw_X
	$System_Drawing_Point.Y = $B_Right.Location.Y
	$CB_Store.Location = $System_Drawing_Point
	
	# Show hide start pic
	if($script:ProgramState -eq "StartNew"){
		if($SCBS.Width -gt '600'){
			$PB_StartPic.Visible = $true
		} else {
			$PB_StartPic.Visible = $false
		}
	} elseif($script:ProgramState -ne "StartExist"){
		$PB_StartPic.Visible = $false
	}	
}

Write-Verbose "Define Close"
$FormEvent_Close={
	Write-Verbose "Form Event Close"
	$timer_backupControl.Stop()
	$timer_norfileConnect.Stop()

	net use $script:DriveLetter /delete | Out-Null
}

#endregion GUI Handlers

#region Form Code
#--Form--#
Write-Verbose "Define Properties Form"
$SCBS.AutoScaleMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 406
$System_Drawing_Size.Width = 484
$SCBS.ClientSize = $System_Drawing_Size
$SCBS.DataBindings.DefaultDataSourceUpdateMode = 0
$script:Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('.\OUIT_Logo.ico')
$SCBS.Icon = $script:Icon
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 445
$System_Drawing_Size.Width = 500
$SCBS.MinimumSize = $System_Drawing_Size
$SCBS.MaximizeBox = $true
$SCBS.Name = "SCBS"
$SCBS.StartPosition = 1
$SCBS.Text = "Service Center Backup Script"
$SCBS.add_Load($FormEvent_Load)
$SCBS.add_Resize($FormEvent_Resize)
$SCBS.add_FormClosing($FormEvent_Close)

#--Timer1--#
Write-Verbose "Define Properties Timer 1"
$timer_backupControl.Enabled = $false
$timer_backupControl.Interval = 500
$timer_backupControl.add_Tick($timer_backupControl_Tick)

#--Timer2--#
Write-Verbose "Define Properties Timer 2"
$timer_norfileConnect.Enabled = $false
$timer_norfileConnect.Interval = 500
$timer_norfileConnect.add_Tick($timer_norfileConnect_Tick)

#--ToolTip Handler--#
Write-Verbose "Define Properties Tool Tips"
$toolTip1.add_Popup($handler_toolTip1_Popup)

#--Picture Box Splash--#
Write-Verbose "Define Properties Picture Box"
$PB_splash.Anchor = 13
$PB_splash.DataBindings.DefaultDataSourceUpdateMode = 0
$PB_splash_ImageLocation = Join-Path $script:WorkingDir "\OUIT_Logo.png"
Write-Debug "Splash Image Location: $PB_splash_ImageLocation"
$PB_splash.Image = [System.Drawing.Image]::FromFile($PB_splash_ImageLocation)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = -2
$PB_splash.Location = $System_Drawing_Point
$PB_splash.Name = "PB_splash"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 120
$System_Drawing_Size.Width = 484
$PB_splash.Size = $System_Drawing_Size
$PB_splash.SizeMode = 3
$PB_splash.TabStop = $False
$SCBS.Controls.Add($PB_splash)

#--Picture Box Start--#
Write-Verbose "Define Properties Picture Box Start"
$PB_StartPic.Anchor = 9
$PB_StartPic.DataBindings.DefaultDataSourceUpdateMode = 0
$PB_StartPic_ImageLocation = Join-Path $script:WorkingDir "\OUIT_Logo-02.png"
Write-Debug "Splash Image Locationn: $PB_StartPic_ImageLocation"
$PB_StartPic.Image = [System.Drawing.Image]::FromFile($PB_StartPic_ImageLocation)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 377
$System_Drawing_Point.Y = 7
$PB_StartPic.Location = $System_Drawing_Point
$PB_StartPic.Name = "PB_StartPic"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 100
$System_Drawing_Size.Width = 100
$PB_StartPic.Size = $System_Drawing_Size
$PB_StartPic.SizeMode = 3
$PB_StartPic.TabStop = $False
$PB_StartPic.Visible = $False
$SCBS.Controls.Add($PB_StartPic)

#--Label HWNumber--#
Write-Verbose "Define Properties Label HWNumber"
$label_HWNumber.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 10
$label_HWNumber.Location = $System_Drawing_Point
$label_HWNumber.Name = "label_HWNumber"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_HWNumber.Size = $System_Drawing_Size
$label_HWNumber.Text = "HW Number"
$SCBS.Controls.Add($label_HWNumber)

#--TB HWNumber--#
Write-Verbose "Define Properties TB HWNumber"
$TB_HWNumber.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 29
$TB_HWNumber.Location = $System_Drawing_Point
$TB_HWNumber.Name = "TB_HWNumber"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$TB_HWNumber.Size = $System_Drawing_Size
$TB_HWNumber.TabIndex = 9
$SCBS.Controls.Add($TB_HWNumber)

#--Label Caller--#
Write-Verbose "Define Properties Label Caller"
$label_Caller.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 9
$label_Caller.Location = $System_Drawing_Point
$label_Caller.Name = "label_Caller"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_Caller.Size = $System_Drawing_Size
$label_Caller.Text = "Caller 4x4"
$label_Caller.TextAlign = 256
$SCBS.Controls.Add($label_Caller)

#--TB Caller--#
Write-Verbose "Define Properties TB Caller"
$TB_Caller.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 29
$TB_Caller.Location = $System_Drawing_Point
$TB_Caller.Name = "TB_Caller"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$TB_Caller.Size = $System_Drawing_Size
$TB_Caller.TabIndex = 10
$SCBS.Controls.Add($TB_Caller)

#--Label Caller Phone--#
Write-Verbose "Define Properties Label Caller Phone"
$label_callerPhone.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 118
$System_Drawing_Point.Y = 10
$label_callerPhone.Location = $System_Drawing_Point
$label_callerPhone.Name = "label_callerPhone"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_callerPhone.Size = $System_Drawing_Size
$label_callerPhone.Text = "Caller Phone"
$label_callerPhone.add_Click($handler_label6_Click)
$SCBS.Controls.Add($label_callerPhone)

#--TB Caller Phone--#
Write-Verbose "Define Properties TB Caller Phone"
$TB_CallerPhone.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 118
$System_Drawing_Point.Y = 28
$TB_CallerPhone.Location = $System_Drawing_Point
$TB_CallerPhone.Name = "TB_CallerPhone"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$TB_CallerPhone.Size = $System_Drawing_Size
$TB_CallerPhone.TabIndex = 11
$TB_CallerPhone.add_TextChanged($handler_textBox5_TextChanged)
$SCBS.Controls.Add($TB_CallerPhone)

#--Label User--#
Write-Verbose "Define Properties Label User"
$label_User.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 52
$label_User.Location = $System_Drawing_Point
$label_User.Name = "label_User"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_User.Size = $System_Drawing_Size
$label_User.Text = "User 4x4"
$label_User.TextAlign = 256
$SCBS.Controls.Add($label_User)

#--TB User--#
Write-Verbose "Define Properties TB User"
$TB_User.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 70
$TB_User.Location = $System_Drawing_Point
$TB_User.Name = "TB_User"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$TB_User.Size = $System_Drawing_Size
$TB_User.TabIndex = 12
$SCBS.Controls.Add($TB_User)

#--Label User Phone--#
Write-Verbose "Define Properties Label User Phone"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 118
$System_Drawing_Point.Y = 53
$label_userPhone.Location = $System_Drawing_Point
$label_userPhone.Name = "label_userPhone"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_userPhone.Size = $System_Drawing_Size
$label_userPhone.Text = "User Phone"
$label_userPhone.add_Click($handler_label7_Click)
$SCBS.Controls.Add($label_userPhone)

#--TB User Phone--#
Write-Verbose "Define Properties TB User Phone"
$TB_UserPhone.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 118
$System_Drawing_Point.Y = 70
$TB_UserPhone.Location = $System_Drawing_Point
$TB_UserPhone.Name = "TB_UserPhone"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$TB_UserPhone.Size = $System_Drawing_Size
$TB_UserPhone.TabIndex = 13
$TB_UserPhone.add_TextChanged($handler_textBox6_TextChanged)
$SCBS.Controls.Add($TB_UserPhone)

#--Label Warranty--#
Write-Verbose "Define Properties Label Warranty"
$label_Warranty.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 225
$System_Drawing_Point.Y = 13
$label_Warranty.Location = $System_Drawing_Point
$label_Warranty.Name = "label_Warranty"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_Warranty.Size = $System_Drawing_Size
$label_Warranty.Text = "Warranty"
$SCBS.Controls.Add($label_Warranty)

#--CoB Warranty--#
Write-Verbose "Define Properties CoB Warranty"
$CoB_Warranty.DataBindings.DefaultDataSourceUpdateMode = 0
$CoB_Warranty.FormattingEnabled = $True
$CoB_Warranty.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$CoB_Warranty.AutoCompleteSource = 256
$CoB_Warranty.Items.Add("Dell Warranty")|Out-Null
$CoB_Warranty.Items.Add("Faculty/Staff")|Out-Null
$CoB_Warranty.Items.Add("HP Warranty")|Out-Null
$CoB_Warranty.Items.Add("Lenovo Warranty")|Out-Null
$CoB_Warranty.Items.Add("Triage Services")|Out-Null
$CoB_Warranty.Items.Add("Out of Warranty")|Out-Null
$CoB_Warranty.Items.Add("None (Bootcamp, VM, etc)")|Out-Null
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 224
$System_Drawing_Point.Y = 28
$CoB_Warranty.Location = $System_Drawing_Point
$CoB_Warranty.Name = "CoB_Warranty"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 122
$CoB_Warranty.Size = $System_Drawing_Size
$CoB_Warranty.TabIndex = 14
$CoB_Warranty.add_SelectedIndexChanged($handler_CoB_Warranty_SelectedIndexChanged)
$SCBS.Controls.Add($CoB_Warranty)

#--Label Make--#
Write-Verbose "Define Properties Label Make"
$label_Make.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 352
$System_Drawing_Point.Y = 10
$label_Make.Location = $System_Drawing_Point
$label_Make.Name = "label_Make"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_Make.Size = $System_Drawing_Size
$label_Make.Text = "Make"
$SCBS.Controls.Add($label_Make)

#--CoB Make--#
Write-Verbose "Define Properties CB Make"
$CoB_Make.DataBindings.DefaultDataSourceUpdateMode = 0
$CoB_Make.FormattingEnabled = $True
$CoB_Make.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$CoB_Make.AutoCompleteSource = 256
$CoB_Make.Items.Add("Apple")|Out-Null
$CoB_Make.Items.Add("Dell")|Out-Null
$CoB_Make.Items.Add("HP")|Out-Null
$CoB_Make.Items.Add("Lenovo")|Out-Null
$CoB_Make.Items.Add("Other")|Out-Null
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 352
$System_Drawing_Point.Y = 28
$CoB_Make.Location = $System_Drawing_Point
$CoB_Make.Name = "CB_Make"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 121
$CoB_Make.Size = $System_Drawing_Size
$CoB_Make.TabIndex = 15
$SCBS.Controls.Add($CoB_Make)

#--Label Partnership Group--#
Write-Verbose "Define Properties Label Partnership"
$label_PartnershipGroup.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 224
$System_Drawing_Point.Y = 53
$label_PartnershipGroup.Location = $System_Drawing_Point
$label_PartnershipGroup.Name = "label_PartnershipGroup"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_PartnershipGroup.Size = $System_Drawing_Size
$label_PartnershipGroup.Text = "Partnership Group"
$SCBS.Controls.Add($label_PartnershipGroup)

#--CoB Partnership Group--#
Write-Verbose "Define Properties CoB Partnership"
$CoB_PartnershipGroup.DataBindings.DefaultDataSourceUpdateMode = 0
$CoB_PartnershipGroup.FormattingEnabled = $true
$CoB_PartnershipGroup.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$CoB_PartnershipGroup.AutoCompleteSource = 256
foreach($_ in $partnership_translation.Keys){
	$CoB_PartnershipGroup.Items.Add($_) | Out-Null
}
$CoB_PartnershipGroup.Sorted = $true
<#
$CoB_PartnershipGroup.Items.Add("Administration Faculty/Staff")|Out-Null
$CoB_PartnershipGroup.Items.Add("Athletics (Faculty/Staff)")|Out-Null
$CoB_PartnershipGroup.Items.Add("College of Architecture (All)")|Out-Null
$CoB_PartnershipGroup.Items.Add("College of Earth and Energy (All)")|Out-Null
$CoB_PartnershipGroup.Items.Add("College of Engineering (All)")|Out-Null
$CoB_PartnershipGroup.Items.Add("College of Journalism (IT Store - Apple)")|Out-Null
$CoB_PartnershipGroup.Items.Add("Honors (Faculty/Staff)")|Out-Null
$CoB_PartnershipGroup.Items.Add("Housing & Food Services (Faculty/Staff)")|Out-Null
$CoB_PartnershipGroup.Items.Add("Information Technology (All)")|Out-Null
$CoB_PartnershipGroup.Items.Add("Musical Theatre (Faculty/Staff)")|Out-Null
$CoB_PartnershipGroup.Items.Add("Physical Plant (Faculty/Staff)")|Out-Null
$CoB_PartnershipGroup.Items.Add("Research and Graduate College (Faculty/Staff)")|Out-Null
$CoB_PartnershipGroup.Items.Add("School of Art and Art History (Lab PCs Only)")|Out-Null
$CoB_PartnershipGroup.Items.Add("Student Affairs (Faculty/Staff)")|Out-Null
$CoB_PartnershipGroup.Items.Add("Technology Development/CCEW (Faculty/Staff)")|Out-Null
$CoB_PartnershipGroup.Items.Add("No Partnership")|Out-Null
#>
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 224
$System_Drawing_Point.Y = 70
$CoB_PartnershipGroup.Location = $System_Drawing_Point
$CoB_PartnershipGroup.Name = "CoB_PartnershipGroup"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 249
$CoB_PartnershipGroup.Size = $System_Drawing_Size
$CoB_PartnershipGroup.TabIndex = 16
$SCBS.Controls.Add($CoB_PartnershipGroup)

#--Label SN--#
Write-Verbose "Define Properties Label SN"
$label_SN.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 93
$label_SN.Location = $System_Drawing_Point
$label_SN.Name = "label_SN"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 206
$label_SN.Size = $System_Drawing_Size
$label_SN.Text = "Service Tag / Serial Number"
$SCBS.Controls.Add($label_SN)

#--TB SN--#
Write-Verbose "Define Properties TB SN"
$TB_SN.Anchor = 13
$TB_SN.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 111
$TB_SN.Location = $System_Drawing_Point
$TB_SN.Name = "TB_SN"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 461
$TB_SN.Size = $System_Drawing_Size
$TB_SN.TabIndex = 17
$SCBS.Controls.Add($TB_SN)

#--Label Pickup--#
Write-Verbose "Define Properties Label Pickup"
$label_Pickup.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 134
$label_Pickup.Location = $System_Drawing_Point
$label_Pickup.Name = "label_Pickup"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_Pickup.Size = $System_Drawing_Size
$label_Pickup.Text = "Pickup Location"
$label_Pickup.TextAlign = 256
$SCBS.Controls.Add($label_Pickup)

#--TB Pickup--#
Write-Verbose "Define Properties TB Pickup"
$TB_Pickup.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 152
$TB_Pickup.Location = $System_Drawing_Point
$TB_Pickup.Name = "TB_Pickup"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 227
$TB_Pickup.Size = $System_Drawing_Size
$TB_Pickup.TabIndex = 18
$SCBS.Controls.Add($TB_Pickup)

#--Label Dropoff--#
Write-Verbose "Define Properties Label Dropoff"
$label_Dropoff.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 246
$label_Dropoff_Draw_Y = 134
$System_Drawing_Point.Y = $label_Dropoff_Draw_Y
$label_Dropoff.Location = $System_Drawing_Point
$label_Dropoff.Name = "label_Dropoff"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_Dropoff.Size = $System_Drawing_Size
$label_Dropoff.Text = "Dropoff Location"
$label_Dropoff.TextAlign = 256
$SCBS.Controls.Add($label_Dropoff)

#--TB Dropoff--#
Write-Verbose "Define Properties TB Dropoff"
$TB_Dropoff.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 246
$TB_Dropoff_Draw_Y = 152
$System_Drawing_Point.Y = $TB_Dropoff_Draw_Y
$TB_Dropoff.Location = $System_Drawing_Point
$TB_Dropoff.Name = "TB_Dropoff"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 227
$TB_Dropoff.Size = $System_Drawing_Size
$TB_Dropoff.TabIndex = 19
$SCBS.Controls.Add($TB_Dropoff)

#--Label Description--#
Write-Verbose "Define Properties Label Short Description"
$label_Description.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 175
$label_Description.Location = $System_Drawing_Point
$label_Description.Name = "label_Description"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_Description.Size = $System_Drawing_Size
$label_Description.Text = "Short Description"
$label_Description.add_Click($handler_label9_Click)
$SCBS.Controls.Add($label_Description)

#--TB Description--#
Write-Verbose "Define Properties TB Description"
$TB_Description.Anchor = 13
$TB_Description.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 193
$TB_Description.Location = $System_Drawing_Point
$TB_Description.Name = "TB_Description"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 461
$TB_Description.Size = $System_Drawing_Size
$TB_Description.TabIndex = 20
$SCBS.Controls.Add($TB_Description)

#--Label Comments--#
Write-Verbose "Define Properties Label Comments"
$label_Comments.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 216
$label_Comments.Location = $System_Drawing_Point
$label_Comments.Name = "label_Comments"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$label_Comments.Size = $System_Drawing_Size
$label_Comments.Text = "Comments"
$label_Comments.TextAlign = 256
$SCBS.Controls.Add($label_Comments)

#--RTB Comments--#
Write-Verbose "Define Properties RTB Comments"
$RTB_Comments.Anchor = 15
$RTB_Comments.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 234
$RTB_Comments.Location = $System_Drawing_Point
$RTB_Comments.Name = "RTB_Comments"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 115
$System_Drawing_Size.Width = 459
$RTB_Comments.Size = $System_Drawing_Size
$RTB_Comments.TabIndex = 21
$RTB_Comments.Text = ""
$SCBS.Controls.Add($RTB_Comments)

#--Button New Ticket--#
Write-Verbose "Define Properties Button New Ticket"
$B_NewTicket.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 138
$B_NewTicket_Draw_Y = 131
$System_Drawing_Point.Y = $B_NewTicket_Draw_Y
$B_NewTicket.Location = $System_Drawing_Point
$B_NewTicket.Name = "B_NewTicket"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 60
$B_NewTicket_Size_W = 100
$System_Drawing_Size.Width = $B_NewTicket_Size_W
$B_NewTicket.Size = $System_Drawing_Size
$B_NewTicket.TabIndex = 0
$B_NewTicket.Text = "Create New Hardware Ticket"
$B_NewTicket.UseVisualStyleBackColor = $True
$B_NewTicket.add_Click($B_NewTicket_OnClick)
$SCBS.Controls.Add($B_NewTicket)

#--Button Update Ticket--#
Write-Verbose "Define Properties Button Update Ticket"
$B_UpdateTicket.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 245
$B_UpdateTicket_Draw_Y = 131
$System_Drawing_Point.Y = $B_UpdateTicket_Draw_Y
$B_UpdateTicket.Location = $System_Drawing_Point
$B_UpdateTicket.Name = "B_UpdateTicket"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 60
$System_Drawing_Size.Width = 100
$B_UpdateTicket.Size = $System_Drawing_Size
$B_UpdateTicket.TabIndex = 1
$B_UpdateTicket.Text = "Update Existing Hardware Ticket"
$B_UpdateTicket.UseVisualStyleBackColor = $True
$B_UpdateTicket.add_Click($B_UpdateTicket_OnClick)
$SCBS.Controls.Add($B_UpdateTicket)

#--Button Right--#
Write-Verbose "Define Properties Button 1"
$B_Right.Anchor = 10
$B_Right.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 398
$System_Drawing_Point.Y = 355
$B_Right.Location = $System_Drawing_Point
$B_Right.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$B_Right.Size = $System_Drawing_Size
$B_Right.TabIndex = 30
$B_Right.Text = "Next"
$B_Right.UseVisualStyleBackColor = $True
$B_Right.add_Click($B_Right_OnClick)
$SCBS.Controls.Add($B_Right)

#--Button Left--#
Write-Verbose "Define Properties Button 2"
$B_Left.DataBindings.DefaultDataSourceUpdateMode = 0
$B_Left.Anchor = 6
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 355
$B_Left.Location = $System_Drawing_Point
$B_Left.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$B_Left.Size = $System_Drawing_Size
$B_Left.TabIndex = 31
$B_Left.Text = "Back"
$B_Left.UseVisualStyleBackColor = $True
$B_Left.Visible = $False
$B_Left.add_Click($B_Left_OnClick)
$SCBS.Controls.Add($B_Left)

#--TreeView--#
Write-Verbose "Define Properties TreeView"
$treeView.Anchor = 15
$treeView.BorderStyle = 1
$treeView.CheckBoxes = $True
$treeView.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 12
$treeView.Location = $System_Drawing_Point
$treeView.Name = "treeView"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 335
$System_Drawing_Size.Width = 461
$treeView.Size = $System_Drawing_Size
$treeView.Visible = $False
$SCBS.Controls.Add($treeView)

#--Button Advanced--#
$B_Advanced.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 211
$System_Drawing_Point.Y = 355
$B_Advanced.Location = $System_Drawing_Point
$B_Advanced.Name = "B_Advanced"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 60
$B_Advanced.Size = $System_Drawing_Size
$B_Advanced.TabIndex = 29
$B_Advanced.Text = "Advanced"
$B_Advanced.UseVisualStyleBackColor = $True
$B_Advanced.add_Click($B_Advanced_OnClick)
$SCBS.Controls.Add($B_Advanced)

#--CB Monolithic--#
$CB_Monolithic.CheckAlign = 64
$CB_Monolithic.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 138
$System_Drawing_Point.Y = 355
$CB_Monolithic.Location = $System_Drawing_Point
$CB_Monolithic.Name = "CB_Monolithic"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$CB_Monolithic.Size = $System_Drawing_Size
$CB_Monolithic.TabIndex = 27
$CB_Monolithic.Text = "Monolithic"
$CB_Monolithic.TextAlign = 64
$CB_Monolithic.UseVisualStyleBackColor = $True
$SCBS.Controls.Add($CB_Monolithic)

#--CB Store--#
$CB_Store.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 245
$System_Drawing_Point.Y = 355
$CB_Store.Location = $System_Drawing_Point
$CB_Store.Name = "CB_Store"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$CB_Store.Size = $System_Drawing_Size
$CB_Store.TabIndex = 26
$CB_Store.Text = "Store"
$CB_Store.UseVisualStyleBackColor = $True
$SCBS.Controls.Add($CB_Store)

#--RTB Output--#
Write-Verbose "Define Properties RTB Output"
$RTB_Output.Anchor = 15
$RTB_Output.DataBindings.DefaultDataSourceUpdateMode = 0
$RTB_Output.Font = New-Object System.Drawing.Font("Consolas",8,0,3,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 12
$RTB_Output.Location = $System_Drawing_Point
$RTB_Output.Name = "richTextBox1"
$RTB_Output.ReadOnly = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 335
$System_Drawing_Size.Width = 461
$RTB_Output.Size = $System_Drawing_Size
$RTB_Output.Text = ""
$RTB_Output.WordWrap = $true
$RTB_Output.Visible = $false
$SCBS.Controls.Add($RTB_Output)

#--Status Bar--#
Write-Verbose "Define Properties Status Bar"
$statusBar.DataBindings.DefaultDataSourceUpdateMode = 0
$statusBar.Font = New-Object System.Drawing.Font("Consolas",8.25,0,3,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 355
$statusBar.Location = $System_Drawing_Point
$statusBar.Name = "statusBar"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 22
$System_Drawing_Size.Width = 485
$statusBar.Size = $System_Drawing_Size
$statusBar.Text = ""
$SCBS.Controls.Add($statusBar)

Write-Verbose "Form Definitions Complete"
#endregion Form Code

# Save the initial state of the form
$InitialFormWindowState = $SCBS.WindowState
# Init the OnLoad event to correct the initial state of the form
$SCBS.add_Load($OnLoadForm_StateCorrection)
# Show the Form
$SCBS.ShowDialog()| Out-Null

} # End Form Function

#region Common Functions

function popupError() {
	#--Types--#
	# # | Options 			  | Return 	#
	# 0 | OK					  | 1		#
	# 1 | OK, Cancel			  | 1,2		#
	# 2 | Abort, Retry, Ignore	  | 3,4,5	#
	# 3 | Yes, No, Cancel 		  | 6,7,2	#
	# 4 | Yes, No				  | 6,7		#
	# 5 | Retry, Cancel			  | 4,2		#
	# 6 | Cancel, Try Again, Cont.| 2,10,11	#
	param(
		[Parameter(Position = 1,Mandatory = $true)][String]$title,
		[Parameter(Position = 2,Mandatory = $true)][String]$msg,
		[Parameter(Position = 3,Mandatory = $false)][Int]$type = '0',
		[Parameter(Position = 4,Mandatory = $false)][Int]$timeout = '0'
	)
	
	Write-Warning $msg
	$popWindow = New-Object -ComObject Wscript.Shell
	return $popWindow.Popup($msg, $timeout, $title, $type)
}

function popupBalloon() {
	# Popups a notification balloon in the Windows system tray.
    param(
        [Parameter(Position = 1,Mandatory = $true)][String]$title,
        [Parameter(Position = 2,Mandatory = $true)][String]$msg,
        [Parameter(Position = 3,Mandatory = $false)]
        [ValidateSet("None", "Warning", "Error")]
        [String]$type = "None",
        [Parameter(Position = 4,Mandatory = $false)][String]$timeout = 0
    )
    Write-Verbose ("Popup Balloon: " + $msg)

    switch($type){
        "None" { $toolTipIcon = [system.windows.forms.tooltipicon]::None }
        "Warning" { $toolTipIcon = [system.windows.forms.tooltipicon]::Warning }
        "Error" { $toolTipIcon = [system.windows.forms.tooltipicon]::Error }
    }

    $notify = new-object system.windows.forms.notifyicon
    $notify.icon = $script:Icon
    $notify.visible = $true
    $notify.showballoontip($timeout, $title, $msg, $toolTipIcon)
}

function updateRTB() {
	# Redraws the text in the Status RTB
	$RTB_Output.Clear()
	foreach($Node in $script:CheckedNodes){
		$RTB_Output.AppendText($Node.Tag)
		$RTB_Output.AppendText("`n")
		$RTB_Output.Select($RTB_Output.Text.Length, 0)
		$RTB_Output.ScrolltoCaret()
	}
	$SCBS.Update()
}

function updateBarGraphic($msg) {
	# Draws waiting graphic defines in $script:progArr after the provided message 
	$script:progArrCount++
	$script:progArrCount = $script:progArrCount%$script:progArr.Length
	#Write-Debug ("ProgArrCount: " + $script:progArrCount)
	$statusBar.Text = $msg + $script:progArr[$script:progArrCount]
}

function setGuiState($guiState) {
# Set the program GUI to the supplied value

	#-------- List of GUI Items ---------#
	#									 #
	# $TB_HWNumber						 #
	# $PB_splash 		splash screen pic#
	# $B_UpdateTicket					 #
	# $B_NewTicket		 				 #
	# $B_Right			next/cancel/exit #
	# $B_Left			back 			 #
	# $label_HWNumber					 #
	# $TB_HWNumber						 #
	# $RTB_Output		status output	 #
	# $label_Caller						 #
	# $TB_Caller						 #
	# $label_User				 		 #
	# $TB_User							 #
	# $label_callerPhone 				 #
	# $TB_CallerPhone					 #
	# $label_userPhone					 #
	# $TB_UserPhone						 #
	# $label_Warranty					 #
	# $CoB_Warranty						 #
	# $label_PartnershipGroup			 #
	# $CoB_PartnershipGroup				 #
	# $label_Pickup						 #
	# $TB_Pickup						 #
	# $label_Dropoff					 #	
	# $TB_Dropoff						 #
	# $label_Description				 #
	# $TB_Description					 #
	# $label_Comments					 #
	# $RTB_Comments						 #
	# $treeView			file picker		 #
	#									 #
	# $statusBar		status bar 		 #
	
	switch($guiState){
		"Splash" {
			$script:ProgramState = "Splash"
			[bool]$script:NoReturn = $false			#controlls allowing return to Splash
			
			$PB_splash.Visible = $true
			$B_NewTicket.Visible = $true
			$B_UpdateTicket.Visible = $true
			$B_Right.Visible = $false
			$B_Left.Visible = $false
			$label_HWNumber.Visible = $false
			$TB_HWNumber.Visible = $false
			$RTB_Output.Visible = $false
			$label_Caller.Visible = $false
			$TB_Caller.Visible = $false
			$label_User.Visible = $false
			$TB_User.Visible = $false			
			$label_callerPhone.Visible = $false
			$TB_CallerPhone.Visible = $false
			$label_userPhone.Visible = $false
			$TB_UserPhone.Visible = $false
			$label_Warranty.Visible = $false
			$CoB_Warranty.Visible = $false
			$label_Make.Visible = $false
			$CoB_Make.Visible = $false
			$label_PartnershipGroup.Visible = $false
			$CoB_PartnershipGroup.Visible = $false
			$label_SN.Visible = $false
			$TB_SN.Visible = $false
			$label_Pickup.Visible = $false
			$TB_Pickup.Visible = $false
			$label_Dropoff.Visible = $false
			$TB_Dropoff.Visible = $false
			$label_Description.Visible = $false
			$TB_Description.Visible = $false
			$label_Comments.Visible = $false
			$RTB_Comments.Visible = $false
			$treeView.Visible = $false
			$B_Advanced.Visible = $false
			$CB_Monolithic.Visible = $false
			$CB_Store.Visible = $false
			$PB_StartPic.Visible = $false
			
			$B_Right.Text = "Next"
			$B_Left.Text = "Back"
			#$statusBar.Text = ""
		}
		
		"StartNew" {
			$script:ProgramState = "StartNew"
			
			$PB_splash.Visible = $false
			$B_NewTicket.Visible = $false
			$B_UpdateTicket.Visible = $false
			$B_Right.Visible = $true
			$B_Left.Visible = !$script:NoReturn
			$label_HWNumber.Visible = $false
			$TB_HWNumber.Visible = $false
			$RTB_Output.Visible = $false
			$label_Caller.Visible = $true
			$TB_Caller.Visible = $true
			$label_User.Visible = $true
			$TB_User.Visible = $true			
			$label_callerPhone.Visible = $true
			$TB_CallerPhone.Visible = $true
			$label_userPhone.Visible = $true
			$TB_UserPhone.Visible = $true
			$label_Warranty.Visible = $true
			$CoB_Warranty.Visible = $true
			$label_Make.Visible = $true
			$CoB_Make.Visible = $true
			$label_PartnershipGroup.Visible = $true
			$CoB_PartnershipGroup.Visible = $true
			$label_SN.Visible = $true
			$TB_SN.Visible = $true
			$label_Pickup.Visible = $true
			$TB_Pickup.Visible = $true
			$label_Dropoff.Visible = $true
			$TB_Dropoff.Visible = $true
			$label_Description.Visible = $true
			$TB_Description.Visible = $true
			$label_Comments.Visible = $true
			$RTB_Comments.Visible = $true
			$treeView.Visible = $false
			$B_Advanced.Visible = $false
			$CB_Monolithic.Visible = $false
			$CB_Store.Visible = $false
			
			if($SCBS.Width -gt '600'){
				$PB_StartPic.Visible = $true
			} else {
				$PB_StartPic.Visible = $false
			}
			
			$B_Right.Text = "Next"
			$B_Left.Text = "Back"
			#$statusBar.Text = ""
		}
		
		"StartExist" {
			$script:ProgramState = "StartExist"
			
			$PB_splash.Visible = $false
			$B_NewTicket.Visible = $false
			$B_UpdateTicket.Visible = $false
			$B_Right.Visible = $true
			$B_Left.Visible = $true
			$label_HWNumber.Visible = $true
			$TB_HWNumber.Visible = $true
			$RTB_Output.Visible = $false
			$label_Caller.Visible = $false
			$TB_Caller.Visible = $false
			$label_User.Visible = $false
			$TB_User.Visible = $false			
			$label_callerPhone.Visible = $false
			$TB_CallerPhone.Visible = $false
			$label_userPhone.Visible = $false
			$TB_UserPhone.Visible = $false
			$label_Warranty.Visible = $false
			$CoB_Warranty.Visible = $false
			$label_Make.Visible = $false
			$CoB_Make.Visible = $false
			$label_PartnershipGroup.Visible = $false
			$CoB_PartnershipGroup.Visible = $false
			$label_SN.Visible = $false
			$TB_SN.Visible = $false
			$label_Pickup.Visible = $false
			$TB_Pickup.Visible = $false
			$label_Dropoff.Visible = $false
			$TB_Dropoff.Visible = $false
			$label_Description.Visible = $false
			$TB_Description.Visible = $false
			$label_Comments.Visible = $false
			$RTB_Comments.Visible = $false
			$treeView.Visible = $false
			$B_Advanced.Visible = $false
			$CB_Monolithic.Visible = $false
			$CB_Store.Visible = $false
			$PB_StartPic.Visible = $true
			
			$B_Right.Text = "Next"
			$B_Left.Text = "Back"
			#$statusBar.Text = ""
		}
		
		"Select" {
			$script:ProgramState = "Select"
			$script:NoReturn = $true	#controlls allowing return to splash
			
			$PB_splash.Visible = $false
			$B_NewTicket.Visible = $false
			$B_UpdateTicket.Visible = $false
			$B_Right.Visible = $true
			if($script:StartState -eq "New"){
				$B_Left.Visible = $true
			} elseif($script:StartState -eq "Exist"){
				$B_Left.Visible = $false
			}
			$label_HWNumber.Visible = $false
			$TB_HWNumber.Visible = $false
			$RTB_Output.Visible = $false
			$label_Caller.Visible = $false
			$TB_Caller.Visible = $false
			$label_User.Visible = $false
			$TB_User.Visible = $false			
			$label_callerPhone.Visible = $false
			$TB_CallerPhone.Visible = $false
			$label_userPhone.Visible = $false
			$TB_UserPhone.Visible = $false
			$label_Warranty.Visible = $false
			$CoB_Warranty.Visible = $false
			$label_Make.Visible = $false
			$CoB_Make.Visible = $false
			$label_PartnershipGroup.Visible = $false
			$CoB_PartnershipGroup.Visible = $false
			$label_SN.Visible = $false
			$TB_SN.Visible = $false
			$label_Pickup.Visible = $false
			$TB_Pickup.Visible = $false
			$label_Dropoff.Visible = $false
			$TB_Dropoff.Visible = $false
			$label_Description.Visible = $false
			$TB_Description.Visible = $false
			$label_Comments.Visible = $false
			$RTB_Comments.Visible = $false
			$treeView.Visible = $true
			$B_Advanced.Visible = $true
			$CB_Monolithic.Visible = $false
			$CB_Store.Visible = $false
			$PB_StartPic.Visible = $false
			
			$B_Advanced.Enabled = $true
			$CB_Monolithic.Enabled = $true
			$CB_Store.Enabled = $true
			
			$B_Right.Text = "Next"
			$B_Left.Text = "Back"
			$SCBS.Text = "SCBS: " + $script:hwNumber
			$statusBar.Text = "Select Files and Folders to Backup"
		}
		
		"Confirm" {
			$script:ProgramState = "Confirm"
			
			$PB_splash.Visible = $false
			$B_NewTicket.Visible = $false
			$B_UpdateTicket.Visible = $false	
			$B_Right.Visible = $true
			$B_Left.Visible = $true
			$label_HWNumber.Visible = $false
			$TB_HWNumber.Visible = $false
			$RTB_Output.Visible = $true
			$label_Caller.Visible = $false
			$TB_Caller.Visible = $false
			$label_User.Visible = $false
			$TB_User.Visible = $false
			$label_callerPhone.Visible = $false
			$TB_CallerPhone.Visible = $false
			$label_userPhone.Visible = $false
			$TB_UserPhone.Visible = $false
			$label_Warranty.Visible = $false
			$CoB_Warranty.Visible = $false
			$label_Make.Visible = $false
			$CoB_Make.Visible = $false
			$label_PartnershipGroup.Visible = $false
			$CoB_PartnershipGroup.Visible = $false
			$label_SN.Visible = $false
			$TB_SN.Visible = $false
			$label_Pickup.Visible = $false
			$TB_Pickup.Visible = $false
			$label_Dropoff.Visible = $false
			$TB_Dropoff.Visible = $false
			$label_Description.Visible = $false
			$TB_Description.Visible = $false
			$label_Comments.Visible = $false
			$RTB_Comments.Visible = $false
			$treeView.Visible = $false
			$PB_StartPic.Visible = $false
			
			$B_Advanced.Enabled = $false
			$CB_Monolithic.Enabled = $false
			$CB_Store.Enabled = $false
			
			$B_Right.Text = "Backup"
			$B_Left.Text = "Back"
			$statusBar.Text = "Confirm Files and Folders to Backup"
		}
		
		"Backup" {
			$script:ProgramState = "Backup"
		
			#Disable button to prevent accidental canceling.
			#This is re-enabled on first tick of timer.
			$B_Right.Enabled = $false 
		
			$PB_splash.Visible = $false
			$B_NewTicket.Visible = $false
			$B_UpdateTicket.Visible = $false
			$B_Right.Visible = $true
			$B_Left.Visible = $false
			$label_HWNumber.Visible = $false
			$TB_HWNumber.Visible = $false
			$RTB_Output.Visible = $true
			$label_Caller.Visible = $false
			$TB_Caller.Visible = $false
			$label_User.Visible = $false
			$TB_User.Visible = $false
			$label_callerPhone.Visible = $false
			$TB_CallerPhone.Visible = $false
			$label_userPhone.Visible = $false
			$TB_UserPhone.Visible = $false
			$label_Warranty.Visible = $false
			$CoB_Warranty.Visible = $false
			$label_Make.Visible = $false
			$CoB_Make.Visible = $false
			$label_PartnershipGroup.Visible = $false
			$CoB_PartnershipGroup.Visible = $false
			$label_SN.Visible = $false
			$TB_SN.Visible = $false
			$label_Pickup.Visible = $false
			$TB_Pickup.Visible = $false
			$label_Dropoff.Visible = $false
			$TB_Dropoff.Visible = $false
			$label_Description.Visible = $false
			$TB_Description.Visible = $false
			$label_Comments.Visible = $false
			$RTB_Comments.Visible = $false
			$treeView.Visible = $false
			$PB_StartPic.Visible = $false
			
			$B_Right.Text = "Cancel"
			$B_Left.Text = ""
			$statusBar.Text = "Backing up... "
		}
		
		"Done" {
			$script:ProgramState = "Done"
		
			$PB_splash.Visible = $false
			$B_NewTicket.Visible = $false
			$B_UpdateTicket.Visible = $false
			$B_Right.Visible = $true
			$B_Left.Visible = $false
			$label_HWNumber.Visible = $false
			$TB_HWNumber.Visible = $false
			$RTB_Output.Visible = $true
			$label_Caller.Visible = $false
			$TB_Caller.Visible = $false
			$label_User.Visible = $false
			$TB_User.Visible = $false
			$label_callerPhone.Visible = $false
			$TB_CallerPhone.Visible = $false
			$label_userPhone.Visible = $false
			$TB_UserPhone.Visible = $false
			$label_Warranty.Visible = $false
			$CoB_Warranty.Visible = $false
			$label_Make.Visible = $false
			$CoB_Make.Visible = $false
			$label_PartnershipGroup.Visible = $false
			$CoB_PartnershipGroup.Visible = $false
			$label_SN.Visible = $false
			$TB_SN.Visible = $false
			$label_Pickup.Visible = $false
			$TB_Pickup.Visible = $false
			$label_Dropoff.Visible = $false
			$TB_Dropoff.Visible = $false
			$label_Description.Visible = $false
			$TB_Description.Visible = $false
			$label_Comments.Visible = $false
			$RTB_Comments.Visible = $false
			$treeView.Visible = $false
			$PB_StartPic.Visible = $false
			
			$B_Right.Text = "Exit"
			$B_Left.Text = ""
			$statusBar.Text = "Backup Done."
		}
		
		"Error" {
			# Error is used when verfiyBackup fails or when the user cancels the backup.
			# Allows user to return select state.
			$script:ProgramState = "Error"
		
			$PB_splash.Visible = $false
			$B_NewTicket.Visible = $false
			$B_UpdateTicket.Visible = $false
			$B_Right.Visible = $true
			$B_Left.Visible = $true
			$label_HWNumber.Visible = $false
			$TB_HWNumber.Visible = $false
			$RTB_Output.Visible = $true
			$label_Caller.Visible = $false
			$TB_Caller.Visible = $false
			$label_User.Visible = $false
			$TB_User.Visible = $false
			$label_callerPhone.Visible = $false
			$TB_CallerPhone.Visible = $false
			$label_userPhone.Visible = $false
			$TB_UserPhone.Visible = $false
			$label_Warranty.Visible = $false
			$CoB_Warranty.Visible = $false
			$label_Make.Visible = $false
			$CoB_Make.Visible = $false
			$label_PartnershipGroup.Visible = $false
			$CoB_PartnershipGroup.Visible = $false
			$label_SN.Visible = $false
			$TB_SN.Visible = $false
			$label_Pickup.Visible = $false
			$TB_Pickup.Visible = $false
			$label_Dropoff.Visible = $false
			$TB_Dropoff.Visible = $false
			$label_Description.Visible = $false
			$TB_Description.Visible = $false
			$label_Comments.Visible = $false
			$RTB_Comments.Visible = $false
			$treeView.Visible = $false
			$PB_StartPic.Visible = $false
			
			$B_Right.Text = "Exit"
			$B_Left.Text = "Back"
			$statusBar.Text = "Error: Backup did not complete!"
		}
	}
}

#endregion Common Functions

#region Startup Functions
 
function setupStartEnv() {
	# Attempt to create the backup directory, if it exists, create a directory in it instead.
	if($script:netPath.Length -le '3'){ # Detecting lenth of the netPath so this code only runs once.
		# Norfile Connect Timer sets $script:netPath to the drive letter
		do{
			$script:netPath = Join-Path $script:netPath $script:hwNumber
			New-Item -ItemType Directory -Path $script:netPath -ErrorAction SilentlyContinue
		} until($?)
		Write-Debug $script:netPath
		
		# Gathers the root folder structer for the file picker includes all internal hard drives.
		# Only run this once so we save selections.
		buildTreeView 
		
		# Find the number of computer cores, used to determine how many concurrent compression jobs to run.
		$script:procCount = @(Get-WmiObject win32_processor)[0].NumberofCores
	}
}

function verifyVersion() {
	# Checks version agains version control files on norfile
	# This very simple version control checks for a file named ($script:VersionNo + ".version")
	# Version files are stored in \\norfile.net.ou.edu\it-cxservices\backups\
	Write-Verbose "Verify Version"
	[bool]$script:ValidVersion = $false
	if((Get-ChildItem $script:DriveLetter -Name) -contains ($script:VersionNo + ".version")){
		[bool]$script:ValidVersion = $true
		return 
	}
}

function makeMapNorfileJob() {
	Write-Verbose "Make Norfile Job"
	$mapJob = New-Object PSObject -Property @{
		"Name" = "";
		"Job" = "";
	}
	$mapJob.Name = "Map Job"
	Add-Member -InputObject $mapJob -Name Map -MemberType ScriptMethod -Value {
		#get a random open drive letter to map norfile to. 
		#Setting this here in case there is an issue with the drive letter it can try a different one (hopefully, ~90% success rate)
		$script:DriveLetter = ls function:[g-z]: -n | ?{ !(test-path $_) } | random	
		
		#disable button till the map is done
		$B_Right.Enabled = $false
		
		$parameters = @{
			"NorfileFailOver" = $script:NorfileFailOver;
			"NorFileIP" = $script:NorFileIP;
			"StaticNetPath" = $script:StaticNetPath;
			"DriveLetter" = $script:DriveLetter;
		}
		
		$this.Job = Start-Job -ArgumentList $parameters -ScriptBlock {
			param($parameters)
			$net = New-Object -ComObject WScript.Network
			if($parameters.NorfileFailOver){
				$net.MapNetworkDrive($parameters.DriveLetter, $parameters.NorFileIP, $false, "username", "password") # Redacted
			} else {
				$net.MapNetworkDrive($parameters.DriveLetter, $parameters.StaticNetPath, $false, "username", "password") # Redacted 
			}
		}
		Write-Verbose "Map Job Started"
	}
	Add-Member -InputObject $mapJob -Name State -MemberType ScriptMethod -Value {
		#get the state of the job
		if($this.Job.State -eq $null){
			return "Unprocessed"
		} else {
			return $this.Job.State
		}
	}
	
	return $mapJob
}

#endregion Startup Functions

#region File Picker Functions

function addNode() {
	# Adds a child node to the provided (selected) node in the directory tree
    param ( 
        $selectedNode, 
        $name, 
        $tag 
    ) 
	
    $newNode = new-object System.Windows.Forms.TreeNode  
    $newNode.Name = $name 
    $newNode.Text = $name 
    $newNode.Tag = $tag 
    $selectedNode.Nodes.Add($newNode) | Out-Null 
    return $newNode 
}

function buildTreeView() {
	# Creates a treeView object and populates it with list of internal harddrives.
	# Each Drive, Directory or File is displayed in the treeView as a Node.
	# All nodes have an AfterSelect method. 
	# If node is of type Container the subdirectory is scanned and the nodes added.
	
	# If there's data in the treeView, clear it out
	if($treeView.Nodes.Count -gt '0'){
		$treeView = New-Object System.Windows.Forms.TreeView
	}		

	#sets up method that expands a directory when selected. Directory is only scanned on selection. 
	#Scanning entire directory structure on run would take too long
    $treeView.add_AfterSelect({
		$ErrorActionPreference = "Stop"		# Used to catch non-terminating errors, reset in Finally block
		try{
			Get-ChildItem $this.SelectedNode.Tag | Out-Null		# Trip Security if we don't have permissions to this folder
			if([String]::IsNullOrEmpty($(Get-ChildItem -Path $this.SelectedNode.Tag | Select Name))){
				popupError "Folder Empty" "The selected folder is empty." -timeout 1
			} elseif($this.SelectedNode.Nodes.Count -eq '0' -and (Test-Path $this.SelectedNode.Tag -PathType Container)){
				
				$folders = Get-ChildItem $this.SelectedNode.Tag | Select-Object Name,FullName
		        foreach($folder in $folders) { 
		            $childNode = addNode $this.SelectedNode $folder.Name $folder.FullName 
		        }
				$SCBS.update()
				$this.SelectedNode.Expand()
			}
		} catch {
			if($Error[0].FullyQualifiedErrorId -match "DirUnauthorizedAccessError"){
				popupError "Access Denied" $("Access to " + $this.SelectedNode.Tag + " directory was denied.") -timeout 1
			}
			if($Error[0].FullyQualifiedErrorId -match "ParameterArgumentValidationErrorNullNotAllowed"){
				continue
			}
		} finally {
			$ErrorActionPreference = "Continue"
		}	
    })
	
	$drives = Get-WmiObject -Query "SELECT * from win32_logicaldisk where drivetype='3'"
    
	foreach($drive in $drives) {
		$parentNode = addNode $treeview $drive.DeviceID $($drive.DeviceID + "\")
		$folders = Get-ChildItem $($drive.DeviceID + "\") | Select-Object Name,FullName
        foreach($folder in $folders) { 
            $childNode = addNode $parentNode $folder.Name $folder.FullName 
        } 
	}
}

function getCheckedNodes() {
	# Recursivly cylces through all nodes collecting one that have been selected
	# and adds them to the provided array
    param(
	    [ValidateNotNull()]
	    [System.Windows.Forms.TreeNodeCollection] $NodeCollection,
	    [ValidateNotNull()]
	    [System.Collections.ArrayList]$CheckedNodes
	)
    
    foreach($Node in $NodeCollection)
    {
        if($Node.Checked)
        {
            [void]$CheckedNodes.Add($Node)
        }
        getCheckedNodes $Node.Nodes $CheckedNodes
    }
}

#endregion File Picker Functions

#region Backup Job Functions

function makeJob ($name, $path) {
	# Creates the backup job object
	$jobObject = New-Object PSObject -Property $jobObjectProperties
	$jobObject.Name = $name
	$jobObject.Path = $path
	Add-Member -InputObject $jobObject -Name Backup -MemberType ScriptMethod -Value {
		#run the job
		if($CB_Monolithic.Checked){
			$destination = Join-Path $script:netPath ($script:hwNumber + ".7z")
		} else {
			$destination = Join-Path $script:netPath ($this.Name + ".7z")
		}
		$parameters = @{
			"destination" = $destination;
			"source" = $this.Path;
			"netPath" = $script:netPath;
			"workingDir" = $script:WorkingDir;
			"store" = $CB_Store.Checked;
		}
		
		$this.Job = Start-Job -ArgumentList $parameters -ScriptBlock {
			param($parameters)
			cd $parameters.workingDir			
			if($parameters.store){
				$cmdTime = Measure-Command {.\7z.exe a -mx0 -t7z -ms $parameters.destination $parameters.source}
			} else {
				$cmdTime = Measure-Command {.\7z.exe a -mx5 -t7z -ms $parameters.destination $parameters.source}
			}
			$($parameters.destination + "," + $cmdTime.TotalSeconds.toString()) | Out-File $(Join-Path $parameters.netPath "transferLog.csv") -Append
		} 
		Write-Debug "Job Started"
		Write-Debug $this.Job.State
		Write-Debug	$this.Name
		Write-Debug $this.Path
	}
	Add-Member -InputObject $jobObject -Name State -MemberType ScriptMethod -Value {
		#get the state of the job
		if($this.Job.State -eq $null){
			return "Unprocessed"
		} else {
			return $this.Job.State
		}
	} 
	
	return $jobObject
}

function makeJobQueue () {
	$jobQueue = New-Object PSObject
	Add-Member -InputObject $jobQueue -Name JobList -Value @() -MemberType NoteProperty
	Add-Member -InputObject $jobQueue -Name RunCount -MemberType ScriptMethod -Value {
		Write-Verbose "RunCount"
		# RunCounter is how many jobs are currently run
		[int]$runCounter = '0'
		foreach($_ in $this.JobList){
			if($_.State() -eq "Running"){
				$runCounter++
			}
		}
		Write-Debug "Run count: $runCounter"
		return $runCounter
	}
	Add-Member -InputObject $jobQueue -Name RunNext -MemberType ScriptMethod -Value {
		Write-Verbose "RunNext"
		# Finds the next unprocessed job and runs it
		# Returns more if there are more jobs to run
		# Returns wait if all jobs are running
		# Returns done if there are no more jobs to run and all jobs are done
		
		$statusVar = "done"
		
		# Check there are resources to run more jobs
		if($script:jobQueue.RunCount() -le $script:procCount){
			# Monolithic mode tells the program to store the backup in one single archive.
			if($CB_Monolithic.Checked){ 
				# Check for a running job
				foreach($_ in $this.JobList){
					Write-Debug $_
					if($_.State() -eq "Running"){
						$statusVar = "wait"
						break
					}
				}
				if($statusVar -eq "done"){
					# Here if no jobs are running implying there is no write lock on archive
					foreach($_ in $this.JobList){
						# Find a job to run
						if($_.State() -eq "Unprocessed"){
							$_.Backup()
							$statusVar = "wait"
							break
						}
					}
				}
			} else {
				# Not monolithic
				foreach($_ in $this.JobList){
					Write-Debug $_
					if($_.State() -eq "Unprocessed"){
						# Here if we need to run another job
						$_.Backup()
						$statusVar = "more"
						break
					} elseif($_.State() -eq "Running"){
						# Here if jobs are running
						# Only sets wait if state is done so it doesn't overwrite a more
						$statusVar = "wait"
					}
				}
			}
		} else {
			# Wait for resources to run next
			$statusVar = "wait"
		}
		
		if($statusVar -eq "more"){Write-Verbose "More jobs to run"}
		if($statusVar -eq "wait"){Write-Verbose "Waiting on jobs to finish"}
		if($statusVar -eq "done"){Write-Verbose "All jobs done"}
		
		return $statusVar
	}
	Add-Member -InputObject $jobQueue -Name UpdateGUI -MemberType ScriptMethod -Value {
		Write-Verbose "UpdateGUI"
		
		$statusArray = @()
		$badStatus = "AtBreakpoint","Blocked","Disconnected","Failed","Stopped","Stopping","Suspended","Suspending"
		foreach($_ in $this.JobList){
			if($script:ErrorList -contains $_.Name){
				$statusArray += "[Error]" + $_.Path
			} elseif($_.State() -eq "Unprocessed"){
				$statusArray += $_.Path
			} elseif($_.State() -eq "Completed"){
				$statusArray += "[Done]" + $_.Path
			} elseif($_.State() -eq "Running"){
				$statusArray += "[Running]" + $_.Path
			} elseif($badStatus -contains $_.State()){
				$statusArray += "[Error]" + $_.Path
			}
		}
		
		# Clear text
		$RTB_Output.Clear()
		
		# Create sort filters
		$sort0 = { if($_ -match '\[Error\]'){$matches[1]} }
		$sort1 = { if($_ -match '\[Running\]'){$matches[1]} }
		$sort2 = { if($_ -match '\[Done\]'){$matches[1]} }
		
		# Sort by Not Started > Error > Running > Done
		$statusArray = $statusArray | Sort-Object $sort0,$sort1,$sort2
		
		# There are changes that need to be displayed
		foreach($_ in $statusArray){
			$RTB_Output.AppendText($_)
			$RTB_Output.AppendText("`n")
			$RTB_Output.Select($RTB_Output.Text.Length, 0)
			$RTB_Output.ScrolltoCaret()
		}
		
		$SCBS.Update()
	}
	return $jobQueue
}

function manageJobQueue() {
	switch($script:jobQueue.RunNext()){
		"more" {
			manageJobQueue
		}
		"wait" {
			# Waiting for jobs to finish running
			return
		}
		"done" {
			# All jobs done, verify backup
			$script:jobQueue.UpdateGUI()
			verifyBackup
			if($script:BackupState){
				# Here if there are no other jobs to run and the backup is verified
				popupBalloon "SCBS" "Backup Complete."
				
				$data = @{
					"backuplocation" = $("\\norfile.net.ou.edu\it-cxservices\backups" + $script:netPath.split(":")[1]);
					"completedby" = "scbssa";
					"backupState" = "Backup Completed";
					"date" = Get-Date -Format "yyyy-MM-dd"
				}
				snQuery -backupEnd -sysID $script:SysID -data $data
				
				Write-Verbose "Verified and Done"
				setGuiState "Done" 
			} else {
				# Here if backup failed
				popupBalloon "SCBS" "Backup Failed"
				
				Write-Warning "Error and Done"
				setGuiState "Error" 
			}
		}
	}
}

function verifyBackup() {
	# Verifies that all files have been backuped by checking their existance
	# Returns false if a file does not exist and backup failed
	# Returns true if backup succeeded
	
	Write-Verbose "Verifying Backup"
	$statusBar.Text = "Verifying Backup"
	
	$nameList = @()
	foreach($_ in (Get-ChildItem $script:netPath | Select-Object Name -ExpandProperty Name)){
		Write-Debug $_.Trim(".7z")
		$nameList += $_.Trim(".7z")
	}
	
	[Boolean]$script:BackupState = $true # Backup state variable	
	
	if($CB_Monolithic.Checked){
		# Only need to check existance of one file
		if($nameList -notcontains $script:hwNumber -and $script:ErrorList -notcontains $script:hwNumber){
			# File is missing and it's not in the list of known Errors
			($script:ErrorList += $script:hwNumber) 2> Out-Null # Add it to error list
			
			popupError "! Critcal Error !" "Backup did not complete properly!" 0x0
			$script:BackupState = $false
		}
	} else {
		# Check that each selected node has an associated file
		foreach($Node in $script:CheckedNodes){
			# Get a list of all files by name in the backup directory
			Write-Debug ("NodeName: " + $Node.Name)
			
			if($nameList -notcontains $Node.Name -and $script:ErrorList -notcontains $Node.Name){
				# File is missing and it's not in the list of known Errors
				($script:ErrorList += $Node.Name) 2> Out-Null # Add it to error list
				
				$answer = popupError "! Critcal Error !" $($Node.tag + " did not backup properly!") 0x5
				# If retry set the state on the proper job object
				if($answer -eq '4'){
					[bool]$backuprestarted = $false
					foreach($_ in $script:jobQueue.JobList){
						# Find the node in the job queue to restart
						if($_.Path -eq $Node.Tag){
							$_.Backup()
							$timer_backupControl.Start()
							$backuprestarted = $true
							return
						}
					}
					if(!$backuprestarted){
						popupError "!Critcal Error!" "Cannot restart backup!"
						$script:BackupState = $false
					}
				} else {
					$script:BackupState = $false
				}
			}
		}
	}
	
	# Check that the system info backuped
	if(!(Test-Path ($script:netPath + "\compInfo.txt"))){
		$script:ErrorList += "Collect System Information"
		$answer = popupError "!Critcal Error!" $("Failed to collect system information!") 0x5
		if($answer -eq '4'){
			[bool]$backuprestarted = $false
			foreach($_ in $script:jobQueue.JobList){
				if($_.Path -match "Collect System Information"){
					$_.Backup()
					$timer_backupControl.Start()
					$backuprestarted = $true
					return
				}
			}
			if(!$backuprestarted){
				popupError "!Critcal Error!" "Cannot restart backup!"
				$script:BackupState = $false
			}
		} else {
			$script:BackupState = $false
		}
	}
	
	# Backup verified
}

#endregion Backup Job Functions

#region Service-Now Query Functions

function validateData() {
	param(
		[Parameter(Mandatory = $true)][Hashtable]$data
	)

	#Validates form data before creating ticket.
	Write-Verbose "Validate Data"
	
	#Sets to true and switched to false if something is invalid.
	[bool]$script:validData = $true
	[string]$valDataErrorList = ""
	
	#--Caller 4x4--#
	if($data.caller4x4 -notmatch $script:regex_4x4){
		$valDataErrorList += "`n'" + $data.caller4x4 + "' is an invalid 4x4."
		$script:validData = $false
	}
	
	#--Caller Phone--#
	if($data.callerPhone -notmatch $script:regex_phoneLong -and $data.callerPhone -notmatch $script:regex_phoneLocal){
		$valDataErrorList += "`n'" + $data.callerPhone + "' is an invalid phone number. Please ensure you include the area code for non 325 numbers."
		$script:validData = $false
	}
	
	#--User 4x4--#
	if($data.user4x4 -notmatch $script:regex_4x4){
		$valDataErrorList += "`n'" + $data.user4x4 + "' is an invalid 4x4."
		$script:validData = $false
	}
	
	#--User Phone--#
	if($data.userPhone -notmatch $script:regex_phoneLong -and $data.userPhone -notmatch $script:regex_phoneLocal){
		$valDataErrorList += "`n'" + $data.userPhone + "' is an invalid phone number. Please ensure you include the area code for non 325 numbers."
		$script:validData = $false
	}
	
	#--Warranty--#
	if([String]::IsNullOrEmpty($data.warranty)){
		$valDataErrorList += "`nYou must select a warranty. If the device is no longer under warranty select 'Out of Warranty'."
		$script:validData = $false
	}
	
	#--Make--#
	if([String]::IsNullOrEmpty($data.make)){
		$valDataErrorList += "`nYou must enter the make (Manufaturer) of the device."
		$script:validData = $false
	}
	
	#--Partnership--#
	if([String]::IsNullOrEmpty($data.partnership)){
		$valDataErrorList += "`nYou must select a partnership group. Select 'No Partnership' if user is not covered by another group."
		$script:validData = $false
	}
	
	#--Serial--#
	if([String]::IsNullOrEmpty($data.serial)){
		$valDataErrorList += "`nYou must enter a serial number, service tag or other identifying number."
		$script:validData = $false
	}
	
	#--Pickup Location--#
	if([String]::IsNullOrEmpty($data.pickup)){
		$valDataErrorList += "`nYou must enter a serial number, service tag or other identifying number."
		$script:validData = $false
	} elseif($data.pickup -notmatch $script:regex_inputSanitizer){
		$valDataErrorList += "`nThere are invalid characters in the pickup location."
		$script:validData = $false
	}
	
	#--Dropoff Location--#
	if([String]::IsNullOrEmpty($data.dropoff)){
		$valDataErrorList += "`nYou must supply a dropoff location."
		$script:validData = $false
		#return
	} elseif($data.dropoff -notmatch $script:regex_inputSanitizer){
		$valDataErrorList += "`nThere are invalid characters in the dropoff location."
		$script:validData = $false
	}
	
	#--Short Description--#
	if([String]::IsNullOrEmpty($data.description)){
		$valDataErrorList += "`nYou must supply a short description."
		$script:validData = $false
	} elseif($data.description -notmatch $script:regex_inputSanitizer){
		$valDataErrorList += "`nYou have invalid characters in the short description."
		$script:validData = $false
	}
	
	#--Comments--#
	if($data.comments -notmatch $script:regex_inputSanitizer){
		$valDataErrorList += "`nYou have invalid characters in the comments section."
		$script:validData = $false
	}
	
	#--Display Errors--#
	if(!$script:validData){
		popupError "Invalid Input" $valDataErrorList 
	}
}

function verifyConnection() {
	#Verify Service-Now Connection, provides DNS issue protection/detection
	Write-Verbose "Verify Connection"
	if((ping $script:StaticServiceNowPath -n 1) -match "Ping request could not find host"){
		#Test Failed, attempt dns fallback
		#set a variable so we know to use the failover later
		#[bool]$script:ServiceNowFailOver = $true
		if([string]$(ping $script:ServiceNowIP -n 1) -match "Destination host unreachable."){
			popupError "Network Error" "Cannot access Service-Now. Backup not possible." 
			$SCBS.Close()
		} else {
			#Test Succeeded but there is a DNS issue. Inform user and proceed.
			$script:StaticServiceNowPath = $script:ServiceNowIP
			$script:StaticNetPath = $script:NorFileIP
			popupError "Information" "There was an issue with DNS lookup for Service-Now. Fallback to IP Address has succeeded. Backup will continue..."
		}
	} else {
		Write-Verbose "Service-Now Reachable"
	}
}

function snQuery() {
	# Sends a web query to Service-Now to gather data, update, or create a hardware ticket.
	#
	#----Valid $data Properties-----#
	#								#
	#	caller4x4	assignmentgroup	#
	#	user4x4		serial			#
	#	userPhone	description		#
	#	date		callerPhone		#
	#	model		backuplocation	#
	#	chassis		completedby		#
	#	username	pickup			#
	#	password	dropoff			#
	#	comments	backupState		#
	#	warranty	make			#
	#	os			partnership		#
	
	[CmdletBinding(DefaultParametersetName='None')]
	param(
		[Parameter(ParameterSetName='Verify',Mandatory = $false)][Switch]$verify,
		[Parameter(ParameterSetName='New',Mandatory = $false)][Switch]$new,
		[Parameter(ParameterSetName='Update',Mandatory = $false)][Switch]$update,
		[Parameter(ParameterSetName='BackupEnd',Mandatory = $false)][Switch]$backupEnd,
		
		[Parameter(ParameterSetName='New',Mandatory = $true)]
		[Parameter(ParameterSetName='Update',Mandatory = $true)]
		[Parameter(ParameterSetName='BackupEnd',Mandatory = $true)]
		[Parameter(Mandatory = $false)][Hashtable]$data,
		
		[Parameter(ParameterSetName='Verify',Mandatory = $true)]
		[Parameter(ParameterSetName='Update',Mandatory = $false)]
		[Parameter(ParameterSetName='BackupEnd',Mandatory = $false)]
		[Parameter(Mandatory = $false)][String]$hwNumber,
		
		[Parameter(ParameterSetName='Verify',Mandatory = $false)]
		[Parameter(ParameterSetName='Update',Mandatory = $true)]
		[Parameter(ParameterSetName='BackupEnd',Mandatory = $true)]
		[Parameter(Mandatory = $false)][String]$sysID
	)
	Write-Verbose "Service-Now Query"
	
	$statusBar.Text = "Conntacting Service-Now"
	
	[bool]$script:lastWebRequest = $false	# Keep track of status of last web request
	
	function ConvertHashTo-Xml() {
		param(
			[Parameter(Mandatory = $true)]
			[Hashtable]$hash,
			[Parameter(Mandatory = $true)]
			[String]$head,
			[Parameter(Mandatory = $true)]
			[String]$foot
		)
		
		$xml = $head
		foreach($_ in $hash.Keys){
			$xml += "<" + $_ + ">" + $hash.$_ + "</" + $_ + ">"
		}
		$xml += $foot
		return $xml
	}
	
	function build-PayloadHash() {
		# Dynamically build payload hash based on input.
		$payloadHash = @{}
		if($data.caller4x4){$payloadHash.add("opened_by",$data.caller4x4)}
		if($data.user4x4){$payloadHash.add("u_customer",$data.user4x4)}
		if($data.userPhone){$payloadHash.add("u_preferred_phone",$data.userPhone)}
		if($data.model){$payloadHash.add("u_computer_model",$data.model)}
		if($data.chassis){$payloadHash.add("u_computer_type",$data.chassis)}
		if($data.username){$payloadHash.add("u_login",$data.username)}
		if($data.password){$payloadHash.add("u_password",$data.password)}
		if($data.assignmentgroup){$payloadHash.add("assignment_group",$data.assignmentgroup)}
		if($data.serial){$payloadHash.add("u_service_tag",$data.serial)}
		if($data.description){$payloadHash.add("short_description",$data.description)}
		if($data.date){$payloadHash.add("u_dobu",$data.date)}	#yyyy-MM-dd eg get-date -Format "yyyy-MM-dd"
		if($data.backuplocation){$payloadHash.add("u_buloc",$data.backuplocation)}
		if($data.completedby){$payloadHash.add("u_bucompby",$data.completedby)}	
		if($data.backupState){$payloadHash.add("u_data_backup",$dataBackupState_translation.($data.backupState))}
		if($data.warranty){$payloadHash.add("u_warranty",$warranty_translation.($data.warranty))}
		if($data.make){$payloadHash.add("u_computer_make",$make_translation.($data.make))}
		if($data.os){$payloadHash.add("u_operationg_system",$os_translation.($data.os))}	# That's not a typo, it's wrong in SN
		if($data.partnership){$payloadHash.add("u_service_level_agreement",$partnership_translation.($data.partnership))}
		# These fields don't exist in SN yet, adding them to comments field instead
		#if($data.pickup){$payloadHash.add("",$data.pickup)}
		#if($data.dropoff){$payloadHash.add("",$data.dropoff)}
		#if($data.callerPhone){$payloadHash.add("",$data.callerPhone)}
		if($data.callerPhone -or $data.pickup -or $data.dropoff){	# If any are not empty
			if(!$data.comments){$data.add("comments","")} 	# Check if comments field exists
			if($data.callerPhone){
				$data.comments += "`nCaller phone number: " + $data.callerPhone
			}
			if($data.pickup){
				$data.comments += "`nPickup Location: " + $data.pickup
			}
			if($data.dropoff){
				$data.comments += "`nDropoff Location: " + $data.dropoff
			}
		}
		if($data.comments){$payloadHash.add("work_notes",$data.comments)}
		
		# Convert to deliverable format
		[string]$payloadXML = ConvertHashTo-XML -hash $payloadHash -head '<?xml version="1.0"?><request><entry>' -foot '</entry></request>' 
		Write-Debug "Payload XML: $payloadXML"
		return $payloadXML
	}
	
	$payloadXML = build-PayloadHash

	try{
		$ErrorActionPreference = "Stop"
	#--Verify--#
		if($verify) {
			Write-Verbose "SN Query Verify"
			$B_Right.Enabled = $false
			$B_Left.Enabled = $false
			
			# Verify existance of case by getting case info
			$Uri = "https://" + $script:StaticServiceNowPath 	# Root URL
			$Uri += "/api/now/table/u_hwcheckin?" 				# Access u_hwcheckin table api
			$Uri += "sysparm_display_value=true"				# Tells api to return display (readable) values of data
			$Uri += "&sysparm_fields=number,sys_id"				# Fields to return
			$Uri += "&sysparm_query=number=" + $hwNumber		# HW Number to look up
			Write-Debug "Query Uri: $Uri"
						
			try {
				make-WebRequest -Uri $Uri -Method Get 
				
				[System.Xml.XmlDocument]$xml = $script:lastWebRequestResult
				$result = $xml.response.result
				
				Write-Verbose $("hwNumber: " + $result.number)
				if($hwNumber -match $result.number){
					# HW ticket exists
					Write-Verbose "HW ticket exists"
					$script:SysID = $result.sys_id
					Write-Verbose "SysID: $script:SysID"
					$script:lastWebRequest = $true
				} else {
					Write-Error "Returned HW number doesn't match requested HW number. How'd this happen?"
					$script:lastWebRequest = $false
				}
			} catch {
				if($error[0].Exception -contains "404"){
					popupError "Invalid Input" "No hardware ticket was found with that number."
				}
				$script:lastWebRequest = $false
			} finally {
				$B_Right.Enabled = $true
				$B_Left.Enabled = $true
			}
	#--Backup--#
		} elseif($backupEnd) {
			Write-Verbose "SN Query Backup End"
			# Update case with backup start/end time
			$Uri = "https://" + $script:StaticServiceNowPath 	# Root URL
			$Uri += "/api/now/table/u_hwcheckin/" 				# Access u_hwcheckin tables
			$Uri += $script:SysID								# Access table
			
			Write-Debug "Query Uri: $Uri"
			
			make-WebRequest -Uri $Uri -Method Put -Payload $payloadXML
			$script:lastWebRequest = $true
	#--Update--#
		} elseif($update) {
			#update the case with provided information
			Write-Verbose "SN Query Update"
			$B_Right.Enabled = $false
			$B_Left.Enabled = $false
			
			$Uri = "https://" + $script:StaticServiceNowPath	# Root URL
			$Uri += "/api/now/table/u_hwcheckin/" 				# Access u_hwcheckin tables
			$Uri += $script:SysID								# Access table
			$Uri += "?sysparm_display_value=true"				# Tell api to return display (readable) values of data
			$Uri += "&sysparm_fields=u_customer,opened_by" 		# Fields to return
			Write-Debug "Query Uri: $Uri"
			
			make-WebRequest -Uri $Uri -Method Put -Payload $payloadXML
			
			$script:lastWebRequest = $true
			$B_Right.Enabled = $true
			$B_Left.Enabled = $true
	#--New--#
		} elseif($new) {
			#create a new case
			Write-Verbose "SN Query New"
			$B_Right.Enabled = $false
			$B_Left.Enabled = $false

			$Uri = "https://" + $script:StaticServiceNowPath 	# Root URL
			$Uri += "/api/now/table/u_hwcheckin?" 				# Access u_hwcheckin table api
			$Uri += "sysparm_display_value=true"				# Tells api to return display (readable) values of data
			$Uri += "&sysparm_fields=number,u_customer,opened_by,sys_id" # Fields to return
			Write-Debug "Query Uri: $Uri"
			
			Write-Debug "PayloadXML: $payloadXML"
			
			make-WebRequest -Uri $Uri -Method Post -Payload $payloadXML  
			
			[System.Xml.XmlDocument]$xml = $script:lastWebRequestResult
			$result = $xml.response.result
			$script:hwNumber = $result.number
			$script:SysID = $result.sys_id
			Write-Verbose "SysID: $script:SysID"
			
			$script:NewCaseCreated = $true
			$script:lastWebRequest = $true
			$B_Right.Enabled = $true
			$B_Left.Enabled = $true
		}
	} catch {
		if($error[0].Exception -contains "400"){
			$msg = "The request URI does not match the APIs in the system, or the operation failed for unknown reasons. Invalid headers can also cause this error."
			$msg += "`nThis is an unrecoverable error. Check ServiceNow for changes and please report this error to Services Team."
			popupError "400: Bad Request" $msg
			Write-Error "400: Bad Request"
			$SCBS.Close()
		}
		if($error[0].Exception -contains "401"){
			$msg = "The ServiceNow request was rejected due to bad credentials. The user is not authorized to use the API."
			$msg += "`nThis is an unrecoverable error. Please report this error to Services Team."
			popupError "401: Unauthorized" $msg
			Write-Error "401: Unauthorized"
			$SCBS.Close()
		}
		if($error[0].Exception -contains "403"){
			$msg = "A bad web request was made to Service-Now. The requested operation is not permitted for the user. This error can also be caused by ACL failures, or business rule or data policy constraints."
			$msg += "`nThis is an unrecoverable error. Check ServiceNow for changes and please report this error to Services Team."
			popupError "403: Forbidden" $msg
			Write-Error "403: Forbidden"
			$SCBS.Close()
		}
		if($error[0].Exception -contains "404"){
			$msg = "The requested resource was not found. This can be caused by an ACL constraint or if the resource does not exist."
			$msg += "`nThis is an unrecoverable error. Check ServiceNow for changes and please report this error to Services Team."
			popupError "404: Not found" $msg
			Write-Error "404: Not found"
			$SCBS.Close()
		}
		if($error[0].Exception -contains "405"){
			$msg = "The HTTP action is not allowed for the requested REST API, or it is not supported by any API."
			$msg += "`nThis is an unrecoverable error. Check ServiceNow for changes and please report this error to Services Team."
			popupError "405: Method not allowed" $msg
			Write-Error "405: Method not allowed"
			$SCBS.Close()
		}
		if($error[0].Exception -contains "408"){
			$msg = "Comunication to ServiceNow timed out."
			$msg += "`nPlease address network settings on this computer and try your request again. If error persists please report this error to the Services Team."
			popupError "408: Request Time-Out" $msg
			Write-Error "408: Request Time-Out"
		}
		if($error[0].Exception -contains "500"){
			$msg = "A bad web request was made to Service-Now."
			$msg += "`nThis is an unrecoverable error. Check ServiceNow for changes and please report this error to Services Team."
			popupError "500: Server Error" $msg
			Write-Error "500: Server Error"
			$SCBS.Close()
		}
		if($error[0].Exception -contains "502"){
			$msg = "A network error was encountered while making a web request to ServiceNow."
			$msg += "`nPlease address network settings on this computer and try your request again. If error persists please report this error to the Services Team."
			popupError "502: Bad Gateway" $msg
			Write-Error "502: Bad Gateway"
		}
		if($error[0].Exception -contains "504"){
			$msg = "A network error was encountered while making a web request to ServiceNow."
			$msg += "`nPlease address network settings on this computer and try your request again. If error persists please report this error to the Services Team."
			popupError "504: Gateway Time-Out" $msg
			Write-Error "504: Gateway Time-Out"
		}
		
		$unhandledErrors = ("302,303,305,402,405,406,407,409,410,411,412,413,414,415,501,503,504,505").split(",")
		foreach($_ in $unhandledErrors){
			if($error[0].Exception -contains $_){
				$msg = "An unhandled network error was encounted while making a web requst to ServiceNow."
				$msg = "`nThis error may be unrecoverable."
				$msg += "`nPlease address network settings on this computer and try your request again. If error persists please report this error to the Services Team."
				popupError ($_ + ": Unhandled Error") $msg
				Write-Error ($_ + ": Unhandled Error")
			}
		}
		
		$script:lastWebRequest = $false
	} finally {
		$ErrorActionPreference = "Continue"
	}
}

function make-WebRequest() {
	param(
		[Parameter(Mandatory = $true)]
		[String]$Uri,		
		[Parameter(Mandatory = $true)]
		[ValidateSet("Post", "Get", "Put", "Delete", "Head", "Options", "Trace", "Merge", "Patch")]
		[String]$Method,
		[Parameter(Mandatory = $false)]
		[String]$Payload,
		[Parameter(Mandatory = $false)]
		[Hashtable]$Headers,
		[Parameter(Mandatory = $false)]
		[ValidateSet("application/json", "application/xml", "text/xml")]
		[String]$ContentType = "text/xml",
		[Parameter(Mandatory = $false)]
		[ValidateSet("application/json", "application/xml", "text/xml")]
		[String]$Accept = "text/xml",
		[Parameter(Mandatory = $false)]
		[String]$UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows CE)",
		[Parameter(Mandatory = $false)]
		[Int]$Timeout = '50000'
	)
	Write-Verbose "SN API Call"
	
	if($PSBoundParameters.ContainsKey('Headers')){
		Write-Verbose "Headers Supplied"
		if(!$Headers.Authorization){
			Write-Verbose "No Authorization Header Supplied: Adding default Authorization"
			$Headers.add("Authorization", "Basic xxxxx") # Redacted
		}
	} else {
		Write-Verbose "No Headers Supplied: Adding default Authorization"
		$Headers = @{
			"Authorization" = "Basic xxxxx"; # Redacted
		}
	}
	
	function webr() {
		# Function configures and makes web requests
		Write-Verbose "Initiate WebRequest"
		$webRequest = [System.Net.HttpWebRequest]::Create($Uri)
		Write-Verbose "Configure WebRequest"
		$webRequest.Method = $Method
		$webRequest.ContentType = $ContentType
		$webRequest.Accept = $Accept
		$webRequest.Timeout = $Timeout
		$webRequest.ReadWriteTimeout = '50000'
		$webRequest.UserAgent = $UserAgent
		#$webRequest.ContentLength = ''
		
		Write-Verbose "Setting Headers"
		foreach($_ in $Headers.Keys){
			$webRequest.Headers.add($_, $Headers.$_)
		}
		
		if(("Post","Put","Merge","Patch") -contains $Method -and ![String]::IsNullOrEmpty($Payload)){
			Write-Verbose "Adding Payload"
			#Here if we need to add a payload to the request
			$Body = [Byte[]][Char[]]$Payload
			$stream = $webRequest.GetRequestStream()
			$stream.Write($Body, 0, $Body.Length)
			$stream.Flush()
			$stream.Close()
		}
		
		Write-Verbose "Getting Response Stream"
		$readStream = New-Object System.IO.StreamReader $webRequest.GetResponse().GetResponseStream()
		$script:lastWebRequestResult = $readStream.ReadToEnd()
		Write-Verbose ("Last Web Request Result: " + $script:lastWebRequestResult)
		
		#Cleanup
		Write-Verbose "Cleanup"
		$readStream.Dispose()
		$readStream.Close()
		$webRequest.GetResponse().Close()
		$webRequest = $null
	}
	
	$timeoutCount = '0'
	function webrCtrl() {
		# Function controls program flow of web request to handle and retry timed-out requests. Which happen more often then it should.
		try{
			$ErrorActionPreference = "Stop"
			if($timeoutCount -lt 5){
				webr
			} else {
				$answer = popupError "Operation Timeout" "Unable to contact Service-Now." 5
				if($answer -eq '4'){
					$timeoutCount = '0'
					webrCtrl
				} else {
					$SCBS.Close()
				}
			}
		} catch {
			if($Error[0].Exception -match "The operation has timed out"){
				Write-Debug "Timeout Caught"
				$timeoutCount++
				webrCtrl
			}
		} finally {
			$ErrorActionPreference = "Continue"
		}
	}
	
	webrCtrl	
	Write-Verbose "SN Query Done"
	return
}

#endregion Service-Now Query Functions

#region Data Functions

function getChassis() {
	# Returns chassis type (form factor) of the computer.
	$chassis = gwmi -Class Win32_SystemEnclosure | Select-Object chassistypes
	# Desktop, Laptop, Tablet, Other
	switch($chassis.chassistypes[0]){
		"3" {return "Desktop"}
		"4" {return "Desktop"}
		"5" {return "Other"}
		"6" {return "Desktop"}
		"7" {return "Desktop"}
		"8" {return "Laptop"}
		"9" {return "Laptop"}
		"10" {return "Laptop"}
		"11" {return "Tablet"}
		"12" {return "Other"}
		"13" {return "Other"}
		"14" {return "Laptop"}
		"15" {return "Other"}
		"16" {return "Other"}
		"17" {return "Other"}
		"18" {return "Other"}
		"19" {return "Other"}
		"20" {return "Other"}
		"21" {return "Other"}
		"22" {return "Other"}
		"23" {return "Other"}
		"24" {return "Other"}
		default {return "Other"}
	}
}

function getOS() {
	# Gets a simplified name of windows version. Verboseness is limited by SN field input options.
	$version = gwmi -Class Win32_OperatingSystem | Select-Object Version -ExpandProperty Version
	
	switch -regex ($version){
		"10.*" {return "Windows 10"}
		"6.3.*" {return "Windows 8"}
		"6.2.*" {return "Windows 8"}
		"6.1.*" {return "Windows 7"}
		"6.0.*" {return "Windows Vista"}
		"5.2.*" {return "Windows XP"}
		"5.1.*" {return "Windows XP"}
		default {return "Windows XP"}
	}	
}

#endregion Data Functions

try {
	#--Main Run Sequence--#
	$VerbosePreference = "Continue"
	$DebugPreference = "Continue"
	
	$gh_pd = (Get-Host).PrivateData
	$gh_pd.VerboseForegroundColor = "Green"
	$gh_pd.DebugForegroundColor = "Yellow"
	$gh_pd.WarningForegroundColor = "Red"
	$gh_pd.ErrorForegroundColor = "White"
	$gh_pd.ErrorBackgroundColor = "Red"
	Write-Verbose "Start"
	
	Write-Debug ("WorkingDir: " + $(Split-Path $(Resolve-Path $myInvocation.MyCommand.Path)))
	
	#region Hashtables 
	# Property and Translation Hashtables
	Write-Verbose "Define Property Hashtables"
	
	$jobQueueProperties = @{
		# Methods for the jobQueue. No properteis actually exists, methods recorded here.
		# JobList = @()
		# RunCount()
		# RunNext()
		# UpdateGUI()
	}
	
	$jobObjectProperties = @{
		# Properties and methods for jobObjects.
		"Name" = "";
		"Path" = "";
		"Job" = "";
		# Backup()
		# State()
	}
	
	$warranty_translation = @{
		# Translation matrix for CoB menu items to ServiceNow menu items.
		# SCBS Value = ServiceNow Value
		"Dell Warranty" = "Dell Warranty";
		"Faculty/Staff"= "Faculty/Staff";
		"HP Warranty" = "HP Warrany";
		"Lenovo Warranty" = "Lenovo Warranty";
		"Triage Services" = "Triage Services";
		"Out of Warranty" = "None";
		"None (Bootcamp, VM, etc)" = "none";
	}
	
	$make_translation = @{
		# Translation matrix for CoB menu items to ServiceNow menu items.
		# SCBS Value = ServiceNow Value
		"Apple" = "Apple";
		"Dell" = "Dell";
		"HP" = "HP";
		"Lenovo" = "Lenovo";
		"Other" = "Other";
	}
	
	$os_translation = @{
		# Translation matrix for CoB menu items to ServiceNow menu items.
		# SCBS Value = ServiceNow Value
		"Windows 7" = "Windows 7";
		"Windows 8" = "Windows 8";
		"Windows Vista" = "Windows Vista";		
		"Windows XP" = "Windows XP";
		"Windows 10" = "Windows 8"; # Change when Operating System field in SN gets a Win10 option.
	}
	
	<# Dev translations
	$partnership_translation = @{
		# Translation matrix for CoB menu items to ServiceNow menu items
		# SCBS Value = ServiceNow Value
		"Administration Faculty/Staff" = "Administration_Faculty/Staff";
		"Athletics (Faculty/Staff)" = "Athletics_FSOnly";
		"College of Architecture (All)" = "COA";
		"College of Earth and Energy (Faculty/Staff)" = "MCEE";
		"College of Engineering (All)" = "COE";
		"College of Fine Arts (Faculty/Staff)" = "Fine_Arts_FSOnly";
		"College of Intl. Studies (Faculty/Staff)" = "International_FSOnly";
		"Enrollment and Student Financial (Faculty/Staff)" = "Enrollment_FSOnly";
		"Graduate College (All)" = "Research_Grad_College_FSOnly";
		"Honors College (Faculty/Staff)" = "Honors_FSOnly";
		"Information Technology (All)" = "IT";
		"Student Affairs (Faculty/Staff)" = "Student_Affairs_FSOnly";
		"University College (Faculty/Staff)" = "University_FSOnly";
		"No Partnership" = "no_sla";
	}
	#>
	
	# Prod translations
	$partnership_translation = @{
		# Translation matrix for CoB_Partnership menu items to ServiceNow menu items
		# Hashtable is also used to dynamically generate CoB_Parternship menu items
		# SCBS Value = ServiceNow Value
		"Administration Faculty/Staff" = "Administration_Faculty/Staff";
		"Athletics (Faculty/Staff)" = "Athletics_FSOnly";
		"College of Architecture (All)" = "COA";
		"College of Earth and Energy (Faculty/Staff)" = "MCEE";
		"College of Engineering (All)" = "COE";
		"College of Fine Arts (Faculty/Staff)" = "Fine_Arts_FSOnly";
		"College of Intl. Studies (Faculty/Staff)" = "International_FSOnly";
		"Enrollment and Student Financial (Faculty/Staff)" = "Enrollment_FSOnly";
		"Graduate College (All)" = "Research_Grad_College_FSOnly";
		"Honors College (Faculty/Staff)" = "Honors_FSOnly";
		"Information Technology (All)" = "IT";
		"Student Affairs (Faculty/Staff)" = "Student_Affairs_FSOnly";
		"University College (Faculty/Staff)" = "University_FSOnly";
		"No Partnership" = "no_sla";
	}
	
	$dataBackupState_translation = @{
		# Translation matrix for CoB menu items to ServiceNow menu items
		# SCBS Value = ServiceNow Value
		"Awaiting Backup" = "Awaiting Backup";
		"Backup Completed" = "Backup Completed";
		"Not Required" = "Not Required";
	}
	
	#endregion Hashtables 

	#region Script Variables
	$script:assignmentGroup			# Assignment Group in ServiceNow to place the Incident. Currently set to Learning Spaces  
	$script:BackupState				# Set to true when the verfiyBackup function has verified all backups are in place.
	$script:CheckedNodes			# Arraylist of nodes that have been selected in the treeview file picker
	$script:DriveLetter				# Drive letter that the script maps norfile too.
	$script:ErrorList				# Tracks errors that have been displayed so they don't popup more than once.
	$script:hwNumber				# Hardware number of Incident. Either provided or set when case is created.
	$script:jobQueue				# Holds jobQueue object that controls the backup process.
	$script:lastWebRequest			# Contains response from last web request.
	$script:lastWebRequestResult	# Bool. Contains status of last web request.
	$script:mapNorfileJob			# Contains mapNorfileJob object that contains methods for mapping Norfile in the background
	$script:netPath					# Variable is used to build and then contain the path to the backup directory in Norfile. eg E:\Backups\HW12345
	$script:NewCaseCreated			# Variable used to track when a new case has been created and update should be used in snQuery.
	$script:NoReturn				# Tracks when user has progressed to far into program to change start type. (After seeing Select, can't return to splash)
	$script:NorfileFailOver			# Tracks whether to use the $script:NorfileIP
	$script:NorFileIP				# Contains IP address of Norfile, used when DNS lookup is not working.
	$script:procCount				# Contains the number of logical processors the computer has. Used in calculating how many simultaneous backup jobs to run. 
	$script:progArr					# Contains an array of charaters used in the progress animation.
	$script:progArrCount			# Contains a number that tracks which character to display next.
	$script:ProgramState			# Tracks what state the program is. Should only be changed with setGuiState function.
	$script:regex_4x4 				# 4x4 validation.
	$script:regex_hw 				# Hardware number validation.
	$script:regex_inputSanitizer 	# Comments input sanitization.
	$script:regex_phoneLocal		# 325 Number validation.
	$script:regex_phoneLong 		# 10 digit number validation. (Must have area code(
	$script:ServiceNowIP			# Contains IP address of Service-Now, used when DNS lookup is not working.
	$script:StartState				# Tracks what state the program started in. Used to differentiate StartNew and StartExist
	$script:StaticNetPath			# Static in the sense that this variable is always correct. Used to build $script:netPath. Overwritten by $script:NorFileIP if DNS fails.
	$script:StaticServiceNowPath	# Static in the sense that this variable is always correct. Overwritten by $script:ServiceNowIP if DNS fails.
	$script:SysID					# ID of the Incident in SN. Used for API calls.
	$script:tickcount				# Tracks how many times the backup timer has ticked.
	$script:validData				# Bool. Tracks if the last call of validateData returned true.
	$script:ValidVersion			# Bool. Tracks if the current version of the program is valid.
	$script:VersionNo				# Contians the current version of the program.
	$script:WorkingDir				# Contians the working directory the script is currently running in. Used to save msinfo32.exe /report locally before copying to Norfile.
	#endregion Script Variables
	
	#region Variable Assignments
	Write-Verbose "Assign Script Variables"
	$script:assignmentGroup = "0982c3d755ad110022f2a8864208f420" 
	$script:ErrorList = @() 		
	$script:hwNumber = ""
	$script:NewCaseCreated = $false 								
	$script:NorFileIP = # Redacted
	#$script:progArr = '|','/','-','\'								
	$script:progArr = '.','..','...',''								
	$script:regex_4x4 = '[a-zA-Z]{1,4}[0-9]{4}'						
	$script:regex_hw = 'hw[0-9]{5}(-[0-9]{1})?'						
	$script:regex_inputSanitizer = '[a-zA-Z0-9,.()\/: _-]'			
	$script:regex_phoneLocal = '5[0-9]{4}'
	$script:regex_phoneLong = '[0-9]{10}'
	$script:ServiceNowIP = # Redacted
	$script:StaticNetPath = # Redacted			
	#$script:StaticServiceNowPath = # Redacted
	$script:StaticServiceNowPath = # Redacted						
	$script:SysID = ""				
	$script:VersionNo = "151015"
	$script:WorkingDir = $(Split-Path $(Resolve-Path $myInvocation.MyCommand.Path))	
	#endregion Variable Assignments

	Write-Verbose "Generate Form"
	GenerateForm
	Write-Verbose "End"
	#--Main Run Sequence End--#
} catch {
	popupError "! Warning !" "Something bad happened; Verify the integrity of the backup."
	Write-Error "The script has encountered an unhandled exception."
	Write-Error $Error[0]
	Exit
}
